import json
import os
import shutil
import re
import pandas as pd
import numpy as np
import unicodedata
from datetime import datetime

def norm(s):
    if not isinstance(s, str): return ''
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn').upper().strip()

def map_client_name(c):
    if not isinstance(c, str): return c
    cn = norm(c)
    if 'DOGA TARIM' in cn:
        return 'ÇİTAR ÇİÇEK TARIM TIC. LTD. STI'
    return c

def clean_budget_sku(row):
    sku = str(row['Código Artículo']).strip().upper()
    pais = str(row['País']).strip().upper()
    if pais != 'ESPAÑA' and len(sku) > 2:
        return sku[:-2]
    return sku

def get_latest_pedidos_file():
    candidates = [
        f for f in os.listdir('.')
        if f.startswith('Pedidos') and f.endswith('.xlsx') and not f.startswith(('temp_', '~$'))
    ]
    if not candidates:
        return 'Pedidos 23.09 budget.xlsx'
    def sort_key(f):
        m = re.search(r'(\d{1,2})\.(\d{1,2})', f)
        if m:
            return (int(m.group(2)), int(m.group(1)), os.path.getmtime(f))
        return (0, 0, os.path.getmtime(f))
    candidates.sort(key=sort_key, reverse=True)
    return candidates[0]

def load_data_sources():
    pedidos_file = get_latest_pedidos_file()
    shutil.copy2(pedidos_file, 'temp_pedidos_input.xlsx')
    df_ped = pd.read_excel('temp_pedidos_input.xlsx')
    df_ped['Cliente'] = df_ped['Cliente'].apply(map_client_name)
    df_ped['CLIENTE_NORM'] = df_ped['Cliente'].apply(norm)
    df_ped['SKU_CLEAN'] = df_ped['CodigoArticulo'].astype(str).str.strip().str.upper()
    df_ped = df_ped[~df_ped['CLIENTE_NORM'].str.contains('SUSTAINABLE', case=False, na=False)].copy()

    src_budget = 'Presupuesto_Ventas_2027_Original_Sin_Aplanar.xlsx' if os.path.exists('Presupuesto_Ventas_2027_Original_Sin_Aplanar.xlsx') else 'Presupuesto_Ventas_2027.xlsx'
    shutil.copy2(src_budget, 'temp_budget_input.xlsx')
    df_bud = pd.read_excel('temp_budget_input.xlsx', sheet_name='Previsión Matriz Horizontal')
    df_bud = df_bud[df_bud['Comercial'].notna() & (~df_bud['Comercial'].astype(str).str.contains('TOTAL', case=False))].copy()

    FACTOR = 0.9479961392983893
    INV_FACTOR = 1.0 / FACTOR
    for idx, row in df_bud.iterrows():
        com = str(row['Comercial']).strip()
        if 'garc' not in norm(com).lower():
            for m_col in ['Sep-26 (u)', 'Oct-26 (u)']:
                v = row[m_col]
                if pd.notna(v) and v > 0 and abs(v - round(v)) > 0.04:
                    restored = round(v * INV_FACTOR, 2)
                    if abs(restored - round(restored)) < 0.05:
                        restored = float(round(restored))
                    df_bud.at[idx, m_col] = restored

    df_bud['Cliente'] = df_bud['Cliente'].apply(map_client_name)
    df_bud['CLIENTE_NORM'] = df_bud['Cliente'].apply(norm)
    df_bud['SKU_CLEAN'] = df_bud.apply(clean_budget_sku, axis=1)

    client_map = df_bud.groupby('CLIENTE_NORM')['Comercial'].first().to_dict()
    client_map[norm('FINCA DOÑA ANA C.B.')] = 'Javier'
    client_map[norm('FITOSANITARIOS CARCAIXENT, S.L.')] = 'Javier'
    client_map[norm('ALMENDRALIA IBÉRICA, S.L.U.')] = 'Javier'
    client_map[norm('TÉCNICAS AGRÍCOLAS, S.A.')] = 'Ricardo'

    df_ped['Comercial'] = df_ped['CLIENTE_NORM'].map(client_map).fillna('Sin Asignar')
    return df_ped, df_bud, pedidos_file

def build_month_dataset(df_ped, df_bud, month_col, period_id, period_name, period_subtitle, period_status, badge_class, is_active=False):
    comerciales = ['Alfonso', 'García', 'Irene', 'Javier', 'Mehmet', 'Pedro', 'Ricardo']
    summary_list = []
    comerciales_data = {}

    for c in comerciales:
        sub_ped = df_ped[df_ped['Comercial'] == c]
        sub_bud = df_bud[df_bud['Comercial'] == c]

        b_tot = float(sub_bud[month_col].sum())
        p_tot = float(sub_ped['Pedidas'].sum())
        s_tot = float(sub_ped['Servidas'].sum())
        pend_tot = float(sub_ped['Pendientes'].sum())
        imp_tot = float(sub_ped['ImporteNeto'].sum())
        cli_cnt = int(sub_ped['Cliente'].nunique())
        if cli_cnt == 0 and len(sub_bud) > 0:
            cli_cnt = int(sub_bud['Cliente'].nunique())

        gap_tot = float(max(0.0, b_tot - p_tot))
        desv_tot = float(p_tot - b_tot)
        pct_tot = float((p_tot / b_tot * 100) if b_tot > 0 else (100.0 if p_tot > 0 else 0.0))

        keys_set = set()
        for _, r in sub_bud.iterrows():
            keys_set.add((r['Cliente'], r['SKU_CLEAN']))
        for _, r in sub_ped.iterrows():
            keys_set.add((r['Cliente'], r['SKU_CLEAN']))

        keys_sorted = sorted(list(keys_set), key=lambda x: (x[0], x[1]))
        lines = []

        for cli, sku in keys_sorted:
            b_match = sub_bud[(sub_bud['Cliente'] == cli) & (sub_bud['SKU_CLEAN'] == sku)]
            p_match = sub_ped[(sub_ped['Cliente'] == cli) & (sub_ped['SKU_CLEAN'] == sku)]

            desc = ""
            if len(b_match) > 0 and pd.notna(b_match['Descripción Artículo'].iloc[0]):
                desc = str(b_match['Descripción Artículo'].iloc[0])
            elif len(p_match) > 0 and pd.notna(p_match['Articulo'].iloc[0]):
                desc = str(p_match['Articulo'].iloc[0])

            b_u = float(b_match[month_col].sum() if len(b_match) > 0 else 0.0)
            p_u = float(p_match['Pedidas'].sum() if len(p_match) > 0 else 0.0)
            s_u = float(p_match['Servidas'].sum() if len(p_match) > 0 else 0.0)
            pend_u = float(p_match['Pendientes'].sum() if len(p_match) > 0 else 0.0)
            imp = float(p_match['ImporteNeto'].sum() if len(p_match) > 0 else 0.0)

            if b_u == 0 and p_u == 0:
                continue

            gap_u = float(max(0.0, b_u - p_u))
            desv_u = float(p_u - b_u)
            pct_u = float((p_u / b_u * 100) if b_u > 0 else (100.0 if p_u > 0 else 0.0))
            
            err_abs_u = float(abs(p_u - b_u))
            acc_u = float((1.0 - err_abs_u / b_u) * 100 if b_u > 0 else (100.0 if p_u == 0 else 0.0))

            if b_u == 0 and p_u > 0:
                est = "Extra Estimación"
            elif pct_u >= 100:
                est = "Superado"
            elif pct_u >= 50:
                est = "En Curso"
            elif p_u > 0:
                est = "Rezagado"
            else:
                est = "Sin Pedido"

            lines.append({
                'cliente': cli,
                'sku': sku,
                'descripcion': desc,
                'budget_uds': b_u,
                'pedidos_uds': p_u,
                'servidas_uds': s_u,
                'pendientes_uds': pend_u,
                'gap_uds': gap_u,
                'desv_uds': desv_u,
                'error_abs_uds': err_abs_u,
                'pct_consec': pct_u,
                'forecast_accuracy': acc_u,
                'importe_neto': imp,
                'estado': est
            })

        com_err_tot = float(sum(l['error_abs_uds'] for l in lines))
        com_acc_tot = float((1.0 - com_err_tot / b_tot) * 100 if b_tot > 0 else 0.0)

        summary_list.append({
            'comercial': c,
            'clientes_activos': cli_cnt,
            'budget_uds': b_tot,
            'pedidos_uds': p_tot,
            'servidas_uds': s_tot,
            'pendientes_uds': pend_tot,
            'gap_uds': gap_tot,
            'desv_uds': desv_tot,
            'error_abs_uds': com_err_tot,
            'pct_consec': pct_tot,
            'forecast_accuracy': com_acc_tot,
            'importe_neto': imp_tot,
            'estado': 'Superado' if pct_tot >= 100 else ('En Curso' if pct_tot >= 50 else 'Rezagado')
        })

        comerciales_data[c] = {
            'comercial': c,
            'kpis': {
                'budget_uds': b_tot,
                'pedidos_uds': p_tot,
                'servidas_uds': s_tot,
                'pendientes_uds': pend_tot,
                'gap_uds': gap_tot,
                'desv_uds': desv_tot,
                'error_abs_uds': com_err_tot,
                'pct_consec': pct_tot,
                'forecast_accuracy': com_acc_tot,
                'importe_neto': imp_tot
            },
            'lineas': lines
        }

    tot_budget_all = float(df_bud[month_col].sum())
    tot_pedidos_all = float(df_ped['Pedidas'].sum())
    tot_error_all = float(sum(item['error_abs_uds'] for item in summary_list))
    global_accuracy = float((1.0 - tot_error_all / tot_budget_all) * 100 if tot_budget_all > 0 else 0.0)

    global_kpis = {
        'budget_uds': tot_budget_all,
        'pedidos_uds': tot_pedidos_all,
        'servidas_uds': float(df_ped['Servidas'].sum()),
        'pendientes_uds': float(df_ped['Pendientes'].sum()),
        'gap_uds': float(sum(item['gap_uds'] for item in summary_list)),
        'desv_uds': float(tot_pedidos_all - tot_budget_all),
        'error_abs_uds': tot_error_all,
        'pct_consec': float((tot_pedidos_all / tot_budget_all * 100) if tot_budget_all > 0 else 0.0),
        'forecast_accuracy': global_accuracy,
        'importe_neto': float(df_ped['ImporteNeto'].sum()),
        'fecha_corte': '01/10/2026',
        'fecha_generacion': datetime.now().strftime('%d/%m/%Y %H:%M')
    }

    return {
        'id': period_id,
        'nombre': period_name,
        'subtitulo': period_subtitle,
        'estado_periodo': period_status,
        'badge_class': badge_class,
        'is_active': is_active,
        'global_kpis': global_kpis,
        'resumen_comerciales': summary_list,
        'comerciales': comerciales_data
    }

def build_accumulated_dataset(d_sep, d_oct):
    comerciales = ['Alfonso', 'García', 'Irene', 'Javier', 'Mehmet', 'Pedro', 'Ricardo']
    summary_list = []
    comerciales_data = {}

    for c in comerciales:
        c_sep = d_sep['comerciales'][c]
        c_oct = d_oct['comerciales'][c]

        b_tot = c_sep['kpis']['budget_uds'] + c_oct['kpis']['budget_uds']
        p_tot = c_sep['kpis']['pedidos_uds'] + c_oct['kpis']['pedidos_uds']
        s_tot = c_sep['kpis']['servidas_uds'] + c_oct['kpis']['servidas_uds']
        pend_tot = c_oct['kpis']['pendientes_uds'] if c_oct['kpis']['pendientes_uds'] > 0 else c_sep['kpis']['pendientes_uds']
        imp_tot = c_sep['kpis']['importe_neto'] + c_oct['kpis']['importe_neto']
        cli_cnt = max(c_sep['kpis'].get('clientes_activos', 0), c_oct['kpis'].get('clientes_activos', 0))

        # Combinar líneas por (cliente, sku)
        lines_dict = {}
        for l in c_sep['lineas']:
            k = (l['cliente'], l['sku'])
            lines_dict[k] = {
                'cliente': l['cliente'],
                'sku': l['sku'],
                'descripcion': l['descripcion'],
                'budget_uds': l['budget_uds'],
                'pedidos_uds': l['pedidos_uds'],
                'servidas_uds': l['servidas_uds'],
                'pendientes_uds': l['pendientes_uds'],
                'importe_neto': l['importe_neto']
            }
        for l in c_oct['lineas']:
            k = (l['cliente'], l['sku'])
            if k in lines_dict:
                lines_dict[k]['budget_uds'] += l['budget_uds']
                lines_dict[k]['pedidos_uds'] += l['pedidos_uds']
                lines_dict[k]['servidas_uds'] += l['servidas_uds']
                lines_dict[k]['pendientes_uds'] = max(lines_dict[k]['pendientes_uds'], l['pendientes_uds'])
                lines_dict[k]['importe_neto'] += l['importe_neto']
                if not lines_dict[k]['descripcion'] and l['descripcion']:
                    lines_dict[k]['descripcion'] = l['descripcion']
            else:
                lines_dict[k] = {
                    'cliente': l['cliente'],
                    'sku': l['sku'],
                    'descripcion': l['descripcion'],
                    'budget_uds': l['budget_uds'],
                    'pedidos_uds': l['pedidos_uds'],
                    'servidas_uds': l['servidas_uds'],
                    'pendientes_uds': l['pendientes_uds'],
                    'importe_neto': l['importe_neto']
                }

        lines = []
        for k, v in sorted(lines_dict.items(), key=lambda x: (x[0][0], x[0][1])):
            b_u = v['budget_uds']
            p_u = v['pedidos_uds']
            if b_u == 0 and p_u == 0:
                continue
            gap_u = max(0.0, b_u - p_u)
            desv_u = p_u - b_u
            pct_u = (p_u / b_u * 100) if b_u > 0 else (100.0 if p_u > 0 else 0.0)
            err_abs_u = abs(p_u - b_u)
            acc_u = (1.0 - err_abs_u / b_u) * 100 if b_u > 0 else (100.0 if p_u == 0 else 0.0)

            if b_u == 0 and p_u > 0:
                est = "Extra Estimación"
            elif pct_u >= 100:
                est = "Superado"
            elif pct_u >= 50:
                est = "En Curso"
            elif p_u > 0:
                est = "Rezagado"
            else:
                est = "Sin Pedido"

            v.update({
                'gap_uds': gap_u,
                'desv_uds': desv_u,
                'error_abs_uds': err_abs_u,
                'pct_consec': pct_u,
                'forecast_accuracy': acc_u,
                'estado': est
            })
            lines.append(v)

        com_err_tot = float(sum(l['error_abs_uds'] for l in lines))
        com_acc_tot = float((1.0 - com_err_tot / b_tot) * 100 if b_tot > 0 else 0.0)
        gap_tot = float(sum(l['gap_uds'] for l in lines))
        desv_tot = float(p_tot - b_tot)
        pct_tot = float((p_tot / b_tot * 100) if b_tot > 0 else 0.0)

        summary_list.append({
            'comercial': c,
            'clientes_activos': cli_cnt,
            'budget_uds': b_tot,
            'pedidos_uds': p_tot,
            'servidas_uds': s_tot,
            'pendientes_uds': pend_tot,
            'gap_uds': gap_tot,
            'desv_uds': desv_tot,
            'error_abs_uds': com_err_tot,
            'pct_consec': pct_tot,
            'forecast_accuracy': com_acc_tot,
            'importe_neto': imp_tot,
            'estado': 'Superado' if pct_tot >= 100 else ('En Curso' if pct_tot >= 50 else 'Rezagado')
        })

        comerciales_data[c] = {
            'comercial': c,
            'kpis': {
                'budget_uds': b_tot,
                'pedidos_uds': p_tot,
                'servidas_uds': s_tot,
                'pendientes_uds': pend_tot,
                'gap_uds': gap_tot,
                'desv_uds': desv_tot,
                'error_abs_uds': com_err_tot,
                'pct_consec': pct_tot,
                'forecast_accuracy': com_acc_tot,
                'importe_neto': imp_tot
            },
            'lineas': lines
        }

    tot_budget_all = d_sep['global_kpis']['budget_uds'] + d_oct['global_kpis']['budget_uds']
    tot_pedidos_all = d_sep['global_kpis']['pedidos_uds'] + d_oct['global_kpis']['pedidos_uds']
    tot_servidas_all = d_sep['global_kpis']['servidas_uds'] + d_oct['global_kpis']['servidas_uds']
    tot_pendientes_all = d_oct['global_kpis']['pendientes_uds'] if d_oct['global_kpis']['pendientes_uds'] > 0 else d_sep['global_kpis']['pendientes_uds']
    tot_error_all = float(sum(item['error_abs_uds'] for item in summary_list))
    tot_gap_all = float(sum(item['gap_uds'] for item in summary_list))
    global_accuracy = float((1.0 - tot_error_all / tot_budget_all) * 100 if tot_budget_all > 0 else 0.0)

    global_kpis = {
        'budget_uds': tot_budget_all,
        'pedidos_uds': tot_pedidos_all,
        'servidas_uds': tot_servidas_all,
        'pendientes_uds': tot_pendientes_all,
        'gap_uds': tot_gap_all,
        'desv_uds': float(tot_pedidos_all - tot_budget_all),
        'error_abs_uds': tot_error_all,
        'pct_consec': float((tot_pedidos_all / tot_budget_all * 100) if tot_budget_all > 0 else 0.0),
        'forecast_accuracy': global_accuracy,
        'importe_neto': d_sep['global_kpis']['importe_neto'] + d_oct['global_kpis']['importe_neto'],
        'fecha_corte': '01/10/2026',
        'fecha_generacion': datetime.now().strftime('%d/%m/%Y %H:%M')
    }

    return {
        'id': 'acumulado',
        'nombre': 'Acumulado Campaña (Sep + Oct)',
        'subtitulo': 'Consolidado histórico bimestral 2026',
        'estado_periodo': 'Consolidado',
        'badge_class': 'badge-info',
        'is_active': False,
        'global_kpis': global_kpis,
        'resumen_comerciales': summary_list,
        'comerciales': comerciales_data
    }

def build_all():
    df_ped_sep, df_bud, pedidos_file = load_data_sources()

    # Período Septiembre 2026 (Cerrado con pedidos hasta fin de mes)
    d_sep = build_month_dataset(
        df_ped=df_ped_sep,
        df_bud=df_bud,
        month_col='Sep-26 (u)',
        period_id='2026-09',
        period_name='Septiembre 2026',
        period_subtitle='Mes Cerrado Oficialmente a 30/09/2026',
        period_status='Cerrado',
        badge_class='badge-neutral',
        is_active=False
    )

    # Período Octubre 2026 (Primer día de mes / previsión inicial + pedidos día 1)
    # Por ahora en pedidos de octubre iniciales tomamos pedidos con fecha octubre si existen o vacíos
    df_ped_oct = df_ped_sep.copy()
    if 'FechaPedido' in df_ped_oct.columns:
        df_ped_oct = df_ped_oct[pd.to_datetime(df_ped_oct['FechaPedido']).dt.month == 10].copy()
    else:
        df_ped_oct = df_ped_oct.iloc[0:0].copy()

    # Si aún no han entrado pedidos de octubre, generamos dataset de octubre con la previsión base de Octubre
    d_oct = build_month_dataset(
        df_ped=df_ped_oct,
        df_bud=df_bud,
        month_col='Oct-26 (u)',
        period_id='2026-10',
        period_name='Octubre 2026',
        period_subtitle='Previsión Base Oficial (Día 1 de Octubre)',
        period_status='En Curso',
        badge_class='badge-success',
        is_active=True
    )

    # Período Acumulado Campaña (Sep + Oct)
    d_acum = build_accumulated_dataset(d_sep, d_oct)

    periods_data = {
        'default_period': '2026-10',
        'periods': {
            '2026-10': d_oct,
            '2026-09': d_sep,
            'acumulado': d_acum
        },
        'period_list': [
            {'id': '2026-10', 'nombre': 'Octubre 2026', 'badge': 'EN CURSO', 'badge_class': 'badge-success'},
            {'id': '2026-09', 'nombre': 'Septiembre 2026', 'badge': 'CERRADO', 'badge_class': 'badge-neutral'},
            {'id': 'acumulado', 'nombre': 'Acumulado (Sep + Oct)', 'badge': 'CONSOLIDADO', 'badge_class': 'badge-info'}
        ]
    }

    # Guardar en json y js
    with open('web_dashboard/dashboard_data.json', 'w', encoding='utf-8') as f:
        json.dump(periods_data, f, ensure_ascii=False, indent=2)

    with open('web_dashboard/dashboard_data.js', 'w', encoding='utf-8') as f:
        f.write("window.DASHBOARD_PERIODS_DATA = " + json.dumps(periods_data, ensure_ascii=False) + ";\n")
        # Mantener window.DASHBOARD_DATA apuntando por defecto al mes activo para retrocompatibilidad
        f.write("window.DASHBOARD_DATA = window.DASHBOARD_PERIODS_DATA.periods['2026-10'];\n")

    print("Multi-period dataset generado con éxito:")
    print("- Sep budget:", d_sep['global_kpis']['budget_uds'], "pedidos:", d_sep['global_kpis']['pedidos_uds'])
    print("- Oct budget:", d_oct['global_kpis']['budget_uds'], "pedidos:", d_oct['global_kpis']['pedidos_uds'])
    print("- Acumulado budget:", d_acum['global_kpis']['budget_uds'], "pedidos:", d_acum['global_kpis']['pedidos_uds'])

if __name__ == '__main__':
    build_all()
