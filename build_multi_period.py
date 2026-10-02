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

def safe_read_excel(filepath, sheet_name=0):
    try:
        return pd.read_excel(filepath, sheet_name=sheet_name)
    except Exception:
        import subprocess
        base_name = os.path.basename(filepath)
        temp_file = f"temp_safe_{base_name}"
        subprocess.run(['powershell', '-Command', f'Copy-Item "{filepath}" "{temp_file}" -Force'], check=True)
        return pd.read_excel(temp_file, sheet_name=sheet_name)

def get_latest_october_file():
    candidates = [
        f for f in os.listdir('.')
        if f.startswith('Pedidos') and f.endswith('.xlsx') and not f.startswith(('temp_', '~$'))
    ]
    # October files have day 02 or higher of month 10
    oct_candidates = []
    for f in candidates:
        m = re.search(r'(\d{1,2})\.(\d{1,2})', f)
        if m:
            day, month = int(m.group(1)), int(m.group(2))
            if month == 10 and day >= 2:
                oct_candidates.append((day, month, f))
    if oct_candidates:
        oct_candidates.sort(key=lambda x: (x[1], x[0]), reverse=True)
        return oct_candidates[0][2]
    return 'Pedidos 02.10 comerciales.xlsx' if os.path.exists('Pedidos 02.10 comerciales.xlsx') else None

def load_data_sources():
    # 1. Budget
    src_budget = 'Presupuesto_Ventas_2027_Original_Sin_Aplanar.xlsx' if os.path.exists('Presupuesto_Ventas_2027_Original_Sin_Aplanar.xlsx') else 'Presupuesto_Ventas_2027.xlsx'
    df_bud = safe_read_excel(src_budget, sheet_name='Previsión Matriz Horizontal')
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

    # 2. Pedidos Septiembre (Cierre oficial Pedidos 01.10 comerciales.xlsx)
    sep_file = 'Pedidos 01.10 comerciales.xlsx' if os.path.exists('Pedidos 01.10 comerciales.xlsx') else 'Pedidos 30.09 comerciales.xlsx'
    df_ped_sep = safe_read_excel(sep_file)
    df_ped_sep['Cliente'] = df_ped_sep['Cliente'].apply(map_client_name)
    df_ped_sep['CLIENTE_NORM'] = df_ped_sep['Cliente'].apply(norm)
    df_ped_sep['SKU_CLEAN'] = df_ped_sep['CodigoArticulo'].astype(str).str.strip().str.upper()
    df_ped_sep = df_ped_sep[~df_ped_sep['CLIENTE_NORM'].str.contains('SUSTAINABLE', case=False, na=False)].copy()
    if 'FechaPedido' in df_ped_sep.columns:
        df_ped_sep = df_ped_sep[pd.to_datetime(df_ped_sep['FechaPedido']).dt.month == 9].copy()
    df_ped_sep['Comercial'] = df_ped_sep['CLIENTE_NORM'].map(client_map).fillna('Sin Asignar')

    # 3. Pedidos Octubre (Pedidos 02.10 comerciales.xlsx y posteriores)
    oct_file = get_latest_october_file()
    if oct_file and os.path.exists(oct_file):
        df_ped_oct = safe_read_excel(oct_file)
        df_ped_oct['Cliente'] = df_ped_oct['Cliente'].apply(map_client_name)
        df_ped_oct['CLIENTE_NORM'] = df_ped_oct['Cliente'].apply(norm)
        df_ped_oct['SKU_CLEAN'] = df_ped_oct['CodigoArticulo'].astype(str).str.strip().str.upper()
        df_ped_oct = df_ped_oct[~df_ped_oct['CLIENTE_NORM'].str.contains('SUSTAINABLE', case=False, na=False)].copy()
        if 'FechaPedido' in df_ped_oct.columns:
            df_ped_oct = df_ped_oct[pd.to_datetime(df_ped_oct['FechaPedido']).dt.month == 10].copy()
        df_ped_oct['Comercial'] = df_ped_oct['CLIENTE_NORM'].map(client_map).fillna('Sin Asignar')
    else:
        df_ped_oct = df_ped_sep.iloc[0:0].copy()

    return df_ped_sep, df_ped_oct, df_bud, sep_file, oct_file

def is_line_nacional(cli, sku, b_match, com, df_bud):
    if len(b_match) > 0:
        return (str(b_match['País'].iloc[0]).strip().upper() == 'ESPAÑA')
    if com in ['Alfonso', 'Irene', 'Javier', 'Pedro']:
        return True
    if com in ['Mehmet', 'Ricardo']:
        return False
    # Para García o clientes sin budget
    m = df_bud[df_bud['Cliente'] == cli]
    if len(m) > 0:
        return (str(m['País'].iloc[0]).strip().upper() == 'ESPAÑA')
    sku_clean = str(sku).strip().upper()
    if len(sku_clean) >= 2 and sku_clean[-2:].isalpha():
        return False
    return True

def build_month_dataset(df_ped, df_bud, month_col, period_id, period_name, period_subtitle, period_status, badge_class, fecha_corte='01/10/2026', is_active=False):
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
        desv_tot = float(abs(p_tot - b_tot))
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
            desv_u = float(abs(p_u - b_u))
            pct_u = float((p_u / b_u * 100) if b_u > 0 else (100.0 if p_u > 0 else 0.0))
            
            err_abs_u = desv_u
            # Si el % de forecast accuracy es negativo, pon 0%
            acc_u = float(max(0.0, (1.0 - desv_u / b_u) * 100)) if b_u > 0 else (100.0 if p_u == 0 else 0.0)

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

            is_nac = is_line_nacional(cli, sku, b_match, c, df_bud)

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
                'error_abs_uds': desv_u,
                'pct_consec': pct_u,
                'forecast_accuracy': acc_u,
                'importe_neto': imp,
                'estado': est,
                'is_nacional': is_nac,
                'origen': 'Nacional' if is_nac else 'Exportación'
            })

        com_desv_tot = float(sum(l['desv_uds'] for l in lines))
        # Si es negativo, pon 0%
        com_acc_tot = float(max(0.0, (1.0 - com_desv_tot / b_tot) * 100) if b_tot > 0 else 0.0)

        # Métricas Productos Nacionales (sin siglas de país)
        com_bud_nac = float(sum(l['budget_uds'] for l in lines if l['is_nacional']))
        com_ped_nac = float(sum(l['pedidos_uds'] for l in lines if l['is_nacional']))
        com_serv_nac = float(sum(l['servidas_uds'] for l in lines if l['is_nacional']))
        com_pend_nac = float(sum(l['pendientes_uds'] for l in lines if l['is_nacional']))
        com_gap_nac = float(sum(l['gap_uds'] for l in lines if l['is_nacional']))
        com_desv_nac = float(sum(l['desv_uds'] for l in lines if l['is_nacional']))
        com_acc_nac = float(max(0.0, (1.0 - com_desv_nac / com_bud_nac) * 100) if com_bud_nac > 0 else 0.0)
        com_pct_nac = float((com_ped_nac / com_bud_nac * 100) if com_bud_nac > 0 else (100.0 if com_ped_nac > 0 else 0.0))
        tiene_nac = bool(com_bud_nac > 0 or com_ped_nac > 0)

        summary_list.append({
            'comercial': c,
            'clientes_activos': cli_cnt,
            'budget_uds': b_tot,
            'pedidos_uds': p_tot,
            'servidas_uds': s_tot,
            'pendientes_uds': pend_tot,
            'gap_uds': gap_tot,
            'desv_uds': com_desv_tot,
            'error_abs_uds': com_desv_tot,
            'pct_consec': pct_tot,
            'forecast_accuracy': com_acc_tot,
            'budget_uds_nacional': com_bud_nac,
            'pedidos_uds_nacional': com_ped_nac,
            'servidas_uds_nacional': com_serv_nac,
            'pendientes_uds_nacional': com_pend_nac,
            'gap_uds_nacional': com_gap_nac,
            'desv_uds_nacional': com_desv_nac,
            'pct_consec_nacional': com_pct_nac,
            'forecast_accuracy_nacional': com_acc_nac,
            'tiene_nacional': tiene_nac,
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
                'desv_uds': com_desv_tot,
                'error_abs_uds': com_desv_tot,
                'pct_consec': pct_tot,
                'forecast_accuracy': com_acc_tot,
                'budget_uds_nacional': com_bud_nac,
                'pedidos_uds_nacional': com_ped_nac,
                'servidas_uds_nacional': com_serv_nac,
                'pendientes_uds_nacional': com_pend_nac,
                'gap_uds_nacional': com_gap_nac,
                'desv_uds_nacional': com_desv_nac,
                'pct_consec_nacional': com_pct_nac,
                'forecast_accuracy_nacional': com_acc_nac,
                'tiene_nacional': tiene_nac,
                'importe_neto': imp_tot
            },
            'lineas': lines
        }

    tot_budget_all = float(df_bud[month_col].sum())
    tot_pedidos_all = float(df_ped['Pedidas'].sum())
    tot_desv_all = float(sum(item['desv_uds'] for item in summary_list))
    # Si es negativo, pon 0%
    global_accuracy = float(max(0.0, (1.0 - tot_desv_all / tot_budget_all) * 100) if tot_budget_all > 0 else 0.0)

    tot_bud_nac = float(sum(item['budget_uds_nacional'] for item in summary_list))
    tot_ped_nac = float(sum(item['pedidos_uds_nacional'] for item in summary_list))
    tot_serv_nac = float(sum(item['servidas_uds_nacional'] for item in summary_list))
    tot_pend_nac = float(sum(item['pendientes_uds_nacional'] for item in summary_list))
    tot_gap_nac = float(sum(item['gap_uds_nacional'] for item in summary_list))
    tot_desv_nac = float(sum(item['desv_uds_nacional'] for item in summary_list))
    global_accuracy_nac = float(max(0.0, (1.0 - tot_desv_nac / tot_bud_nac) * 100) if tot_bud_nac > 0 else 0.0)
    global_pct_nac = float((tot_ped_nac / tot_bud_nac * 100) if tot_bud_nac > 0 else (100.0 if tot_ped_nac > 0 else 0.0))

    global_kpis = {
        'budget_uds': tot_budget_all,
        'pedidos_uds': tot_pedidos_all,
        'servidas_uds': float(df_ped['Servidas'].sum()),
        'pendientes_uds': float(df_ped['Pendientes'].sum()),
        'gap_uds': float(sum(item['gap_uds'] for item in summary_list)),
        'desv_uds': tot_desv_all,
        'error_abs_uds': tot_desv_all,
        'pct_consec': float((tot_pedidos_all / tot_budget_all * 100) if tot_budget_all > 0 else 0.0),
        'forecast_accuracy': global_accuracy,
        'budget_uds_nacional': tot_bud_nac,
        'pedidos_uds_nacional': tot_ped_nac,
        'servidas_uds_nacional': tot_serv_nac,
        'pendientes_uds_nacional': tot_pend_nac,
        'gap_uds_nacional': tot_gap_nac,
        'desv_uds_nacional': tot_desv_nac,
        'pct_consec_nacional': global_pct_nac,
        'forecast_accuracy_nacional': global_accuracy_nac,
        'importe_neto': float(df_ped['ImporteNeto'].sum()),
        'fecha_corte': fecha_corte,
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

def build_accumulated_dataset(d_sep, d_oct, fecha_corte='02/10/2026'):
    comerciales = ['Alfonso', 'García', 'Irene', 'Javier', 'Mehmet', 'Pedro', 'Ricardo']
    summary_list = []
    comerciales_data = {}

    for c in comerciales:
        c_sep = d_sep['comerciales'][c]
        c_oct = d_oct['comerciales'][c]

        b_tot = c_sep['kpis']['budget_uds'] + c_oct['kpis']['budget_uds']
        p_tot = c_sep['kpis']['pedidos_uds'] + c_oct['kpis']['pedidos_uds']
        s_tot = c_sep['kpis']['servidas_uds'] + c_oct['kpis']['servidas_uds']
        pend_tot = c_oct['kpis']['pendientes_uds'] + c_sep['kpis']['pendientes_uds']
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
                'importe_neto': l['importe_neto'],
                'is_nacional': l.get('is_nacional', True),
                'origen': l.get('origen', 'Nacional')
            }
        for l in c_oct['lineas']:
            k = (l['cliente'], l['sku'])
            if k in lines_dict:
                lines_dict[k]['budget_uds'] += l['budget_uds']
                lines_dict[k]['pedidos_uds'] += l['pedidos_uds']
                lines_dict[k]['servidas_uds'] += l['servidas_uds']
                lines_dict[k]['pendientes_uds'] += l['pendientes_uds']
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
                    'importe_neto': l['importe_neto'],
                    'is_nacional': l.get('is_nacional', True),
                    'origen': l.get('origen', 'Nacional')
                }

        lines = []
        for k, v in sorted(lines_dict.items(), key=lambda x: (x[0][0], x[0][1])):
            b_u = v['budget_uds']
            p_u = v['pedidos_uds']
            if b_u == 0 and p_u == 0:
                continue
            gap_u = max(0.0, b_u - p_u)
            desv_u = abs(p_u - b_u)
            pct_u = (p_u / b_u * 100) if b_u > 0 else (100.0 if p_u > 0 else 0.0)
            err_abs_u = desv_u
            # Si el % de forecast accuracy es negativo, pon 0%
            acc_u = float(max(0.0, (1.0 - desv_u / b_u) * 100)) if b_u > 0 else (100.0 if p_u == 0 else 0.0)

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
                'error_abs_uds': desv_u,
                'pct_consec': pct_u,
                'forecast_accuracy': acc_u,
                'estado': est
            })
            lines.append(v)

        com_desv_tot = float(sum(l['desv_uds'] for l in lines))
        com_acc_tot = float(max(0.0, (1.0 - com_desv_tot / b_tot) * 100) if b_tot > 0 else 0.0)
        gap_tot = float(sum(l['gap_uds'] for l in lines))
        pct_tot = float((p_tot / b_tot * 100) if b_tot > 0 else 0.0)

        # Métricas Nacionales Acumulado
        com_bud_nac = float(sum(l['budget_uds'] for l in lines if l.get('is_nacional', True)))
        com_ped_nac = float(sum(l['pedidos_uds'] for l in lines if l.get('is_nacional', True)))
        com_serv_nac = float(sum(l['servidas_uds'] for l in lines if l.get('is_nacional', True)))
        com_pend_nac = float(sum(l['pendientes_uds'] for l in lines if l.get('is_nacional', True)))
        com_gap_nac = float(sum(l['gap_uds'] for l in lines if l.get('is_nacional', True)))
        com_desv_nac = float(sum(l['desv_uds'] for l in lines if l.get('is_nacional', True)))
        com_acc_nac = float(max(0.0, (1.0 - com_desv_nac / com_bud_nac) * 100) if com_bud_nac > 0 else 0.0)
        com_pct_nac = float((com_ped_nac / com_bud_nac * 100) if com_bud_nac > 0 else 0.0)
        tiene_nac = bool(com_bud_nac > 0 or com_ped_nac > 0)

        summary_list.append({
            'comercial': c,
            'clientes_activos': cli_cnt,
            'budget_uds': b_tot,
            'pedidos_uds': p_tot,
            'servidas_uds': s_tot,
            'pendientes_uds': pend_tot,
            'gap_uds': gap_tot,
            'desv_uds': com_desv_tot,
            'error_abs_uds': com_desv_tot,
            'pct_consec': pct_tot,
            'forecast_accuracy': com_acc_tot,
            'budget_uds_nacional': com_bud_nac,
            'pedidos_uds_nacional': com_ped_nac,
            'servidas_uds_nacional': com_serv_nac,
            'pendientes_uds_nacional': com_pend_nac,
            'gap_uds_nacional': com_gap_nac,
            'desv_uds_nacional': com_desv_nac,
            'pct_consec_nacional': com_pct_nac,
            'forecast_accuracy_nacional': com_acc_nac,
            'tiene_nacional': tiene_nac,
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
                'desv_uds': com_desv_tot,
                'error_abs_uds': com_desv_tot,
                'pct_consec': pct_tot,
                'forecast_accuracy': com_acc_tot,
                'budget_uds_nacional': com_bud_nac,
                'pedidos_uds_nacional': com_ped_nac,
                'servidas_uds_nacional': com_serv_nac,
                'pendientes_uds_nacional': com_pend_nac,
                'gap_uds_nacional': com_gap_nac,
                'desv_uds_nacional': com_desv_nac,
                'pct_consec_nacional': com_pct_nac,
                'forecast_accuracy_nacional': com_acc_nac,
                'tiene_nacional': tiene_nac,
                'importe_neto': imp_tot
            },
            'lineas': lines
        }

    tot_budget_all = d_sep['global_kpis']['budget_uds'] + d_oct['global_kpis']['budget_uds']
    tot_pedidos_all = d_sep['global_kpis']['pedidos_uds'] + d_oct['global_kpis']['pedidos_uds']
    tot_servidas_all = d_sep['global_kpis']['servidas_uds'] + d_oct['global_kpis']['servidas_uds']
    tot_pendientes_all = d_sep['global_kpis']['pendientes_uds'] + d_oct['global_kpis']['pendientes_uds']
    tot_desv_all = float(sum(item['desv_uds'] for item in summary_list))
    tot_gap_all = float(sum(item['gap_uds'] for item in summary_list))
    global_accuracy = float(max(0.0, (1.0 - tot_desv_all / tot_budget_all) * 100) if tot_budget_all > 0 else 0.0)

    tot_bud_nac = float(sum(item['budget_uds_nacional'] for item in summary_list))
    tot_ped_nac = float(sum(item['pedidos_uds_nacional'] for item in summary_list))
    tot_serv_nac = float(sum(item['servidas_uds_nacional'] for item in summary_list))
    tot_pend_nac = float(sum(item['pendientes_uds_nacional'] for item in summary_list))
    tot_gap_nac = float(sum(item['gap_uds_nacional'] for item in summary_list))
    tot_desv_nac = float(sum(item['desv_uds_nacional'] for item in summary_list))
    global_accuracy_nac = float(max(0.0, (1.0 - tot_desv_nac / tot_bud_nac) * 100) if tot_bud_nac > 0 else 0.0)
    global_pct_nac = float((tot_ped_nac / tot_bud_nac * 100) if tot_bud_nac > 0 else 0.0)

    global_kpis = {
        'budget_uds': tot_budget_all,
        'pedidos_uds': tot_pedidos_all,
        'servidas_uds': tot_servidas_all,
        'pendientes_uds': tot_pendientes_all,
        'gap_uds': tot_gap_all,
        'desv_uds': tot_desv_all,
        'error_abs_uds': tot_desv_all,
        'pct_consec': float((tot_pedidos_all / tot_budget_all * 100) if tot_budget_all > 0 else 0.0),
        'forecast_accuracy': global_accuracy,
        'budget_uds_nacional': tot_bud_nac,
        'pedidos_uds_nacional': tot_ped_nac,
        'servidas_uds_nacional': tot_serv_nac,
        'pendientes_uds_nacional': tot_pend_nac,
        'gap_uds_nacional': tot_gap_nac,
        'desv_uds_nacional': tot_desv_nac,
        'pct_consec_nacional': global_pct_nac,
        'forecast_accuracy_nacional': global_accuracy_nac,
        'importe_neto': d_sep['global_kpis']['importe_neto'] + d_oct['global_kpis']['importe_neto'],
        'fecha_corte': fecha_corte,
        'fecha_generacion': datetime.now().strftime('%d/%m/%Y %H:%M')
    }

    return {
        'id': 'acumulado',
        'nombre': 'Acumulado Campaña (Sep + Oct)',
        'subtitulo': f'Consolidado histórico bimestral 2026 (Corte a {fecha_corte})',
        'estado_periodo': 'Consolidado',
        'badge_class': 'badge-info',
        'is_active': False,
        'global_kpis': global_kpis,
        'resumen_comerciales': summary_list,
        'comerciales': comerciales_data
    }

def build_all():
    df_ped_sep, df_ped_oct, df_bud, sep_file, oct_file = load_data_sources()

    # Período Septiembre 2026 (Cerrado oficial al 30/09/2026)
    d_sep = build_month_dataset(
        df_ped=df_ped_sep,
        df_bud=df_bud,
        month_col='Sep-26 (u)',
        period_id='2026-09',
        period_name='Septiembre 2026',
        period_subtitle='Mes Cerrado Oficialmente a 30/09/2026',
        period_status='Cerrado',
        badge_class='badge-neutral',
        fecha_corte='30/09/2026',
        is_active=False
    )

    # Período Octubre 2026 (En curso con pedidos actuales de octubre)
    m_oct = re.search(r'(\d{1,2})\.(\d{1,2})', oct_file) if oct_file else None
    oct_corte = f"{m_oct.group(1).zfill(2)}/{m_oct.group(2).zfill(2)}/2026" if m_oct else "02/10/2026"

    d_oct = build_month_dataset(
        df_ped=df_ped_oct,
        df_bud=df_bud,
        month_col='Oct-26 (u)',
        period_id='2026-10',
        period_name='Octubre 2026',
        period_subtitle=f'En Curso (Corte a {oct_corte})',
        period_status='En Curso',
        badge_class='badge-success',
        fecha_corte=oct_corte,
        is_active=True
    )

    # Período Acumulado Campaña (Sep + Oct)
    d_acum = build_accumulated_dataset(d_sep, d_oct, fecha_corte=oct_corte)

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
        # Mantener window.DASHBOARD_DATA apuntando por defecto a Octubre (período en curso)
        f.write("window.DASHBOARD_DATA = window.DASHBOARD_PERIODS_DATA.periods['2026-10'];\n")

    print("Multi-period dataset generado con éxito:")
    print(f"- Sep (cierre {d_sep['global_kpis']['fecha_corte']}): budget={d_sep['global_kpis']['budget_uds']}, pedidos={d_sep['global_kpis']['pedidos_uds']}, accuracy={d_sep['global_kpis']['forecast_accuracy']:.1f}%")
    print(f"- Oct (corte {d_oct['global_kpis']['fecha_corte']}): budget={d_oct['global_kpis']['budget_uds']}, pedidos={d_oct['global_kpis']['pedidos_uds']}, accuracy={d_oct['global_kpis']['forecast_accuracy']:.1f}%")
    print(f"- Acumulado (corte {d_acum['global_kpis']['fecha_corte']}): budget={d_acum['global_kpis']['budget_uds']}, pedidos={d_acum['global_kpis']['pedidos_uds']}, accuracy={d_acum['global_kpis']['forecast_accuracy']:.1f}%")

if __name__ == '__main__':
    build_all()
