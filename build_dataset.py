import json
import os
import pandas as pd
import numpy as np
import unicodedata
from datetime import datetime

def norm(s):
    if not isinstance(s, str): return ''
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn').upper().strip()

def build_dataset():
    import shutil
    shutil.copy2('Pedidos 23.09 budget.xlsx', 'temp_pedidos_input.xlsx')
    df_ped = pd.read_excel('temp_pedidos_input.xlsx')
    def map_client_name(c):
        if not isinstance(c, str): return c
        cn = norm(c)
        if 'DOGA TARIM' in cn:
            return 'ÇITAR ÇIÇEK TARIM TIC. LTD. STI'
        return c

    df_ped['Cliente'] = df_ped['Cliente'].apply(map_client_name)
    df_ped['CLIENTE_NORM'] = df_ped['Cliente'].apply(norm)
    df_ped['SKU_CLEAN'] = df_ped['CodigoArticulo'].astype(str).str.strip().str.upper()
    df_ped = df_ped[~df_ped['CLIENTE_NORM'].str.contains('SUSTAINABLE', case=False, na=False)].copy()

    shutil.copy2('Presupuesto_Ventas_2027.xlsx', 'temp_budget_input.xlsx')
    df_bud = pd.read_excel('temp_budget_input.xlsx', sheet_name='Previsión Matriz Horizontal')
    df_bud = df_bud[df_bud['Comercial'].notna() & (~df_bud['Comercial'].astype(str).str.contains('TOTAL', case=False))].copy()
    df_bud['Cliente'] = df_bud['Cliente'].apply(map_client_name)
    df_bud['CLIENTE_NORM'] = df_bud['Cliente'].apply(norm)

    def clean_budget_sku(row):
        sku = str(row['Código Artículo']).strip().upper()
        pais = str(row['País']).strip().upper()
        if pais != 'ESPAÑA' and len(sku) > 2:
            return sku[:-2]
        return sku

    df_bud['SKU_CLEAN'] = df_bud.apply(clean_budget_sku, axis=1)

    client_map = df_bud.groupby('CLIENTE_NORM')['Comercial'].first().to_dict()
    client_map[norm('FINCA DOÑA ANA C.B.')] = 'Javier'
    client_map[norm('FITOSANITARIOS CARCAIXENT, S.L.')] = 'Javier'
    client_map[norm('ALMENDRALIA IBÉRICA, S.L.U.')] = 'Javier'

    df_ped['Comercial'] = df_ped['CLIENTE_NORM'].map(client_map).fillna('Sin Asignar')

    comerciales = ['Alfonso', 'García', 'Irene', 'Javier', 'Mehmet', 'Pedro', 'Ricardo']

    summary_list = []
    comerciales_data = {}

    for c in comerciales:
        sub_ped = df_ped[df_ped['Comercial'] == c]
        sub_bud = df_bud[df_bud['Comercial'] == c]

        b_tot = float(sub_bud['Sep-26 (u)'].sum())
        p_tot = float(sub_ped['Pedidas'].sum())
        s_tot = float(sub_ped['Servidas'].sum())
        pend_tot = float(sub_ped['Pendientes'].sum())
        imp_tot = float(sub_ped['ImporteNeto'].sum())
        cli_cnt = int(sub_ped['Cliente'].nunique())
        gap_tot = float(max(0.0, b_tot - p_tot))
        desv_tot = float(p_tot - b_tot)
        pct_tot = float((p_tot / b_tot * 100) if b_tot > 0 else (100.0 if p_tot > 0 else 0.0))

        summary_list.append({
            'comercial': c,
            'clientes_activos': cli_cnt,
            'budget_uds': b_tot,
            'pedidos_uds': p_tot,
            'servidas_uds': s_tot,
            'pendientes_uds': pend_tot,
            'gap_uds': gap_tot,
            'desv_uds': desv_tot,
            'pct_consec': pct_tot,
            'importe_neto': imp_tot,
            'estado': 'Superado' if pct_tot >= 100 else ('En Curso' if pct_tot >= 50 else 'Rezagado')
        })

        # Detalle de líneas
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

            b_u = float(b_match['Sep-26 (u)'].sum() if len(b_match) > 0 else 0.0)
            p_u = float(p_match['Pedidas'].sum() if len(p_match) > 0 else 0.0)
            s_u = float(p_match['Servidas'].sum() if len(p_match) > 0 else 0.0)
            pend_u = float(p_match['Pendientes'].sum() if len(p_match) > 0 else 0.0)
            imp = float(p_match['ImporteNeto'].sum() if len(p_match) > 0 else 0.0)

            # Excluir líneas que tienen todo 0 en Septiembre (sin budget y sin pedidos)
            if b_u == 0 and p_u == 0:
                continue

            gap_u = float(max(0.0, b_u - p_u))
            desv_u = float(p_u - b_u)
            pct_u = float((p_u / b_u * 100) if b_u > 0 else (100.0 if p_u > 0 else 0.0))

            if b_u == 0 and p_u > 0:
                est = "Extra Budget"
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
                'pct_consec': pct_u,
                'importe_neto': imp,
                'estado': est
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
                'pct_consec': pct_tot,
                'importe_neto': imp_tot
            },
            'lineas': lines
        }

    global_kpis = {
        'budget_uds': float(df_bud['Sep-26 (u)'].sum()),
        'pedidos_uds': float(df_ped['Pedidas'].sum()),
        'servidas_uds': float(df_ped['Servidas'].sum()),
        'pendientes_uds': float(df_ped['Pendientes'].sum()),
        'gap_uds': float(sum(item['gap_uds'] for item in summary_list)),
        'pct_consec': float((df_ped['Pedidas'].sum() / df_bud['Sep-26 (u)'].sum() * 100)),
        'importe_neto': float(df_ped['ImporteNeto'].sum()),
        'fecha_corte': '22/09/2026',
        'fecha_generacion': datetime.now().strftime('%d/%m/%Y %H:%M')
    }

    dataset = {
        'global_kpis': global_kpis,
        'resumen_comerciales': summary_list,
        'comerciales': comerciales_data
    }
    return dataset

if __name__ == '__main__':
    d = build_dataset()
    with open('dashboard_data.json', 'w', encoding='utf-8') as f:
        json.dump(d, f, ensure_ascii=False, indent=2)
    os.makedirs('web_dashboard', exist_ok=True)
    with open('web_dashboard/dashboard_data.json', 'w', encoding='utf-8') as f:
        json.dump(d, f, ensure_ascii=False, indent=2)
    print("dashboard_data.json generado correctamente en raíz y web_dashboard/. Resumen global:")
    print(d['global_kpis'])
