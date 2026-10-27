import pandas as pd
import numpy as np
import unicodedata

def norm(s):
    if not isinstance(s, str): return ''
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn').upper().strip()

# 1. Load Pedidos
df_ped = pd.read_excel('Pedidos 23.09 budget.xlsx')
df_ped['CLIENTE_NORM'] = df_ped['Cliente'].apply(norm)
df_ped['SKU_CLEAN'] = df_ped['CodigoArticulo'].astype(str).str.strip().str.upper()

# Exclude Sustainable Agro Solutions
df_ped = df_ped[~df_ped['CLIENTE_NORM'].str.contains('SUSTAINABLE', case=False, na=False)].copy()

# 2. Load Budget
df_bud = pd.read_excel('Presupuesto_Ventas_2027.xlsx', sheet_name='Previsión Matriz Horizontal')
df_bud = df_bud[df_bud['Comercial'].notna() & (~df_bud['Comercial'].astype(str).str.contains('TOTAL', case=False))].copy()
df_bud['CLIENTE_NORM'] = df_bud['Cliente'].apply(norm)

def clean_budget_sku(row):
    sku = str(row['Código Artículo']).strip().upper()
    pais = str(row['País']).strip().upper()
    if pais != 'ESPAÑA' and len(sku) > 2:
        return sku[:-2]
    return sku

df_bud['SKU_CLEAN'] = df_bud.apply(clean_budget_sku, axis=1)

# Client to comercial mapping
client_map = df_bud.groupby('CLIENTE_NORM')['Comercial'].first().to_dict()

# Explicit user assignments
client_map[norm('FINCA DOÑA ANA C.B.')] = 'Javier'
client_map[norm('FITOSANITARIOS CARCAIXENT, S.L.')] = 'Javier'
client_map[norm('ALMENDRALIA IBÉRICA, S.L.U.')] = 'Javier'

df_ped['Comercial'] = df_ped['CLIENTE_NORM'].map(client_map).fillna('Sin Asignar')

print("=== REVISED TOTALS BY COMERCIAL ===")
com_summary = []
comerciales = ['Alfonso', 'García', 'Irene', 'Javier', 'Mehmet', 'Pedro', 'Ricardo']

for c in comerciales:
    sub_ped = df_ped[df_ped['Comercial'] == c]
    sub_bud = df_bud[df_bud['Comercial'] == c]
    
    bud_uds = sub_bud['Sep-26 (u)'].sum()
    ped_uds = sub_ped['Pedidas'].sum()
    serv_uds = sub_ped['Servidas'].sum()
    pend_uds = sub_ped['Pendientes'].sum()
    imp_neto = sub_ped['ImporteNeto'].sum()
    
    gap_uds = max(0.0, bud_uds - ped_uds)
    pct_consec = (ped_uds / bud_uds * 100) if bud_uds > 0 else (100.0 if ped_uds > 0 else 0.0)
    desv_uds = ped_uds - bud_uds
    
    com_summary.append({
        'Comercial': c,
        'Budget Sep (u)': round(bud_uds, 2),
        'Pedidos Sep (u)': round(ped_uds, 2),
        'Servidas (u)': round(serv_uds, 2),
        'Pendientes (u)': round(pend_uds, 2),
        'Falta Conseguir (u)': round(gap_uds, 2),
        'Desviación (u)': round(desv_uds, 2),
        '% Consecución': round(pct_consec, 1),
        'Importe Neto (€)': round(imp_neto, 2)
    })

df_res = pd.DataFrame(com_summary)
print(df_res.to_string(index=False))

total_bud = df_res['Budget Sep (u)'].sum()
total_ped = df_res['Pedidos Sep (u)'].sum()
total_gap = df_res['Falta Conseguir (u)'].sum()
total_pct = (total_ped / total_bud * 100) if total_bud > 0 else 0.0
total_imp = df_res['Importe Neto (€)'].sum()

print("\n--- TOTAL CONSOLIDADO ---")
print(f"Budget Sep: {total_bud:,.2f} u")
print(f"Pedidos Sep: {total_ped:,.2f} u")
print(f"Falta Conseguir (Gap): {total_gap:,.2f} u")
print(f"% Consecución Global: {total_pct:.1f}%")
print(f"Importe Neto Total: {total_imp:,.2f} €")
