import pandas as pd
import unicodedata

def norm(s):
    if not isinstance(s, str): return ''
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn').upper().strip()

def map_client_name(c):
    if not isinstance(c, str): return c
    cn = norm(c)
    if 'DOGA TARIM' in cn:
        return 'ÇITAR ÇIÇEK TARIM TIC. LTD. STI'
    return c

df_bud = pd.read_excel('temp_budget_input.xlsx', sheet_name='Previsión Matriz Horizontal')
df_bud = df_bud[df_bud['Comercial'] == 'Mehmet'].copy()
df_bud['Cliente'] = df_bud['Cliente'].apply(map_client_name)
df_bud['CLIENTE_NORM'] = df_bud['Cliente'].apply(norm)

def clean_budget_sku(row):
    sku = str(row['Código Artículo']).strip().upper()
    pais = str(row['País']).strip().upper()
    if pais != 'ESPAÑA' and len(sku) > 2:
        return sku[:-2]
    return sku

df_bud['SKU_CLEAN'] = df_bud.apply(clean_budget_sku, axis=1)

df_ped = pd.read_excel('temp_pedidos_input.xlsx')
df_ped['Cliente'] = df_ped['Cliente'].apply(map_client_name)
df_ped['CLIENTE_NORM'] = df_ped['Cliente'].apply(norm)
df_ped['SKU_CLEAN'] = df_ped['CodigoArticulo'].astype(str).str.strip().str.upper()
df_ped_mehmet = df_ped[df_ped['Cliente'].str.contains('CITAR|ÇITAR', case=False, na=False)].copy()

keys_set = set()
for _, r in df_bud.iterrows():
    keys_set.add((r['Cliente'], r['SKU_CLEAN']))
for _, r in df_ped_mehmet.iterrows():
    keys_set.add((r['Cliente'], r['SKU_CLEAN']))

keys_sorted = sorted(list(keys_set), key=lambda x: (x[0], x[1]))

print("=== MEHMET LINES AFTER UNIFICATION (OMITTING ZEROS) ===")
for cli, sku in keys_sorted:
    b_match = df_bud[(df_bud['Cliente'] == cli) & (df_bud['SKU_CLEAN'] == sku)]
    p_match = df_ped_mehmet[(df_ped_mehmet['Cliente'] == cli) & (df_ped_mehmet['SKU_CLEAN'] == sku)]
    
    b_u = float(b_match['Sep-26 (u)'].sum() if len(b_match) > 0 else 0.0)
    p_u = float(p_match['Pedidas'].sum() if len(p_match) > 0 else 0.0)
    s_u = float(p_match['Servidas'].sum() if len(p_match) > 0 else 0.0)
    pend_u = float(p_match['Pendientes'].sum() if len(p_match) > 0 else 0.0)
    
    if b_u == 0 and p_u == 0:
        continue
        
    gap_u = max(0.0, b_u - p_u)
    desv_u = p_u - b_u
    pct_u = (p_u / b_u * 100) if b_u > 0 else 100.0
    
    desc = ""
    if len(b_match) > 0 and pd.notna(b_match['Descripción Artículo'].iloc[0]):
        desc = str(b_match['Descripción Artículo'].iloc[0])
    elif len(p_match) > 0 and pd.notna(p_match['Articulo'].iloc[0]):
        desc = str(p_match['Articulo'].iloc[0])
        
    print(f"{cli} | {sku} ({desc}) | Budget: {b_u:,.2f} | Pedidos: {p_u:,.2f} | Pendientes: {pend_u:,.2f} | Gap: {gap_u:,.2f} | %: {pct_u:.1f}%")
