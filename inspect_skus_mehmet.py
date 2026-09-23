import pandas as pd
import unicodedata

def norm(s):
    if not isinstance(s, str): return ''
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn').upper().strip()

df_bud = pd.read_excel('temp_budget_input.xlsx', sheet_name='Previsión Matriz Horizontal')
df_bud = df_bud[df_bud['Comercial'] == 'Mehmet']

def clean_budget_sku(row):
    sku = str(row['Código Artículo']).strip().upper()
    pais = str(row['País']).strip().upper()
    if pais != 'ESPAÑA' and len(sku) > 2:
        return sku[:-2]
    return sku

df_bud['SKU_CLEAN'] = df_bud.apply(clean_budget_sku, axis=1)

doga = df_bud[df_bud['Cliente'].str.contains('DOGA', case=False, na=False)]
citar = df_bud[df_bud['Cliente'].str.contains('CITAR|ÇITAR', case=False, na=False)]

print("=== DOGA TARIM SKUs in Sep-26 ===")
print(doga[['Código Artículo', 'SKU_CLEAN', 'Descripción Artículo', 'Sep-26 (u)']])

print("\n=== ÇITAR ÇIÇEK SKUs in Sep-26 ===")
print(citar[['Código Artículo', 'SKU_CLEAN', 'Descripción Artículo', 'Sep-26 (u)']])

df_ped = pd.read_excel('temp_pedidos_input.xlsx')
ped_citar = df_ped[df_ped['Cliente'].str.contains('CITAR|ÇITAR', case=False, na=False)]
print("\n=== ÇITAR ÇIÇEK Pedidos ===")
print(ped_citar[['CodigoArticulo', 'Articulo', 'Pedidas', 'Servidas', 'Pendientes']])
