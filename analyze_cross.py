import pandas as pd
import unicodedata

def norm(s):
    if not isinstance(s, str): return ''
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn').upper().strip()

df_ped = pd.read_excel('Pedidos 23.09 budget.xlsx')
df_bud = pd.read_excel('Presupuesto_Ventas_2027.xlsx', sheet_name='Previsión Matriz Horizontal')
df_bud = df_bud[df_bud['Comercial'].notna() & (~df_bud['Comercial'].astype(str).str.contains('TOTAL', case=False))]

df_ped['CLIENTE_NORM'] = df_ped['Cliente'].apply(norm)
df_bud['CLIENTE_NORM'] = df_bud['Cliente'].apply(norm)

# Client -> Comercial in budget
client_com = df_bud.groupby('CLIENTE_NORM')['Comercial'].first().to_dict()
client_pais = df_bud.groupby('CLIENTE_NORM')['País'].first().to_dict()

df_ped['Comercial_Budget'] = df_ped['CLIENTE_NORM'].map(client_com)
df_ped['Pais_Budget'] = df_ped['CLIENTE_NORM'].map(client_pais)

print("=== PEDIDOS POR COMERCIAL ASIGNADO (SEGÚN CLIENTE EN BUDGET) ===")
ped_by_com = df_ped.groupby('Comercial_Budget', dropna=False).agg(
    Lineas=('Pedido', 'count'),
    Clientes=('Cliente', 'nunique'),
    Uds_Pedidas=('Pedidas', 'sum'),
    Uds_Servidas=('Servidas', 'sum'),
    Uds_Pendientes=('Pendientes', 'sum'),
    Importe=('ImporteNeto', 'sum')
)
print(ped_by_com)

# Now check SKUs
# In budget, remove 2-letter country code for non-Spain
def clean_budget_sku(row):
    sku = str(row['Código Artículo']).strip().upper()
    pais = str(row['País']).strip().upper()
    if pais != 'ESPAÑA' and len(sku) > 2:
        return sku[:-2]
    return sku

df_bud['SKU_CLEAN'] = df_bud.apply(clean_budget_sku, axis=1)
df_ped['SKU_CLEAN'] = df_ped['CodigoArticulo'].astype(str).str.strip().str.upper()

# Check SKU overlap
budget_skus = set(df_bud['SKU_CLEAN'].unique())
ped_skus = set(df_ped['SKU_CLEAN'].unique())

print(f"\nTotal SKUs en Pedidos: {len(ped_skus)}")
print(f"Total SKUs en Budget (normalizados): {len(budget_skus)}")
print(f"SKUs de Pedidos que existen en Budget: {len(ped_skus.intersection(budget_skus))} / {len(ped_skus)}")
unmatched_skus = ped_skus - budget_skus
if unmatched_skus:
    print(f"SKUs de Pedidos no encontrados en Budget ({len(unmatched_skus)}):")
    for s in unmatched_skus:
        sample_ped = df_ped[df_ped['SKU_CLEAN']==s]
        print(f"  - {s} ({sample_ped['Articulo'].iloc[0]}) | Pedidas: {sample_ped['Pedidas'].sum()} | Clientes: {sample_ped['Cliente'].unique()[:2]}")
