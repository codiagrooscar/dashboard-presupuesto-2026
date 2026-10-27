import pandas as pd
import unicodedata

def norm(s):
    if not isinstance(s, str): return ''
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn').upper().strip()

df_ped = pd.read_excel('Pedidos 23.09 budget.xlsx')
print('=== PEDIDOS INFO ===')
print('Min date:', df_ped['FechaPedido'].min())
print('Max date:', df_ped['FechaPedido'].max())
print('Total rows:', len(df_ped))
print('Unique clients in Pedidos:', df_ped['Cliente'].nunique())
print('Total Pedidas (u):', df_ped['Pedidas'].sum())
print('Total Servidas (u):', df_ped['Servidas'].sum())
print('Total Pendientes (u):', df_ped['Pendientes'].sum())
print('Total ImporteNeto (€):', df_ped['ImporteNeto'].sum())

# Load budget
df_bud = pd.read_excel('Presupuesto_Ventas_2027.xlsx', sheet_name='Previsión Matriz Horizontal')
df_bud = df_bud[df_bud['Comercial'].notna() & (~df_bud['Comercial'].astype(str).str.contains('TOTAL', case=False))]

df_ped['CLIENTE_NORM'] = df_ped['Cliente'].apply(norm)
df_bud['CLIENTE_NORM'] = df_bud['Cliente'].apply(norm)

# Also check other sources for client to commercial mapping if any
client_to_comercial = {}
for _, r in df_bud.iterrows():
    c_n = r['CLIENTE_NORM']
    if c_n not in client_to_comercial:
        client_to_comercial[c_n] = r['Comercial']

ped_clients = df_ped['CLIENTE_NORM'].unique()
matched = [c for c in ped_clients if c in client_to_comercial]
unmatched = [c for c in ped_clients if c not in client_to_comercial]

print(f'\nClients match: {len(matched)} / {len(ped_clients)} matched directly')
if unmatched:
    print('Unmatched clients in Pedidos:')
    for u in unmatched:
        orig = df_ped[df_ped['CLIENTE_NORM']==u]['Cliente'].iloc[0]
        cod = df_ped[df_ped['CLIENTE_NORM']==u]['CodigoCliente'].iloc[0]
        rows_count = len(df_ped[df_ped['CLIENTE_NORM']==u])
        uds = df_ped[df_ped['CLIENTE_NORM']==u]['Pedidas'].sum()
        print(f'  - Code {cod}: "{orig}" (Rows: {rows_count}, Uds: {uds})')
