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

client_map = {}
for _, r in df_bud.iterrows():
    cn = r['CLIENTE_NORM']
    if cn not in client_map:
        client_map[cn] = {
            'comercial': r['Comercial'],
            'pais': r['País'],
            'clasif': r['Clasificación Cliente']
        }

# Group pedidos by client
clients_summary = []
for c_code, group in df_ped.groupby('CodigoCliente'):
    c_name = group['Cliente'].iloc[0]
    cn = norm(c_name)
    budget_info = client_map.get(cn, None)
    com = budget_info['comercial'] if budget_info else 'NO ASIGNADO'
    pais = budget_info['pais'] if budget_info else 'DESCONOCIDO'
    
    # check budget sep for this client
    bud_client = df_bud[df_bud['CLIENTE_NORM'] == cn]
    sep_bud_uds = bud_client['Sep-26 (u)'].sum() if len(bud_client) > 0 else 0.0
    
    clients_summary.append({
        'Codigo': c_code,
        'Cliente': c_name,
        'Comercial': com,
        'País': pais,
        'Líneas Pedido': len(group),
        'Uds Pedidas': group['Pedidas'].sum(),
        'Uds Servidas': group['Servidas'].sum(),
        'Uds Pendientes': group['Pendientes'].sum(),
        'Importe Pedidos (€)': group['ImporteNeto'].sum(),
        'Budget Sep (u)': sep_bud_uds
    })

res = pd.DataFrame(clients_summary)
res = res.sort_values(by=['Comercial', 'Cliente'])
print(res.to_string(index=False))
