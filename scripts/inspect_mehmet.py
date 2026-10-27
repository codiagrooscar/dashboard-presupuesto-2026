import pandas as pd
import unicodedata

def norm(s):
    if not isinstance(s, str): return ''
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn').upper().strip()

df_bud = pd.read_excel('temp_budget_input.xlsx', sheet_name='Previsión Matriz Horizontal')
df_bud = df_bud[df_bud['Comercial'] == 'Mehmet']
print('Mehmet clients in Budget:')
for cli in df_bud['Cliente'].unique():
    sub = df_bud[df_bud['Cliente'] == cli]
    sep_u = sub['Sep-26 (u)'].sum()
    print(f'  - {cli} (Rows: {len(sub)}, Sep-26: {sep_u})')

df_ped = pd.read_excel('temp_pedidos_input.xlsx')
df_ped['CLIENTE_NORM'] = df_ped['Cliente'].apply(norm)
print('\nMehmet clients in Pedidos:')
for cli in df_ped['Cliente'].unique():
    if any(k in norm(cli) for k in ['DOGA', 'CITAR', 'CICEK']):
        sub = df_ped[df_ped['Cliente'] == cli]
        print(f'  - {cli} (Rows: {len(sub)}, Pedidas: {sub["Pedidas"].sum()})')
