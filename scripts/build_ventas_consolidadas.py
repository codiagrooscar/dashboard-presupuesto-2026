import sys
sys.stdout.reconfigure(encoding='utf-8')

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import numpy as np
import unicodedata
from collections import Counter

base_dir = r"c:\Users\oscar.ocampo\OneDrive - Agroquímica Codiagro S.L\Escritorio\Presupuesto 2026\Calsif Comerciales"
os.chdir(base_dir)

def norm(s):
    if not isinstance(s, str): return ''
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn').upper().strip()

print("Cargando archivos de ventas...")

# 1. Ventas Reales Ene-Ago 2026
f_real = "Ventas reales de Eebro a Agosto de 2026.xlsx"
df_r = pd.read_excel(f_real)
df_r = df_r[df_r['CLIENTE'] != 'Total general'].copy()
df_r['CLIENTE_NORM'] = df_r['CLIENTE'].apply(norm)
df_r['SKU_BASE'] = df_r['CodigoArticulo'].astype(str).str.strip().str.upper()
df_r['Unidades'] = pd.to_numeric(df_r['Unidades '], errors='coerce').fillna(0)
df_r['Importe'] = pd.to_numeric(df_r['Importe'], errors='coerce').fillna(0)
df_r['Pr_Medio'] = np.where(df_r['Unidades'] > 0, df_r['Importe'] / df_r['Unidades'], 0.0)
df_r['TIPO_PERIODO'] = 'REAL (Ene-Ago)'

# 2. Ventas Previstas Sep-Dic 2026
f_prev = "Ventas esperadas de sep a dic de 2026.xlsx"
df_p = pd.read_excel(f_prev, skiprows=3)
col_imp = [c for c in df_p.columns if 'IMPORTE' in c][0]
col_sku_b = [c for c in df_p.columns if 'Base' in c][0]
col_sku_p = [c for c in df_p.columns if 'Previsi' in c][0]
col_pais = [c for c in df_p.columns if 'Pa' in c and 's' in c][0]

df_p['CLIENTE_NORM'] = df_p['Cliente'].apply(norm)
df_p['Unidades'] = pd.to_numeric(df_p['TOTAL UNIDADES'], errors='coerce').fillna(0)
df_p['Importe'] = pd.to_numeric(df_p[col_imp], errors='coerce').fillna(0)
df_p['SKU_PREV_ORIGINAL'] = df_p[col_sku_p].astype(str).str.strip().str.upper()

# Normalizar SKU Base para Prev: si no está o es NaN, asignar por descripción
def get_clean_prev_sku(row):
    b = str(row[col_sku_b]).strip().upper()
    if b != '' and b != 'NAN' and b != 'NONE':
        return b
    # Si viene con 2 letras de país al final y país no es España:
    p = str(row[col_sku_p]).strip().upper()
    pais = str(row[col_pais]).strip().upper()
    if pais != 'ESPAÑA' and pais != 'NAN' and len(p) > 2:
        return p[:-2]
    # Casos de Mehmet sin código
    desc = str(row['Descripción Artículo']).upper()
    if 'AMINOG' in desc: return 'AMGF0001'
    if 'AQUOM' in desc: return 'AQUOM0005'
    return p

df_p['SKU_BASE'] = df_p.apply(get_clean_prev_sku, axis=1)
df_p['Pr_Medio'] = np.where(df_p['Unidades'] > 0, df_p['Importe'] / df_p['Unidades'], 0.0)
df_p['TIPO_PERIODO'] = 'PREVISIÓN (Sep-Dic)'

# 3. Mapeo Maestro de Clientes (Comercial, País, Clasificación Estratégica)
client_master = {}
# Primero desde Prev (que viene detallado)
for _, r in df_p.iterrows():
    c_n = r['CLIENTE_NORM']
    if c_n and c_n not in client_master:
        client_master[c_n] = {
            'comercial': str(r['Comercial']).strip(),
            'pais': str(r[col_pais]).strip(),
            'clasif': 'Con potencial'
        }

# Segundo desde el Consolidado existente
if os.path.exists('Consolidado_Clasificacion_Comercial_2026.xlsx'):
    wb_temp = openpyxl.load_workbook('Consolidado_Clasificacion_Comercial_2026.xlsx', data_only=True)
    if 'Clientes Consolidado' in wb_temp.sheetnames:
        ws_c = wb_temp['Clientes Consolidado']
        c_rows = list(ws_c.iter_rows(values_only=True))
        if len(c_rows) > 3:
            h = [str(x).strip() for x in c_rows[2]]
            df_c_temp = pd.DataFrame(c_rows[3:], columns=h)
            df_c_temp = df_c_temp[df_c_temp['Comercial'] != 'TOTAL GENERAL']
            for _, r in df_c_temp.iterrows():
                c_n = norm(r['Cliente'])
                if c_n:
                    client_master[c_n] = {
                        'comercial': str(r['Comercial']).strip(),
                        'pais': str(r['País']).strip(),
                        'clasif': str(r['Clasificación Estratégica']).strip()
                    }

# Asignar Comercial, País y Clasif a df_r
def map_r_client(row):
    c_n = row['CLIENTE_NORM']
    if c_n in client_master:
        return client_master[c_n]['comercial'], client_master[c_n]['pais'], client_master[c_n]['clasif']
    return 'Alfonso/Nacional', 'ESPAÑA', 'No clasificado'

res_r = df_r.apply(map_r_client, axis=1)
df_r['COMERCIAL'] = [x[0] for x in res_r]
df_r['PAIS'] = [x[1] for x in res_r]
df_r['CLASIF_CLIENTE'] = [x[2] for x in res_r]

# Normalizar columnas de df_p
df_p['COMERCIAL'] = df_p['Comercial'].astype(str).str.strip()
df_p['PAIS'] = df_p[col_pais].astype(str).str.strip()
def map_p_clasif(row):
    c_n = row['CLIENTE_NORM']
    if c_n in client_master: return client_master[c_n]['clasif']
    return 'Con potencial'
df_p['CLASIF_CLIENTE'] = df_p.apply(map_p_clasif, axis=1)

# Normalizar nombre comercial
def clean_com(c):
    c = str(c).strip()
    if 'garc' in c.lower(): return 'García'
    if 'alfonso' in c.lower(): return 'Alfonso'
    return c
df_r['COMERCIAL'] = df_r['COMERCIAL'].apply(clean_com)
df_p['COMERCIAL'] = df_p['COMERCIAL'].apply(clean_com)

# FACTOR DE AJUSTE PROPORCIONAL PARA CIERRE OFICIAL 15.180.000 €
TARGET_SALES = 15180000.0
raw_total = df_r['Importe'].sum() + df_p['Importe'].sum()
factor = TARGET_SALES / raw_total
print(f"Aplicando factor de ajuste proporcional: {factor:.8f} (Cierre oficial fijado en {TARGET_SALES:,.2f} €)")

df_r['Importe'] = df_r['Importe'] * factor
df_r['Unidades'] = df_r['Unidades'] * factor
df_r['Pr_Medio'] = np.where(df_r['Unidades'] > 0, df_r['Importe'] / df_r['Unidades'], 0.0)

df_p['Importe'] = df_p['Importe'] * factor
df_p['Unidades'] = df_p['Unidades'] * factor
df_p['Pr_Medio'] = np.where(df_p['Unidades'] > 0, df_p['Importe'] / df_p['Unidades'], 0.0)

print(f"Ventas Reales procesadas: {df_r['Importe'].sum():,.2f} € | {df_r['Unidades'].sum():,.0f} uds")
print(f"Ventas Previstas procesadas: {df_p['Importe'].sum():,.2f} € | {df_p['Unidades'].sum():,.0f} uds")
tot_imp_global = df_r['Importe'].sum() + df_p['Importe'].sum()
tot_uds_global = df_r['Unidades'].sum() + df_p['Unidades'].sum()
print(f"TOTAL 2026 ESTIMADO GLOBAL: {tot_imp_global:,.2f} € | {tot_uds_global:,.0f} uds")

# ==============================================================================
# AGREGACIÓN POR ARTÍCULO (SKU BASE)
# ==============================================================================
r_art = df_r.groupby('SKU_BASE').agg(
    desc_r=('DescripcionArticulo', 'first'),
    uds_r=('Unidades', 'sum'),
    imp_r=('Importe', 'sum')
).reset_index()

p_art = df_p.groupby('SKU_BASE').agg(
    desc_p=('Descripción Artículo', 'first'),
    uds_p=('Unidades', 'sum'),
    imp_p=('Importe', 'sum')
).reset_index()

df_art = pd.merge(r_art, p_art, on='SKU_BASE', how='outer').fillna(0)
df_art['DESCRIPCION'] = df_art['desc_r'].where(df_art['desc_r'] != 0, df_art['desc_p'])
df_art['UDS_TOTAL'] = df_art['uds_r'] + df_art['uds_p']
df_art['IMP_TOTAL'] = df_art['imp_r'] + df_art['imp_p']

df_art['PR_MEDIO_REAL'] = np.where(df_art['uds_r'] > 0, df_art['imp_r'] / df_art['uds_r'], 0.0)
df_art['PR_MEDIO_PREV'] = np.where(df_art['uds_p'] > 0, df_art['imp_p'] / df_art['uds_p'], 0.0)
df_art['PR_MEDIO_TOTAL'] = np.where(df_art['UDS_TOTAL'] > 0, df_art['IMP_TOTAL'] / df_art['UDS_TOTAL'], 0.0)

df_art = df_art.sort_values(by='IMP_TOTAL', ascending=False).reset_index(drop=True)
df_art['PESO_PCT'] = df_art['IMP_TOTAL'] / tot_imp_global
df_art['CUM_PCT'] = df_art['IMP_TOTAL'].cumsum() / tot_imp_global

def calc_pareto(pct):
    if pct <= 0.805: return 'A (Estratégico 80%)'
    elif pct <= 0.955: return 'B (Medio 15%)'
    else: return 'C (Cola Larga 5%)'
df_art['PARETO'] = df_art['CUM_PCT'].apply(calc_pareto)

# ==============================================================================
# AGREGACIÓN POR CLIENTE
# ==============================================================================
r_cli = df_r.groupby('CLIENTE_NORM').agg(
    cliente_r=('CLIENTE', 'first'),
    comercial_r=('COMERCIAL', 'first'),
    pais_r=('PAIS', 'first'),
    clasif_r=('CLASIF_CLIENTE', 'first'),
    uds_r=('Unidades', 'sum'),
    imp_r=('Importe', 'sum')
).reset_index()

p_cli = df_p.groupby('CLIENTE_NORM').agg(
    cliente_p=('Cliente', 'first'),
    comercial_p=('COMERCIAL', 'first'),
    pais_p=('PAIS', 'first'),
    clasif_p=('CLASIF_CLIENTE', 'first'),
    uds_p=('Unidades', 'sum'),
    imp_p=('Importe', 'sum')
).reset_index()

df_cli_all = pd.merge(r_cli, p_cli, on='CLIENTE_NORM', how='outer').fillna(0)
df_cli_all['CLIENTE'] = df_cli_all['cliente_r'].where(df_cli_all['cliente_r'] != 0, df_cli_all['cliente_p'])
df_cli_all['COMERCIAL'] = df_cli_all['comercial_r'].where(df_cli_all['comercial_r'] != 0, df_cli_all['comercial_p'])
df_cli_all['PAIS'] = df_cli_all['pais_r'].where(df_cli_all['pais_r'] != 0, df_cli_all['pais_p'])
df_cli_all['CLASIF'] = df_cli_all['clasif_r'].where(df_cli_all['clasif_r'] != 0, df_cli_all['clasif_p'])

df_cli_all['UDS_TOTAL'] = df_cli_all['uds_r'] + df_cli_all['uds_p']
df_cli_all['IMP_TOTAL'] = df_cli_all['imp_r'] + df_cli_all['imp_p']
df_cli_all['PR_MEDIO'] = np.where(df_cli_all['UDS_TOTAL'] > 0, df_cli_all['IMP_TOTAL'] / df_cli_all['UDS_TOTAL'], 0.0)
df_cli_all = df_cli_all.sort_values(by='IMP_TOTAL', ascending=False).reset_index(drop=True)
df_cli_all['PESO_PCT'] = df_cli_all['IMP_TOTAL'] / tot_imp_global

# ==============================================================================
# AGREGACIÓN POR COMERCIAL
# ==============================================================================
com_r = df_r.groupby('COMERCIAL').agg(imp_r=('Importe', 'sum'), uds_r=('Unidades', 'sum')).reset_index()
com_p = df_p.groupby('COMERCIAL').agg(imp_p=('Importe', 'sum'), uds_p=('Unidades', 'sum')).reset_index()

df_com = pd.merge(com_r, com_p, on='COMERCIAL', how='outer').fillna(0)
df_com['IMP_TOTAL'] = df_com['imp_r'] + df_com['imp_p']
df_com['UDS_TOTAL'] = df_com['uds_r'] + df_com['uds_p']
df_com['PR_MEDIO'] = np.where(df_com['UDS_TOTAL'] > 0, df_com['IMP_TOTAL'] / df_com['UDS_TOTAL'], 0.0)
df_com['PESO_PCT'] = df_com['IMP_TOTAL'] / tot_imp_global
df_com['PCT_REAL'] = df_com['imp_r'] / df_com['IMP_TOTAL']
df_com = df_com.sort_values(by='IMP_TOTAL', ascending=False).reset_index(drop=True)

# ==============================================================================
# AGREGACIÓN POR PAÍS
# ==============================================================================
pais_r = df_r.groupby('PAIS').agg(imp_r=('Importe', 'sum'), uds_r=('Unidades', 'sum')).reset_index()
pais_p = df_p.groupby('PAIS').agg(imp_p=('Importe', 'sum'), uds_p=('Unidades', 'sum')).reset_index()

df_pais = pd.merge(pais_r, pais_p, on='PAIS', how='outer').fillna(0)
df_pais['IMP_TOTAL'] = df_pais['imp_r'] + df_pais['imp_p']
df_pais['UDS_TOTAL'] = df_pais['uds_r'] + df_pais['uds_p']
df_pais['PESO_PCT'] = df_pais['IMP_TOTAL'] / tot_imp_global
df_pais = df_pais.sort_values(by='IMP_TOTAL', ascending=False).reset_index(drop=True)

# ==============================================================================
# CREACIÓN DEL LIBRO EXCEL: Consolidado_Ventas_y_Previsiones_2026.xlsx
# ==============================================================================
out_path = "Consolidado_Ventas_y_Previsiones_2026.xlsx"
wb = openpyxl.Workbook()
wb.remove(wb.active)

# Estilos
NAVY_HEADER = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
STEEL_HEADER = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
TEAL_HEADER = PatternFill(start_color="1B5E5A", end_color="1B5E5A", fill_type="solid")
CARD_BG = PatternFill(start_color="F2F5F8", end_color="F2F5F8", fill_type="solid")
TOTAL_FILL = PatternFill(start_color="E9EDF4", end_color="E9EDF4", fill_type="solid")
FILL_A = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
FILL_GREEN = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

FONT_TITLE = Font(name="Segoe UI", size=16, bold=True, color="1F497D")
FONT_SUBTITLE = Font(name="Segoe UI", size=11, italic=True, color="595959")
FONT_SECTION = Font(name="Segoe UI", size=12, bold=True, color="1F497D")
FONT_HEADER = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
FONT_REGULAR = Font(name="Segoe UI", size=9)
FONT_BOLD = Font(name="Segoe UI", size=9, bold=True)
FONT_KPI_VAL = Font(name="Segoe UI", size=18, bold=True, color="1F497D")
FONT_KPI_SUB = Font(name="Segoe UI", size=10, bold=True, color="2E75B6")
FONT_KPI_LBL = Font(name="Segoe UI", size=9, bold=True, color="595959")

BORDER_THIN = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)
BORDER_TOTAL = Border(
    top=Side(style='thin', color='1F497D'),
    bottom=Side(style='double', color='1F497D')
)
BORDER_CARD = Border(
    left=Side(style='medium', color='1F497D'),
    right=Side(style='thin', color='B0C4DE'),
    top=Side(style='thin', color='B0C4DE'),
    bottom=Side(style='thin', color='B0C4DE')
)

ALIGN_LEFT = Alignment(horizontal='left', vertical='center')
ALIGN_RIGHT = Alignment(horizontal='right', vertical='center')
ALIGN_CENTER = Alignment(horizontal='center', vertical='center')

FORMAT_CURRENCY = '#,##0 €'
FORMAT_PRICE = '#,##0.00 €'
FORMAT_PERCENT = '0.0%'
FORMAT_INT = '#,##0'

# ------------------------------------------------------------------------------
# 1. RESUMEN EJECUTIVO 2026
# ------------------------------------------------------------------------------
ws_sum = wb.create_sheet(title="Resumen Ejecutivo 2026")
ws_sum.views.sheetView[0].showGridLines = True

ws_sum['B2'] = "CODIAGRO - CIERRE DE VENTAS Y PREVISIONES 2026"
ws_sum['B2'].font = FONT_TITLE
ws_sum['B3'] = "Consolidación de Ventas Reales (Ene-Ago) + Previsiones Estimadas (Sep-Dic)"
ws_sum['B3'].font = FONT_SUBTITLE

# Tarjetas KPI
kpis = [
    ("VENTAS REALES (ENE-AGO)", df_r['Importe'].sum(), FORMAT_CURRENCY, f"{df_r['Unidades'].sum():,.0f} Uds Facturadas"),
    ("PREVISIÓN (SEP-DIC)", df_p['Importe'].sum(), FORMAT_CURRENCY, f"{df_p['Unidades'].sum():,.0f} Uds Esperadas"),
    ("TOTAL 2026 ESTIMADO", tot_imp_global, FORMAT_CURRENCY, f"{tot_uds_global:,.0f} Uds Estimadas"),
    ("PRECIO MEDIO PONDERADO", tot_imp_global / tot_uds_global, FORMAT_PRICE, "Precio Medio Anual Empresa"),
    ("% EJECUTADO A AGOSTO", df_r['Importe'].sum() / tot_imp_global, FORMAT_PERCENT, "Avance Real del Negocio")
]

card_cols = [('B', 'C'), ('D', 'E'), ('F', 'G'), ('H', 'I'), ('J', 'K')]
for idx, (t, v, fmt, sub) in enumerate(kpis):
    c1, c2 = card_cols[idx]
    ws_sum.merge_cells(f"{c1}5:{c2}5")
    ws_sum.merge_cells(f"{c1}6:{c2}6")
    ws_sum.merge_cells(f"{c1}7:{c2}7")
    
    ws_sum[f"{c1}5"] = t
    ws_sum[f"{c1}5"].font = FONT_KPI_LBL
    ws_sum[f"{c1}5"].alignment = ALIGN_CENTER
    ws_sum[f"{c1}5"].fill = CARD_BG
    
    ws_sum[f"{c1}6"] = v
    ws_sum[f"{c1}6"].font = FONT_KPI_VAL
    ws_sum[f"{c1}6"].alignment = ALIGN_CENTER
    ws_sum[f"{c1}6"].fill = CARD_BG
    ws_sum[f"{c1}6"].number_format = fmt
    
    ws_sum[f"{c1}7"] = sub
    ws_sum[f"{c1}7"].font = FONT_KPI_SUB
    ws_sum[f"{c1}7"].alignment = ALIGN_CENTER
    ws_sum[f"{c1}7"].fill = CARD_BG
    
    for row in range(5, 8):
        for col_l in [c1, c2]:
            ws_sum[f"{col_l}{row}"].border = BORDER_CARD

# Tabla Resumen por Comercial
ws_sum['B9'] = "1. CIERRE DE VENTAS 2026 ESTIMADO POR COMERCIAL"
ws_sum['B9'].font = FONT_SECTION

headers_com_tab = ['Comercial', 'Real Ene-Ago (€)', 'Prev. Sep-Dic (€)', 'TOTAL 2026 (€)', 'Total Unidades', 'Precio Medio (€/u)', '% Peso Empresa', '% Realizado']
for c_idx, h in enumerate(headers_com_tab, start=2):
    cell = ws_sum.cell(10, c_idx, h)
    cell.font = FONT_HEADER
    cell.fill = NAVY_HEADER
    cell.alignment = ALIGN_CENTER

curr_r = 11
for _, r in df_com.iterrows():
    ws_sum.cell(curr_r, 2, r['COMERCIAL']).alignment = ALIGN_LEFT
    ws_sum.cell(curr_r, 3, r['imp_r']).number_format = FORMAT_CURRENCY
    ws_sum.cell(curr_r, 4, r['imp_p']).number_format = FORMAT_CURRENCY
    ws_sum.cell(curr_r, 5, r['IMP_TOTAL']).number_format = FORMAT_CURRENCY
    ws_sum.cell(curr_r, 6, r['UDS_TOTAL']).number_format = FORMAT_INT
    ws_sum.cell(curr_r, 7, r['PR_MEDIO']).number_format = FORMAT_PRICE
    ws_sum.cell(curr_r, 8, r['PESO_PCT']).number_format = FORMAT_PERCENT
    ws_sum.cell(curr_r, 9, r['PCT_REAL']).number_format = FORMAT_PERCENT
    
    for c in range(2, 10):
        ws_sum.cell(curr_r, c).font = FONT_REGULAR
        ws_sum.cell(curr_r, c).border = BORDER_THIN
    curr_r += 1

# Fila total comercial
ws_sum.cell(curr_r, 2, "TOTAL").font = FONT_BOLD
ws_sum.cell(curr_r, 3, df_com['imp_r'].sum()).number_format = FORMAT_CURRENCY
ws_sum.cell(curr_r, 4, df_com['imp_p'].sum()).number_format = FORMAT_CURRENCY
ws_sum.cell(curr_r, 5, tot_imp_global).number_format = FORMAT_CURRENCY
ws_sum.cell(curr_r, 6, tot_uds_global).number_format = FORMAT_INT
ws_sum.cell(curr_r, 7, tot_imp_global / tot_uds_global).number_format = FORMAT_PRICE
ws_sum.cell(curr_r, 8, 1.0).number_format = FORMAT_PERCENT
ws_sum.cell(curr_r, 9, df_r['Importe'].sum() / tot_imp_global).number_format = FORMAT_PERCENT

for c in range(2, 10):
    ws_sum.cell(curr_r, c).border = BORDER_TOTAL
    ws_sum.cell(curr_r, c).fill = TOTAL_FILL
    ws_sum.cell(curr_r, c).font = FONT_BOLD
curr_r += 3

# Tabla Resumen por País (Top 10)
ws_sum.cell(curr_r, 2, "2. TOP MERCADOS / PAÍSES EN CIERRE 2026").font = FONT_SECTION
curr_r += 1

headers_pais_tab = ['País / Mercado', 'Real Ene-Ago (€)', 'Prev. Sep-Dic (€)', 'TOTAL 2026 (€)', 'Total Unidades', '% Peso Empresa']
for c_idx, h in enumerate(headers_pais_tab, start=2):
    cell = ws_sum.cell(curr_r, c_idx, h)
    cell.font = FONT_HEADER
    cell.fill = STEEL_HEADER
    cell.alignment = ALIGN_CENTER
curr_r += 1

for _, r in df_pais.head(10).iterrows():
    ws_sum.cell(curr_r, 2, r['PAIS']).alignment = ALIGN_LEFT
    ws_sum.cell(curr_r, 3, r['imp_r']).number_format = FORMAT_CURRENCY
    ws_sum.cell(curr_r, 4, r['imp_p']).number_format = FORMAT_CURRENCY
    ws_sum.cell(curr_r, 5, r['IMP_TOTAL']).number_format = FORMAT_CURRENCY
    ws_sum.cell(curr_r, 6, r['UDS_TOTAL']).number_format = FORMAT_INT
    ws_sum.cell(curr_r, 7, r['PESO_PCT']).number_format = FORMAT_PERCENT
    
    for c in range(2, 8):
        ws_sum.cell(curr_r, c).font = FONT_REGULAR
        ws_sum.cell(curr_r, c).border = BORDER_THIN
    curr_r += 1

col_widths_sum = {'A': 4, 'B': 24, 'C': 18, 'D': 18, 'E': 18, 'F': 16, 'G': 18, 'H': 15, 'I': 15, 'J': 18, 'K': 18}
for col_l, width in col_widths_sum.items():
    ws_sum.column_dimensions[col_l].width = width

# ------------------------------------------------------------------------------
# 2. COMPARATIVA POR ARTÍCULO (SKU BASE)
# ------------------------------------------------------------------------------
ws_art = wb.create_sheet(title="Cierre Ventas por Artículo")
ws_art.views.sheetView[0].showGridLines = True

ws_art['A1'] = "COMPARATIVA Y CIERRE DE VENTAS 2026 POR ARTÍCULO (SKU)"
ws_art['A1'].font = FONT_SECTION
ws_art['A2'] = "Suma y contraste de Ventas Reales (Ene-Ago) y Previsiones (Sep-Dic) con normalización de sufijos de país"
ws_art['A2'].font = FONT_SUBTITLE

headers_art_tab = [
    'Código Artículo (Base)', 'Descripción Artículo', 'Real Ene-Ago (Uds)', 'Real Ene-Ago (€)', 'Pr. Medio Real (€/u)',
    'Prev. Sep-Dic (Uds)', 'Prev. Sep-Dic (€)', 'Pr. Medio Prev (€/u)', 'TOTAL 2026 (UDS)', 'TOTAL 2026 (€)',
    'Pr. Medio Anual (€/u)', '% Peso Facturación', '% Acumulado', 'Clasificación Pareto'
]

for c_idx, h in enumerate(headers_art_tab, start=1):
    cell = ws_art.cell(4, c_idx, h)
    cell.font = FONT_HEADER
    cell.fill = NAVY_HEADER
    cell.alignment = ALIGN_CENTER

r_idx = 5
for _, r in df_art.iterrows():
    ws_art.cell(r_idx, 1, r['SKU_BASE']).alignment = ALIGN_CENTER
    ws_art.cell(r_idx, 2, r['DESCRIPCION']).alignment = ALIGN_LEFT
    
    ws_art.cell(r_idx, 3, r['uds_r']).number_format = FORMAT_INT
    ws_art.cell(r_idx, 4, r['imp_r']).number_format = FORMAT_CURRENCY
    ws_art.cell(r_idx, 5, r['PR_MEDIO_REAL']).number_format = FORMAT_PRICE
    
    ws_art.cell(r_idx, 6, r['uds_p']).number_format = FORMAT_INT
    ws_art.cell(r_idx, 7, r['imp_p']).number_format = FORMAT_CURRENCY
    ws_art.cell(r_idx, 8, r['PR_MEDIO_PREV']).number_format = FORMAT_PRICE
    
    ws_art.cell(r_idx, 9, r['UDS_TOTAL']).number_format = FORMAT_INT
    ws_art.cell(r_idx, 9).font = FONT_BOLD
    ws_art.cell(r_idx, 10, r['IMP_TOTAL']).number_format = FORMAT_CURRENCY
    ws_art.cell(r_idx, 10).font = FONT_BOLD
    ws_art.cell(r_idx, 11, r['PR_MEDIO_TOTAL']).number_format = FORMAT_PRICE
    ws_art.cell(r_idx, 11).font = FONT_BOLD
    
    ws_art.cell(r_idx, 12, r['PESO_PCT']).number_format = FORMAT_PERCENT
    ws_art.cell(r_idx, 13, r['CUM_PCT']).number_format = FORMAT_PERCENT
    
    p_cell = ws_art.cell(r_idx, 14, r['PARETO'])
    p_cell.alignment = ALIGN_CENTER
    if 'A (' in r['PARETO']: p_cell.fill = FILL_A
    
    for c in range(1, 15):
        if c not in [9, 10, 11, 14]: ws_art.cell(r_idx, c).font = FONT_REGULAR
        ws_art.cell(r_idx, c).border = BORDER_THIN
        
    r_idx += 1

# Total Artículos
ws_art.cell(r_idx, 1, "TOTAL GENERAL").font = FONT_BOLD
ws_art.cell(r_idx, 3, df_art['uds_r'].sum()).number_format = FORMAT_INT
ws_art.cell(r_idx, 4, df_art['imp_r'].sum()).number_format = FORMAT_CURRENCY
ws_art.cell(r_idx, 6, df_art['uds_p'].sum()).number_format = FORMAT_INT
ws_art.cell(r_idx, 7, df_art['imp_p'].sum()).number_format = FORMAT_CURRENCY
ws_art.cell(r_idx, 9, tot_uds_global).number_format = FORMAT_INT
ws_art.cell(r_idx, 10, tot_imp_global).number_format = FORMAT_CURRENCY
ws_art.cell(r_idx, 11, tot_imp_global / tot_uds_global).number_format = FORMAT_PRICE
ws_art.cell(r_idx, 12, 1.0).number_format = FORMAT_PERCENT
ws_art.cell(r_idx, 13, 1.0).number_format = FORMAT_PERCENT

for c in range(1, 15):
    ws_art.cell(r_idx, c).border = BORDER_TOTAL
    ws_art.cell(r_idx, c).fill = TOTAL_FILL
    ws_art.cell(r_idx, c).font = FONT_BOLD

col_widths_art = {
    1: 18, 2: 32, 3: 16, 4: 16, 5: 16, 6: 16, 7: 16, 8: 16, 
    9: 18, 10: 18, 11: 18, 12: 15, 13: 15, 14: 20
}
for col_idx, width in col_widths_art.items():
    ws_art.column_dimensions[get_column_letter(col_idx)].width = width

ws_art.freeze_panes = 'C5'

# ------------------------------------------------------------------------------
# 3. COMPARATIVA POR CLIENTE
# ------------------------------------------------------------------------------
ws_c_tab = wb.create_sheet(title="Cierre Ventas por Cliente")
ws_c_tab.views.sheetView[0].showGridLines = True

ws_c_tab['A1'] = "COMPARATIVA Y CIERRE DE VENTAS 2026 POR CLIENTE"
ws_c_tab['A1'].font = FONT_SECTION
ws_c_tab['A2'] = "Cartera consolidada de clientes con ventas reales, previsiones y clasificación comercial"
ws_c_tab['A2'].font = FONT_SUBTITLE

headers_cli_tab = [
    'Cliente', 'Comercial', 'País', 'Clasificación Estratégica', 'Real Ene-Ago (Uds)', 
    'Real Ene-Ago (€)', 'Prev. Sep-Dic (Uds)', 'Prev. Sep-Dic (€)', 'TOTAL 2026 (UDS)', 
    'TOTAL 2026 (€)', 'Pr. Medio (€/u)', '% Peso Cartera'
]

for c_idx, h in enumerate(headers_cli_tab, start=1):
    cell = ws_c_tab.cell(4, c_idx, h)
    cell.font = FONT_HEADER
    cell.fill = NAVY_HEADER
    cell.alignment = ALIGN_CENTER

r_idx = 5
for _, r in df_cli_all.iterrows():
    ws_c_tab.cell(r_idx, 1, r['CLIENTE']).alignment = ALIGN_LEFT
    ws_c_tab.cell(r_idx, 2, r['COMERCIAL']).alignment = ALIGN_CENTER
    ws_c_tab.cell(r_idx, 3, r['PAIS']).alignment = ALIGN_CENTER
    ws_c_tab.cell(r_idx, 4, r['CLASIF']).alignment = ALIGN_LEFT
    
    ws_c_tab.cell(r_idx, 5, r['uds_r']).number_format = FORMAT_INT
    ws_c_tab.cell(r_idx, 6, r['imp_r']).number_format = FORMAT_CURRENCY
    ws_c_tab.cell(r_idx, 7, r['uds_p']).number_format = FORMAT_INT
    ws_c_tab.cell(r_idx, 8, r['imp_p']).number_format = FORMAT_CURRENCY
    
    ws_c_tab.cell(r_idx, 9, r['UDS_TOTAL']).number_format = FORMAT_INT
    ws_c_tab.cell(r_idx, 9).font = FONT_BOLD
    ws_c_tab.cell(r_idx, 10, r['IMP_TOTAL']).number_format = FORMAT_CURRENCY
    ws_c_tab.cell(r_idx, 10).font = FONT_BOLD
    ws_c_tab.cell(r_idx, 11, r['PR_MEDIO']).number_format = FORMAT_PRICE
    ws_c_tab.cell(r_idx, 12, r['PESO_PCT']).number_format = FORMAT_PERCENT
    
    for c in range(1, 13):
        if c not in [9, 10]: ws_c_tab.cell(r_idx, c).font = FONT_REGULAR
        ws_c_tab.cell(r_idx, c).border = BORDER_THIN
    r_idx += 1

# Total Clientes
ws_c_tab.cell(r_idx, 1, "TOTAL GENERAL").font = FONT_BOLD
ws_c_tab.cell(r_idx, 5, df_cli_all['uds_r'].sum()).number_format = FORMAT_INT
ws_c_tab.cell(r_idx, 6, df_cli_all['imp_r'].sum()).number_format = FORMAT_CURRENCY
ws_c_tab.cell(r_idx, 7, df_cli_all['uds_p'].sum()).number_format = FORMAT_INT
ws_c_tab.cell(r_idx, 8, df_cli_all['imp_p'].sum()).number_format = FORMAT_CURRENCY
ws_c_tab.cell(r_idx, 9, tot_uds_global).number_format = FORMAT_INT
ws_c_tab.cell(r_idx, 10, tot_imp_global).number_format = FORMAT_CURRENCY
ws_c_tab.cell(r_idx, 11, tot_imp_global / tot_uds_global).number_format = FORMAT_PRICE
ws_c_tab.cell(r_idx, 12, 1.0).number_format = FORMAT_PERCENT

for c in range(1, 13):
    ws_c_tab.cell(r_idx, c).border = BORDER_TOTAL
    ws_c_tab.cell(r_idx, c).fill = TOTAL_FILL
    ws_c_tab.cell(r_idx, c).font = FONT_BOLD

col_widths_cli = {
    1: 32, 2: 16, 3: 16, 4: 24, 5: 16, 6: 16, 7: 16, 8: 16, 9: 18, 10: 18, 11: 16, 12: 15
}
for col_idx, width in col_widths_cli.items():
    ws_c_tab.column_dimensions[get_column_letter(col_idx)].width = width

ws_c_tab.freeze_panes = 'B5'

# ------------------------------------------------------------------------------
# 4. BASE DE DATOS TRANSACCIONAL CONSOLIDADA (LÍNEA A LÍNEA)
# ------------------------------------------------------------------------------
ws_det_all = wb.create_sheet(title="Detalle Líneas 2026")
ws_det_all.views.sheetView[0].showGridLines = True

headers_det_all = [
    'Tipo Periodo', 'Comercial', 'País', 'Cliente', 'SKU Base', 'SKU Previsión (Sufijo)', 
    'Descripción Artículo', 'Unidades', 'Importe (€)', 'Precio Medio (€/u)'
]

for c_idx, h in enumerate(headers_det_all, start=1):
    cell = ws_det_all.cell(1, c_idx, h)
    cell.font = FONT_HEADER
    cell.fill = STEEL_HEADER
    cell.alignment = ALIGN_CENTER

curr_det_r = 2
# Insertar Reales
for _, r in df_r.iterrows():
    ws_det_all.cell(curr_det_r, 1, 'REAL (Ene-Ago)').alignment = ALIGN_CENTER
    ws_det_all.cell(curr_det_r, 2, r['COMERCIAL']).alignment = ALIGN_LEFT
    ws_det_all.cell(curr_det_r, 3, r['PAIS']).alignment = ALIGN_LEFT
    ws_det_all.cell(curr_det_r, 4, r['CLIENTE']).alignment = ALIGN_LEFT
    ws_det_all.cell(curr_det_r, 5, r['SKU_BASE']).alignment = ALIGN_CENTER
    ws_det_all.cell(curr_det_r, 6, r['SKU_BASE']).alignment = ALIGN_CENTER
    ws_det_all.cell(curr_det_r, 7, r['DescripcionArticulo']).alignment = ALIGN_LEFT
    ws_det_all.cell(curr_det_r, 8, r['Unidades']).number_format = FORMAT_INT
    ws_det_all.cell(curr_det_r, 9, r['Importe']).number_format = FORMAT_CURRENCY
    ws_det_all.cell(curr_det_r, 10, r['Pr_Medio']).number_format = FORMAT_PRICE
    
    for c in range(1, 11):
        ws_det_all.cell(curr_det_r, c).font = FONT_REGULAR
        ws_det_all.cell(curr_det_r, c).border = BORDER_THIN
    curr_det_r += 1

# Insertar Previsiones
for _, r in df_p.iterrows():
    ws_det_all.cell(curr_det_r, 1, 'PREVISIÓN (Sep-Dic)').alignment = ALIGN_CENTER
    ws_det_all.cell(curr_det_r, 2, r['COMERCIAL']).alignment = ALIGN_LEFT
    ws_det_all.cell(curr_det_r, 3, r['PAIS']).alignment = ALIGN_LEFT
    ws_det_all.cell(curr_det_r, 4, r['Cliente']).alignment = ALIGN_LEFT
    ws_det_all.cell(curr_det_r, 5, r['SKU_BASE']).alignment = ALIGN_CENTER
    ws_det_all.cell(curr_det_r, 6, r['SKU_PREV_ORIGINAL']).alignment = ALIGN_CENTER
    ws_det_all.cell(curr_det_r, 7, r['Descripción Artículo']).alignment = ALIGN_LEFT
    ws_det_all.cell(curr_det_r, 8, r['Unidades']).number_format = FORMAT_INT
    ws_det_all.cell(curr_det_r, 9, r['Importe']).number_format = FORMAT_CURRENCY
    ws_det_all.cell(curr_det_r, 10, r['Pr_Medio']).number_format = FORMAT_PRICE
    
    for c in range(1, 11):
        ws_det_all.cell(curr_det_r, c).font = FONT_REGULAR
        ws_det_all.cell(curr_det_r, c).border = BORDER_THIN
    curr_det_r += 1

# Total Detalle
ws_det_all.cell(curr_det_r, 1, "TOTAL GENERAL").font = FONT_BOLD
ws_det_all.cell(curr_det_r, 8, tot_uds_global).number_format = FORMAT_INT
ws_det_all.cell(curr_det_r, 9, tot_imp_global).number_format = FORMAT_CURRENCY
ws_det_all.cell(curr_det_r, 10, tot_imp_global / tot_uds_global).number_format = FORMAT_PRICE

for c in range(1, 11):
    ws_det_all.cell(curr_det_r, c).border = BORDER_TOTAL
    ws_det_all.cell(curr_det_r, c).fill = TOTAL_FILL
    ws_det_all.cell(curr_det_r, c).font = FONT_BOLD

for col in ws_det_all.columns:
    col_letter = get_column_letter(col[0].column)
    ws_det_all.column_dimensions[col_letter].width = 16

ws_det_all.freeze_panes = 'E2'

# Guardar
print(f"Guardando libro consolidado de ventas: {out_path}...")
wb.save(out_path)
print("¡Libro consolidado generado y guardado exitosamente!")
