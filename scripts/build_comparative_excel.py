import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd

DIR = r'Gastos'
EXCEL_OUT = os.path.join(DIR, 'Comparativa_Presupuestos_Departamentos_vs_Contable.xlsx')

wb = openpyxl.Workbook()
wb.remove(wb.active) # Remove default sheet

# Typography & Palette
FONT_TITLE = Font(name='Calibri', size=16, bold=True, color='1F4E79')
FONT_SUBTITLE = Font(name='Calibri', size=10, italic=True, color='595959')
FONT_SECTION = Font(name='Calibri', size=12, bold=True, color='FFFFFF')
FONT_TH = Font(name='Calibri', size=10, bold=True, color='FFFFFF')
FONT_DATA = Font(name='Calibri', size=10)
FONT_BOLD = Font(name='Calibri', size=10, bold=True)
FONT_TOTAL = Font(name='Calibri', size=11, bold=True, color='1F4E79')
FONT_ALERT = Font(name='Calibri', size=10, bold=True, color='C00000')
FONT_SUCCESS = Font(name='Calibri', size=10, bold=True, color='276A3C')
FONT_NOTE = Font(name='Calibri', size=9, italic=True, color='333333')

FILL_NAVY = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
FILL_HEADER = PatternFill(start_color='2F5597', end_color='2F5597', fill_type='solid')
FILL_SUBHEADER = PatternFill(start_color='41719C', end_color='41719C', fill_type='solid')
FILL_ZEBRA = PatternFill(start_color='F9FAFC', end_color='F9FAFC', fill_type='solid')
FILL_TOTAL = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
FILL_GREEN = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
FILL_YELLOW = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
FILL_ORANGE = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')

BORDER_THIN = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)
BORDER_HEADER = Border(
    left=Side(style='thin', color='FFFFFF'),
    right=Side(style='thin', color='FFFFFF'),
    top=Side(style='thin', color='2F5597'),
    bottom=Side(style='medium', color='1F4E79')
)
BORDER_TOTAL = Border(
    top=Side(style='thin', color='1F4E79'),
    bottom=Side(style='double', color='1F4E79'),
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9')
)

FMT_CURR = '#,##0.00 €;[Red]-#,##0.00 €;"-"'
FMT_PCT = '+0.0%;-0.0%;"0.0%"'
FMT_INT = '#,##0'

ALIGN_LEFT = Alignment(horizontal='left', vertical='center')
ALIGN_RIGHT = Alignment(horizontal='right', vertical='center')
ALIGN_CENTER = Alignment(horizontal='center', vertical='center')
ALIGN_HEADER = Alignment(horizontal='center', vertical='center', wrap_text=True)

# -------------------------------------------------------------
# READ RAW DATA
# -------------------------------------------------------------
# 1. Compras
df_compras_raw = pd.read_excel(os.path.join(DIR, 'Presupuesto_compras_2027_actualizado.xlsx'), sheet_name='Resumen mensual')
compras_m = []
meses_nombres = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic']
for r in range(5, 17):
    mes_name = df_compras_raw.iloc[r, 0]
    sin_res = float(df_compras_raw.iloc[r, 4])
    con_res = float(df_compras_raw.iloc[r, 7])
    compras_m.append({'mes': meses_nombres[r-5], 'sin_res': sin_res, 'con_res': con_res})

# 2. Transporte
df_trans_res = pd.read_excel(os.path.join(DIR, 'Presupuesto_total_transporte_2026_2027.xlsx'), sheet_name='Resumen general')
trans_2026_m = []
for r in range(10, 14):
    trans_2026_m.append({
        'mes': str(df_trans_res.iloc[r, 0]).strip(),
        'resto_nal': float(df_trans_res.iloc[r, 1]),
        'canarias': float(df_trans_res.iloc[r, 2]),
        'export': float(df_trans_res.iloc[r, 3]),
        'total': float(df_trans_res.iloc[r, 4])
    })
trans_2027_m = []
for r in range(19, 31):
    trans_2027_m.append({
        'mes': meses_nombres[r-19],
        'resto_nal': float(df_trans_res.iloc[r, 1]),
        'canarias': float(df_trans_res.iloc[r, 2]),
        'export': float(df_trans_res.iloc[r, 3]),
        'total': float(df_trans_res.iloc[r, 4])
    })

# 3. Mantenimiento
df_maint_raw = pd.read_excel(os.path.join(DIR, 'Presupuesto_mantenimiento_fabrica.xlsx'), header=None)
maint_2027_m = []
for c in range(1, 13):
    base_val = float(df_maint_raw.iloc[9, c])
    conting_val = float(df_maint_raw.iloc[10, c])
    maint_2027_m.append({'mes': meses_nombres[c-1], 'base': base_val, 'con_conting': conting_val})

# Tasks de mantenimiento
maint_tasks = []
for r in range(14, 87):
    row_vals = [df_maint_raw.iloc[r, c] for c in range(13)]
    if pd.notna(row_vals[0]) and pd.notna(row_vals[9]):
        maint_tasks.append({
            'mes': str(row_vals[0]).strip(),
            'area': str(row_vals[1] or '').strip(),
            'categoria': str(row_vals[2] or '').strip(),
            'tipo': str(row_vals[3] or '').strip(),
            'descripcion': str(row_vals[4] or '').strip(),
            'proveedor': str(row_vals[5] or '').strip(),
            'importe': float(row_vals[9])
        })

# 4. Marketing
xl_mkt = pd.ExcelFile(os.path.join(DIR, 'Presupuesto_Marketing_CODIAGRO_2027_V5.xlsx'))
df_mkt_res = xl_mkt.parse('Resumen')
mkt_2027_m = []
for r in range(9, 21):
    mkt_2027_m.append({
        'mes': meses_nombres[r-9],
        'importe': float(df_mkt_res.iloc[r, 5])
    })

# 5. Regulatory
xl_reg = pd.ExcelFile(os.path.join(DIR, 'Presupuesto_regulatory 2026 y 2027.xlsx'))
df_reg_27 = xl_reg.parse('2027')
df_reg_26 = xl_reg.parse('2026')

reg_items_27 = []
for idx, r in df_reg_27.iterrows():
    if pd.notna(r.iloc[1]) and str(r.iloc[1]).strip().lower() != 'total':
        cat = str(r.iloc[0] or '').strip()
        conc = str(r.iloc[1] or '').strip()
        tot = float(r.get('total') or 0)
        months = [float(r.iloc[c]) if pd.notna(r.iloc[c]) else 0.0 for c in range(2, 14)]
        note = str(r.iloc[14]) if len(r) > 14 and pd.notna(r.iloc[14]) else ''
        reg_items_27.append({
            'categoria': cat,
            'concepto': conc,
            'total': tot,
            'meses': months,
            'nota': note
        })

# Current P&L 2027 values for comparison
pnl_ref_2027 = {
    # Compras
    '601000000': {'name': 'Compras de materias primas', 'pnl_2027': 4791964.00, 'pnl_2026': 4065272.93},
    '602000001': {'name': 'Compra de envases y embalajes', 'pnl_2027': 998826.14, 'pnl_2026': 830916.53},
    '602000002': {'name': 'Compra de etiquetas', 'pnl_2027': 101284.89, 'pnl_2026': 85966.48},
    '602000003': {'name': 'Compra de palets', 'pnl_2027': 105364.66, 'pnl_2026': 89435.32},
    '602000006': {'name': 'Compra tapones de envases', 'pnl_2027': 39556.10, 'pnl_2026': 33930.64},
    # Transportes
    '624000001': {'name': 'Transportes nacional', 'pnl_2027': 212857.93, 'pnl_2026': 159898.70},
    '624000002': {'name': 'Transporte internacional (Exportación)', 'pnl_2027': 150774.37, 'pnl_2026': 114577.39},
    '624000003': {'name': 'Transportes Canarias', 'pnl_2027': 88690.80, 'pnl_2026': 69093.88},
    # Mantenimiento
    '622010000': {'name': 'Mantenimiento preventivo general', 'pnl_2027': 17669.40, 'pnl_2026': 7045.20},
    '622010200': {'name': 'Mantenimiento Alcaplant', 'pnl_2027': 30331.56, 'pnl_2026': 26606.58},
    '622010201': {'name': 'Reparaciones Alcaplant', 'pnl_2027': 50245.68, 'pnl_2026': 32054.65},
    '622010300': {'name': 'Mantenimiento Sólidos', 'pnl_2027': 71268.96, 'pnl_2026': 39783.31},
    '622010400': {'name': 'Mantenimiento Líquidos', 'pnl_2027': 10990.80, 'pnl_2026': 8764.59},
    '622010500': {'name': 'Repuestos mantenimiento varios', 'pnl_2027': 9458.40, 'pnl_2026': 7542.61},
    '622010501': {'name': 'Reparación fontanería', 'pnl_2027': 1457.28, 'pnl_2026': 813.48},
    '629000016': {'name': 'Herramientas taller', 'pnl_2027': 6000.36, 'pnl_2026': 2500.15},
    # Marketing
    '627000007': {'name': 'Ferias y congresos', 'pnl_2027': 3855.75, 'pnl_2026': 1696.53},
    '627000008': {'name': 'Marketing y campañas comerciales', 'pnl_2027': 0.00, 'pnl_2026': 9108.58},
    '627000006': {'name': 'Publicidad, catálogos y folletos', 'pnl_2027': 1236.75, 'pnl_2026': 643.11},
    '627000009': {'name': 'Artículos publicidad / promociones', 'pnl_2027': 26500.00, 'pnl_2026': 25894.10},
    '627000004': {'name': 'Relaciones públicas y comerciales', 'pnl_2027': 18104.50, 'pnl_2026': 7965.98},
    # Regulatory
    '629000007': {'name': 'Gastos retirada de residuos', 'pnl_2027': 34000.00, 'pnl_2026': 33169.72},
    '623200000': {'name': 'Gastos marcas y patentes (global)', 'pnl_2027': 0.00, 'pnl_2026': 21459.34},
    '623000003': {'name': 'Servicios registros y notarios', 'pnl_2027': 24000.00, 'pnl_2026': 23694.51},
    '623000004': {'name': 'Gastos certificaciones BCS OKO / Eco', 'pnl_2027': 15253.68, 'pnl_2026': 8897.98},
    '629000008': {'name': 'Cuotas asociaciones (Quimacova)', 'pnl_2027': 9621.84, 'pnl_2026': 6414.57},
    '629000005': {'name': 'Gastos análisis de laboratorios', 'pnl_2027': 11295.24, 'pnl_2026': 6588.88},
    '629000002': {'name': 'Otros gastos de gestión / Envases AEVAE', 'pnl_2027': 11379.82, 'pnl_2026': 11379.82},
}

# Totals by Department (EXCLUDING EXPORT TRANSPORT AS IT IS REPERCUTED TO CUSTOMERS)
tot_compras_27 = sum(m['con_res'] for m in compras_m)
# Transporte neto para Codiagro = Resto Nacional + Canarias (Export se repercute al 100%)
tot_trans_27 = sum(m['resto_nal'] + m['canarias'] for m in trans_2027_m) # 245,670.44 €
tot_trans_export_27 = sum(m['export'] for m in trans_2027_m) # 305,223.30 € (Repercutido a clientes)
tot_maint_27 = sum(m['con_conting'] for m in maint_2027_m) # 91,234.24 €
tot_mkt_27 = sum(m['importe'] for m in mkt_2027_m) # 266,072.40 €
tot_reg_27 = sum(it['total'] for it in reg_items_27) # 233,918.00 €

inmov_reg_marcas = 16300.00
inmov_reg_registros = 25900.00
inmov_reg_nave = 9600.00
inmov_maint = 9410.00

# -------------------------------------------------------------
# SHEET 1: RESUMEN EJECUTIVO
# -------------------------------------------------------------
ws1 = wb.create_sheet(title='Resumen Ejecutivo')
ws1.views.sheetView[0].showGridLines = True

ws1.merge_cells('B2:K2')
ws1['B2'] = 'CODIAGRO · CONTROL PRESUPUESTARIO 2026 - 2027'
ws1['B2'].font = FONT_TITLE

ws1.merge_cells('B3:K3')
ws1['B3'] = 'Comparativa Integral de Presupuestos Departamentales vs P&L (Salvedad: Transporte Exportación Repercutido al Cliente Excluido de Gastos)'
ws1['B3'].font = FONT_SUBTITLE

ws1.merge_cells('B5:K5')
ws1['B5'] = '1. COMPARATIVA GENERAL POR DEPARTAMENTO (PRESUPUESTO 2027)'
ws1['B5'].font = FONT_SECTION
ws1['B5'].fill = FILL_NAVY
ws1['B5'].alignment = Alignment(vertical='center', indent=1)

headers_s1 = [
    'Departamento', 'Cuentas PGC Clave', 'Presupuesto Depto (€)', 'Previsión P&L Actual (€)',
    'Desviación (€)', 'Desviación (%)', 'Partidas Inmovilizables (€)', 'Gasto Operativo Neto (€)',
    'Impacto en EBITDA (€)', 'Estado / Dictamen'
]
for col_idx, h in enumerate(headers_s1, start=2):
    cell = ws1.cell(row=6, column=col_idx, value=h)
    cell.font = FONT_TH
    cell.fill = FILL_HEADER
    cell.alignment = ALIGN_HEADER
    cell.border = BORDER_HEADER
ws1.row_dimensions[6].height = 28

pnl_compras_tot = sum(pnl_ref_2027[k]['pnl_2027'] for k in ['601000000', '602000001', '602000002', '602000003', '602000006'])
# P&L comparable en transporte solo incluye nacional y Canarias, ya que export se repercute
pnl_trans_tot = pnl_ref_2027['624000001']['pnl_2027'] + pnl_ref_2027['624000003']['pnl_2027'] # 301,548.73 €
pnl_maint_tot = sum(pnl_ref_2027[k]['pnl_2027'] for k in ['622010000', '622010200', '622010201', '622010300', '622010400', '622010500', '622010501', '629000016'])
pnl_mkt_tot = sum(pnl_ref_2027[k]['pnl_2027'] for k in ['627000007', '627000008', '627000006', '627000009', '627000004'])
pnl_reg_tot = sum(pnl_ref_2027[k]['pnl_2027'] for k in ['629000007', '623200000', '623000003', '623000004', '629000008', '629000005', '629000002'])

deptos_summary = [
    {'depto': 'Compras y Aprovisionamientos', 'ctas': '601, 6020 (MP, Envases, Etiquetas, Palets)', 'ppto': tot_compras_27, 'pnl': pnl_compras_tot, 'inmov': 0.0, 'dictamen': 'Alineado con necesidades netas (+3% reserva)'},
    {'depto': 'Transportes (Gasto Neto Codiagro)', 'ctas': '624000001, 624000003 (Nacional y Canarias)', 'ppto': tot_trans_27, 'pnl': pnl_trans_tot, 'inmov': 0.0, 'dictamen': 'Ahorro de -55,9k € (-18,5%); Export repercutido al cliente'},
    {'depto': 'Mantenimiento de Fábrica', 'ctas': '6220 (Correctivo, Preventivo, Repuestos)', 'ppto': tot_maint_27, 'pnl': pnl_maint_tot, 'inmov': inmov_maint, 'dictamen': 'Ahorro masivo vs histórico (-106,2k €)'},
    {'depto': 'Marketing y Comunicación', 'ctas': '6270 (Ferias, Campañas, Clientes A/B)', 'ppto': tot_mkt_27, 'pnl': pnl_mkt_tot, 'inmov': 0.0, 'dictamen': 'Plan integral formalizado (1,50% s/ventas)'},
    {'depto': 'Regulatory (Residuos, Marcas, Registros)', 'ctas': '6290, 6230, 6232 (Marcas, AEVAE, Registros)', 'ppto': tot_reg_27, 'pnl': pnl_reg_tot, 'inmov': inmov_reg_marcas + inmov_reg_registros + inmov_reg_nave, 'dictamen': 'Activación clave en Marcas y Registros (51,8k €)'}
]

row_curr = 7
for idx, d in enumerate(deptos_summary):
    fill_row = FILL_ZEBRA if idx % 2 == 1 else PatternFill(fill_type=None)
    ws1.cell(row=row_curr, column=2, value=d['depto']).font = FONT_BOLD
    ws1.cell(row=row_curr, column=3, value=d['ctas']).font = FONT_DATA
    
    c_ppto = ws1.cell(row=row_curr, column=4, value=d['ppto'])
    c_ppto.font = FONT_DATA
    c_ppto.number_format = FMT_CURR
    c_ppto.alignment = ALIGN_RIGHT

    c_pnl = ws1.cell(row=row_curr, column=5, value=d['pnl'])
    c_pnl.font = FONT_DATA
    c_pnl.number_format = FMT_CURR
    c_pnl.alignment = ALIGN_RIGHT

    c_diff = ws1.cell(row=row_curr, column=6, value=f'=D{row_curr}-E{row_curr}')
    c_diff.font = FONT_BOLD
    c_diff.number_format = FMT_CURR
    c_diff.alignment = ALIGN_RIGHT

    c_pct = ws1.cell(row=row_curr, column=7, value=f'=F{row_curr}/E{row_curr}')
    c_pct.font = FONT_DATA
    c_pct.number_format = FMT_PCT
    c_pct.alignment = ALIGN_RIGHT

    c_inmov = ws1.cell(row=row_curr, column=8, value=d['inmov'])
    c_inmov.font = FONT_BOLD if d['inmov'] > 0 else FONT_DATA
    c_inmov.number_format = FMT_CURR
    c_inmov.alignment = ALIGN_RIGHT
    if d['inmov'] > 0: c_inmov.fill = FILL_GREEN

    c_neto = ws1.cell(row=row_curr, column=9, value=f'=D{row_curr}-H{row_curr}')
    c_neto.font = FONT_BOLD
    c_neto.number_format = FMT_CURR
    c_neto.alignment = ALIGN_RIGHT

    c_ebitda = ws1.cell(row=row_curr, column=10, value=f'=E{row_curr}-I{row_curr}')
    c_ebitda.font = FONT_BOLD
    c_ebitda.number_format = FMT_CURR
    c_ebitda.alignment = ALIGN_RIGHT

    c_dict = ws1.cell(row=row_curr, column=11, value=d['dictamen'])
    c_dict.font = FONT_DATA
    c_dict.alignment = ALIGN_LEFT

    for c_i in range(2, 12):
        cell_i = ws1.cell(row=row_curr, column=c_i)
        cell_i.border = BORDER_THIN
        if fill_row.fill_type and c_i != 8: cell_i.fill = fill_row
    row_curr += 1

# Total Row S1
ws1.cell(row=row_curr, column=2, value='TOTAL GASTOS ASUMIDOS POR CODIAGRO').font = FONT_TOTAL
ws1.cell(row=row_curr, column=3, value='').font = FONT_TOTAL
for c_idx, letter in [(4, 'D'), (5, 'E'), (6, 'F'), (8, 'H'), (9, 'I'), (10, 'J')]:
    cell_tot = ws1.cell(row=row_curr, column=c_idx, value=f'=SUM({letter}7:{letter}{row_curr-1})')
    cell_tot.font = FONT_TOTAL
    cell_tot.number_format = FMT_CURR
    cell_tot.alignment = ALIGN_RIGHT
ws1.cell(row=row_curr, column=7, value=f'=F{row_curr}/E{row_curr}').number_format = FMT_PCT
ws1.cell(row=row_curr, column=7).font = FONT_TOTAL
ws1.cell(row=row_curr, column=7).alignment = ALIGN_RIGHT
ws1.cell(row=row_curr, column=11, value='Impacto neto de gasto muy controlado').font = FONT_BOLD
for c_i in range(2, 12):
    c_t = ws1.cell(row=row_curr, column=c_i)
    c_t.border = BORDER_TOTAL
    c_t.fill = FILL_TOTAL

# Callout Alert Box regarding Export Transport
row_curr += 1
ws1.merge_cells(f'B{row_curr}:K{row_curr}')
ws1.cell(row=row_curr, column=2, value='ℹ️ SALVEDAD TRANSPORTE EXPORTACIÓN: El transporte de exportación presupuestado (305.223,30 € en 2027 y 97.236,00 € en Q4 2026) se repercute íntegramente al cliente vía factura (700001624), por lo que NO computa como gasto operativo de Codiagro al tener impacto neto nulo en resultados.').font = FONT_NOTE
ws1.cell(row=row_curr, column=2).fill = FILL_ACCENT_YELLOW if 'FILL_ACCENT_YELLOW' in locals() else FILL_YELLOW
ws1.cell(row=row_curr, column=2).alignment = Alignment(vertical='center', indent=1)
ws1.row_dimensions[row_curr].height = 24

# Section 2: Calendario Mensual
row_curr += 2
ws1.merge_cells(f'B{row_curr}:K{row_curr}')
ws1.cell(row=row_curr, column=2, value='2. CALENDARIO MENSUAL DE GASTOS NETOS PRESUPUESTADOS (EJERCICIO 2027)').font = FONT_SECTION
ws1.cell(row=row_curr, column=2).fill = FILL_NAVY
ws1.cell(row=row_curr, column=2).alignment = Alignment(vertical='center', indent=1)

row_curr += 1
headers_m = ['Mes', 'Compras (€)', 'Transporte Neto (€)', 'Mantenimiento (€)', 'Marketing (€)', 'Regulatory (€)', 'Total Mensual (€)', '% s/Anual', 'Inmovilizable (€)', 'Gasto Neto P&L (€)']
for col_idx, h in enumerate(headers_m, start=2):
    cell = ws1.cell(row=row_curr, column=col_idx, value=h)
    cell.font = FONT_TH
    cell.fill = FILL_HEADER
    cell.alignment = ALIGN_HEADER
    cell.border = BORDER_HEADER
ws1.row_dimensions[row_curr].height = 24

start_m_row = row_curr + 1
for m_idx in range(12):
    m_name = meses_nombres[m_idx]
    c_val = compras_m[m_idx]['con_res']
    # Transporte Neto = Resto Nacional + Canarias (sin export)
    t_val = trans_2027_m[m_idx]['resto_nal'] + trans_2027_m[m_idx]['canarias']
    m_val = maint_2027_m[m_idx]['con_conting']
    mkt_val = mkt_2027_m[m_idx]['importe']
    reg_val = sum(it['meses'][m_idx] for it in reg_items_27)

    inmov_m = 0.0
    for it in reg_items_27:
        conc_l = it['concepto'].lower()
        if 'regularización nave' in conc_l:
            inmov_m += it['meses'][m_idx]
        elif it['categoria'] == 'marcas' and any(k in conc_l for k in ['guatemala', 'europea', 'arabia', 'omán', 'usa', 'nueva']):
            inmov_m += it['meses'][m_idx]
        elif it['categoria'] == 'Registros' and any(k in conc_l for k in ['azerbaiyán', 'sudáfrica', 'guatemala']):
            inmov_m += it['meses'][m_idx]
    if m_idx == 0: inmov_m += 2110.0
    elif m_idx == 3: inmov_m += 2000.0
    elif m_idx == 8: inmov_m += 2300.0
    elif m_idx == 9: inmov_m += 3000.0

    r_row = start_m_row + m_idx
    fill_m = FILL_ZEBRA if m_idx % 2 == 1 else PatternFill(fill_type=None)

    ws1.cell(row=r_row, column=2, value=m_name).font = FONT_BOLD
    ws1.cell(row=r_row, column=2).alignment = ALIGN_CENTER

    for c_i, val in [(3, c_val), (4, t_val), (5, m_val), (6, mkt_val), (7, reg_val)]:
        c_cell = ws1.cell(row=r_row, column=c_i, value=val)
        c_cell.font = FONT_DATA
        c_cell.number_format = FMT_CURR
        c_cell.alignment = ALIGN_RIGHT

    c_tot = ws1.cell(row=r_row, column=8, value=f'=SUM(C{r_row}:G{r_row})')
    c_tot.font = FONT_BOLD
    c_tot.number_format = FMT_CURR
    c_tot.alignment = ALIGN_RIGHT

    c_pct = ws1.cell(row=r_row, column=9, value=f'=H{r_row}/$H${start_m_row + 12}')
    c_pct.font = FONT_DATA
    c_pct.number_format = '0.0%'
    c_pct.alignment = ALIGN_RIGHT

    c_inm = ws1.cell(row=r_row, column=10, value=inmov_m)
    c_inm.font = FONT_BOLD if inmov_m > 0 else FONT_DATA
    c_inm.number_format = FMT_CURR
    c_inm.alignment = ALIGN_RIGHT
    if inmov_m > 0: c_inm.fill = FILL_GREEN

    c_net = ws1.cell(row=r_row, column=11, value=f'=H{r_row}-J{r_row}')
    c_net.font = FONT_BOLD
    c_net.number_format = FMT_CURR
    c_net.alignment = ALIGN_RIGHT

    for c_col in range(2, 12):
        cell_c = ws1.cell(row=r_row, column=c_col)
        cell_c.border = BORDER_THIN
        if fill_m.fill_type and c_col != 10: cell_c.fill = fill_m

tot_m_row = start_m_row + 12
ws1.cell(row=tot_m_row, column=2, value='TOTAL GASTOS NETOS 2027').font = FONT_TOTAL
ws1.cell(row=tot_m_row, column=2).alignment = ALIGN_CENTER

for c_i, letter in [(3, 'C'), (4, 'D'), (5, 'E'), (6, 'F'), (7, 'G'), (8, 'H'), (10, 'J'), (11, 'K')]:
    cell_tot = ws1.cell(row=tot_m_row, column=c_i, value=f'=SUM({letter}{start_m_row}:{letter}{tot_m_row-1})')
    cell_tot.font = FONT_TOTAL
    cell_tot.number_format = FMT_CURR
    cell_tot.alignment = ALIGN_RIGHT
ws1.cell(row=tot_m_row, column=9, value=1.0).number_format = '100.0%'
ws1.cell(row=tot_m_row, column=9).font = FONT_TOTAL
ws1.cell(row=tot_m_row, column=9).alignment = ALIGN_RIGHT

for c_col in range(2, 12):
    cell_c = ws1.cell(row=tot_m_row, column=c_col)
    cell_c.border = BORDER_TOTAL
    cell_c.fill = FILL_TOTAL

# Section 3: Sep-Dic 2026
row_curr = tot_m_row + 3
ws1.merge_cells(f'B{row_curr}:K{row_curr}')
ws1.cell(row=row_curr, column=2, value='3. CIERRE EJERCICIO 2026 (SEPTIEMBRE - DICIEMBRE): TRANSPORTE Y REGULATORY').font = FONT_SECTION
ws1.cell(row=row_curr, column=2).fill = FILL_NAVY
ws1.cell(row=row_curr, column=2).alignment = Alignment(vertical='center', indent=1)

row_curr += 1
headers_26 = ['Concepto / Partida', 'Cuenta PGC', 'Sep (€)', 'Oct (€)', 'Nov (€)', 'Dic (€)', 'Total Sep-Dic (€)', 'P&L Sep-Dic (€)', 'Diferencia (€)', 'Comentario']
for col_idx, h in enumerate(headers_26, start=2):
    cell = ws1.cell(row=row_curr, column=col_idx, value=h)
    cell.font = FONT_TH
    cell.fill = FILL_SUBHEADER
    cell.alignment = ALIGN_HEADER
    cell.border = BORDER_HEADER
ws1.row_dimensions[row_curr].height = 24

reg_26_summary = [
    ('Transporte Nacional (Campillo)', '624000001', 6434.53, 10499.03, 10212.34, 4053.93, 31199.84, 58042.30, 'Ahorro vs previsión P&L (-26.8k €)'),
    ('Transporte Canarias (Sealine)', '624000003', 2450.11, 7990.59, 11300.38, 534.14, 22275.21, 23316.18, 'Alineado con previsión (-1.0k €)'),
    ('Transporte Exportación (Repercutido a Clientes)', '700001624 / 624', 0.00, 0.00, 0.00, 0.00, 0.00, 0.00, 'Repercutido 100% al cliente (97.236 € facturados; 0 € coste neto Codiagro)'),
    ('Residuos Peligrosos y NP', '629000007', 0.0, 4500.0, 4500.0, 4500.0, 13500.00, 11500.00, 'Estimación de 4.500 €/mes en Q4'),
    ('AEVAE (Declaración y Regulariz.)', '629000002', 0.0, 0.0, 0.0, 2916.0, 2916.00, 2916.00, 'Declaración Q3 + regularización provisional'),
    ('Registros y Tenedurías', '623000003', 0.0, 2650.0, 1600.0, 6150.0, 10400.00, 8000.00, 'Guatemala, Sudáfrica y Arabia'),
    ('Marcas Nuevas (Biorad UE e Int)', '203000000 / 6232', 0.0, 900.0, 0.0, 4000.0, 4900.00, 0.00, 'Biorad UE (4.000 €) activable en inmovilizado'),
    ('Inspecciones y Seguridad Industrial', '622000000', 0.0, 0.0, 1500.0, 0.0, 1500.00, 0.00, 'Inspección Antiincendios y APQ (cada 5 años)'),
    ('Certificaciones Ecológicas (BCS)', '623000004', 0.0, 300.0, 0.0, 1200.0, 1500.00, 5084.56, 'Renovación 4 productos')
]

start_26_row = row_curr + 1
for idx, item in enumerate(reg_26_summary):
    r_idx = start_26_row + idx
    fill_r = FILL_ZEBRA if idx % 2 == 1 else PatternFill(fill_type=None)
    ws1.cell(row=r_idx, column=2, value=item[0]).font = FONT_BOLD
    ws1.cell(row=r_idx, column=3, value=item[1]).font = FONT_DATA
    for c_i, val in enumerate(item[2:7], start=4):
        c_c = ws1.cell(row=r_idx, column=c_i, value=val)
        c_c.font = FONT_DATA
        c_c.number_format = FMT_CURR
        c_c.alignment = ALIGN_RIGHT

    c_pnl = ws1.cell(row=r_idx, column=9, value=item[7])
    c_pnl.font = FONT_DATA
    c_pnl.number_format = FMT_CURR
    c_pnl.alignment = ALIGN_RIGHT

    c_diff = ws1.cell(row=r_idx, column=10, value=f'=H{r_idx}-I{r_idx}')
    c_diff.font = FONT_BOLD
    c_diff.number_format = FMT_CURR
    c_diff.alignment = ALIGN_RIGHT

    ws1.cell(row=r_idx, column=11, value=item[8]).font = FONT_DATA

    for c_col in range(2, 12):
        cell_c = ws1.cell(row=r_idx, column=c_col)
        cell_c.border = BORDER_THIN
        if fill_r.fill_type: cell_c.fill = fill_r

tot_26_row = start_26_row + len(reg_26_summary)
ws1.cell(row=tot_26_row, column=2, value='TOTAL GASTOS Q4 2026').font = FONT_TOTAL
ws1.cell(row=tot_26_row, column=3, value='').font = FONT_TOTAL

for c_i, letter in [(4, 'D'), (5, 'E'), (6, 'F'), (7, 'G'), (8, 'H'), (9, 'I'), (10, 'J')]:
    cell_tot = ws1.cell(row=tot_26_row, column=c_i, value=f'=SUM({letter}{start_26_row}:{letter}{tot_26_row-1})')
    cell_tot.font = FONT_TOTAL
    cell_tot.number_format = FMT_CURR
    cell_tot.alignment = ALIGN_RIGHT
ws1.cell(row=tot_26_row, column=11, value='Control estricto de cierre Q4 (sin distorsión de fletes export)').font = FONT_BOLD

for c_col in range(2, 12):
    cell_c = ws1.cell(row=tot_26_row, column=c_col)
    cell_c.border = BORDER_TOTAL
    cell_c.fill = FILL_TOTAL

# Set column widths
ws1.column_dimensions['A'].width = 3
ws1.column_dimensions['B'].width = 36
ws1.column_dimensions['C'].width = 28
ws1.column_dimensions['D'].width = 18
ws1.column_dimensions['E'].width = 18
ws1.column_dimensions['F'].width = 18
ws1.column_dimensions['G'].width = 16
ws1.column_dimensions['H'].width = 20
ws1.column_dimensions['I'].width = 20
ws1.column_dimensions['J'].width = 20
ws1.column_dimensions['K'].width = 44


# -------------------------------------------------------------
# SHEET 2: DETALLE MENSUAL Y MAPEO PGC
# -------------------------------------------------------------
ws2 = wb.create_sheet(title='Detalle Mensual y Mapeo PGC')
ws2.views.sheetView[0].showGridLines = True

ws2.merge_cells('B2:Y2')
ws2['B2'] = 'DETALLE DE PARTIDAS PRESUPUESTARIAS 2027 Y MAPEO A CUENTAS PGC'
ws2['B2'].font = FONT_TITLE

ws2.merge_cells('B3:Y3')
ws2['B3'] = 'Desglose mensual de cada partida, imputación contable, comparativa con P&L actual y clasificación de inmovilizado'
ws2['B3'].font = FONT_SUBTITLE

headers_d = [
    'Depto', 'Subcategoría / Área', 'Partida / Concepto Detallado', 'Proveedor / Destino',
    'Cta PGC Gasto', 'Nombre Cuenta Contable PGC', 'Tipo Contable', 'Cta Inmovilizado Alternativa',
    'Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic',
    'Total 2027 (€)', 'P&L Actual Cta (€)', 'Diferencia (€)', 'Criterio / Justificación Contable'
]

ws2.row_dimensions[5].height = 28
for col_idx, h in enumerate(headers_d, start=2):
    cell = ws2.cell(row=5, column=col_idx, value=h)
    cell.font = FONT_TH
    cell.fill = FILL_HEADER
    cell.alignment = ALIGN_HEADER
    cell.border = BORDER_HEADER

items_detail = []

# 1. COMPRAS
mp_m = [round(compras_m[i]['con_res'] * 0.7755, 2) for i in range(12)]
env_m = [round(compras_m[i]['con_res'] * 0.1601, 2) for i in range(12)]
etiq_m = [round(compras_m[i]['con_res'] * 0.0250, 2) for i in range(12)]
pal_m = [round(compras_m[i]['con_res'] * 0.0260, 2) for i in range(12)]
tap_m = [round(compras_m[i]['con_res'] * 0.0134, 2) for i in range(12)]

items_detail.append({
    'depto': 'Compras', 'subcat': 'Materias Primas', 'concepto': 'Compras Netas MP Formulados (+3% reserva pedidos mín)', 'prov': 'Varios proveedores MP',
    'cta': '601000000', 'cta_name': 'Compras de materias primas', 'tipo': 'Gasto Corriente', 'cta_inm': '-',
    'meses': mp_m, 'tot': sum(mp_m), 'pnl_cta': 4791964.00,
    'crit': 'Valoración a coste de adquisición neto de stock 31/12 según necesidades de producción formulados.'
})
items_detail.append({
    'depto': 'Compras', 'subcat': 'Envases y Embalajes', 'concepto': 'Envases y bidones para producción envasados', 'prov': 'Varios envases',
    'cta': '602000001', 'cta_name': 'Compra de envases y embalajes', 'tipo': 'Gasto Corriente', 'cta_inm': '-',
    'meses': env_m, 'tot': sum(env_m), 'pnl_cta': 998826.14,
    'crit': 'Coste variable directo de envases según lotes presupuestados.'
})
items_detail.append({
    'depto': 'Compras', 'subcat': 'Etiquetas', 'concepto': 'Etiquetas bobina y adhesivas para envasado', 'prov': 'Imprentas gráficas',
    'cta': '602000002', 'cta_name': 'Compra de etiquetas', 'tipo': 'Gasto Corriente', 'cta_inm': '-',
    'meses': etiq_m, 'tot': sum(etiq_m), 'pnl_cta': 101284.89,
    'crit': 'Material auxiliar de acondicionamiento proporcional a las unidades envasadas.'
})
items_detail.append({
    'depto': 'Compras', 'subcat': 'Palets', 'concepto': 'Palets homologados (nacional y export)', 'prov': 'Proveedores palets',
    'cta': '602000003', 'cta_name': 'Compra de palets', 'tipo': 'Gasto Corriente', 'cta_inm': '-',
    'meses': pal_m, 'tot': sum(pal_m), 'pnl_cta': 105364.66,
    'crit': 'Suministros logísticos de transporte y expedición.'
})
items_detail.append({
    'depto': 'Compras', 'subcat': 'Tapones', 'concepto': 'Tapones de seguridad y dosificadores envases', 'prov': 'Fabricantes tapones',
    'cta': '602000006', 'cta_name': 'Compra tapones de envases', 'tipo': 'Gasto Corriente', 'cta_inm': '-',
    'meses': tap_m, 'tot': sum(tap_m), 'pnl_cta': 39556.10,
    'crit': 'Aprovisionamiento auxiliar directo para líneas de envasado.'
})

# 2. TRANSPORTE
items_detail.append({
    'depto': 'Transporte', 'subcat': 'Logística Nacional', 'concepto': 'Transporte Resto Nacional (Campillo)', 'prov': 'Campillo Palmera',
    'cta': '624000001', 'cta_name': 'Transportes nacional', 'tipo': 'Gasto Corriente', 'cta_inm': '-',
    'meses': [m['resto_nal'] for m in trans_2027_m], 'tot': sum(m['resto_nal'] for m in trans_2027_m), 'pnl_cta': 212857.93,
    'crit': 'Coste variable según expediciones peninsulares estimadas (+10% tarifas e incremento combustible).'
})
items_detail.append({
    'depto': 'Transporte', 'subcat': 'Logística Insular', 'concepto': 'Transporte Canarias (Sealine marítimo)', 'prov': 'Sealine',
    'cta': '624000003', 'cta_name': 'Transportes Canarias', 'tipo': 'Gasto Corriente', 'cta_inm': '-',
    'meses': [m['canarias'] for m in trans_2027_m], 'tot': sum(m['canarias'] for m in trans_2027_m), 'pnl_cta': 88690.80,
    'crit': 'Flete marítimo regular a delegaciones y distribuidores canarios.'
})
# Exportación: REPERCUTIDO AL CLIENTE (Excluido de gasto neto)
items_detail.append({
    'depto': 'Transporte', 'subcat': 'Logística Exportación', 'concepto': 'Transporte Exportación (Repercutido 100% al Cliente)', 'prov': 'Transitarios y navieras',
    'cta': '700001624 / 624', 'cta_name': 'Venta transporte rep. / Flete internacional', 'tipo': 'Repercutido a Clientes (Neutro)', 'cta_inm': '-',
    'meses': [0.0]*12, 'tot': 0.0, 'pnl_cta': 0.0,
    'crit': 'El transporte de exportación (305.223,30 €) se factura íntegramente al cliente vía flete repercutido (700001624). Coste neto 0 € para Codiagro, excluido de gastos.'
})

# 3. MANTENIMIENTO
items_detail.append({
    'depto': 'Mantenimiento', 'subcat': 'Correctivo', 'concepto': 'Imprevistos y averías de planta (3.000 €/mes base + contingencia)', 'prov': 'Varios / Taller propio',
    'cta': '622010000', 'cta_name': 'Mantenimiento preventivo/correctivo', 'tipo': 'Gasto Corriente', 'cta_inm': '-',
    'meses': [round(3000.0 * 1.05, 2) for _ in range(12)], 'tot': 36000.0 * 1.05, 'pnl_cta': 17669.40,
    'crit': 'Mantenimiento correctivo ordinario mensual para resolución de incidencias en fábrica.'
})
m_alcaplant = [0.0]*12
m_alcaplant[3] = 528.0 * 1.05
m_alcaplant[7] = 1000.0 * 1.05
m_alcaplant[8] = 800.0 * 1.05
m_alcaplant[9] = 3000.0 * 1.05
items_detail.append({
    'depto': 'Mantenimiento', 'subcat': 'Alcaplant / Molinos', 'concepto': 'Mantenimiento preventivo y molinos Alcaplant (incluye overhaul Oliver&Batlle)', 'prov': 'Maquisel / Talleres Hurtado / Mavic',
    'cta': '622010200', 'cta_name': 'Mantenimiento Alcaplant', 'tipo': 'Inmovilizable (Parcial)', 'cta_inm': '213000000 - Maquinaria',
    'meses': m_alcaplant, 'tot': sum(m_alcaplant), 'pnl_cta': 80577.24,
    'crit': 'Revisión técnica de molinos. Los 3.000 € de Oliver&Batlle en Octubre son activables en maquinaria si suponen una gran reparación que alarga vida útil.'
})
m_solidos = [round(350.0 * 1.05, 2)]*12
m_solidos[10] += 2500.0 * 1.05
items_detail.append({
    'depto': 'Mantenimiento', 'subcat': 'Envasado Sólido', 'concepto': 'Repuestos envasado sólido y repaso de sinfines Proymec', 'prov': 'Elocom Montajes / Proymec',
    'cta': '622010300', 'cta_name': 'Mantenimiento Sólidos', 'tipo': 'Gasto Corriente', 'cta_inm': '-',
    'meses': m_solidos, 'tot': sum(m_solidos), 'pnl_cta': 71268.96,
    'crit': 'Repuestos periódicos mensuales (350 €) y repaso general anual de sinfines (2.500 € en Noviembre).'
})
m_liq = [0.0]*12
m_liq[0] = 400.0 * 1.05
m_liq[3] = 600.0 * 1.05
m_liq[8] = (600.0 + 800.0 + 2300.0) * 1.05
m_liq[9] = (600.0 + 500.0) * 1.05
items_detail.append({
    'depto': 'Mantenimiento', 'subcat': 'Líquidos y Tuberías', 'concepto': 'Valvulería, latiguillos, fontanería y propulsores Proymec', 'prov': 'S.I.R / Irriagro / Proymec',
    'cta': '622010400', 'cta_name': 'Mantenimiento Líquidos', 'tipo': 'Inmovilizable (Parcial)', 'cta_inm': '213000000 - Maquinaria',
    'meses': m_liq, 'tot': sum(m_liq), 'pnl_cta': 12448.08,
    'crit': 'Mantenimiento de reactores y líneas de trasvase. La sustitución integral de válvulas propulsores (2.300 € en Sept) puede capitalizarse.'
})
m_rep = [round(250.0 * 1.05, 2)]*12
m_rep[8] = round(350.0 * 1.05, 2)
items_detail.append({
    'depto': 'Mantenimiento', 'subcat': 'Taller Planta', 'concepto': 'Consumibles de taller y repuestos generales de mantenimiento', 'prov': 'Novastec / Thermocontrol',
    'cta': '622010500', 'cta_name': 'Repuestos mantenimiento varios', 'tipo': 'Gasto Corriente', 'cta_inm': '-',
    'meses': m_rep, 'tot': sum(m_rep), 'pnl_cta': 9458.40,
    'crit': 'Suministros de ferretería, tornillería, engrase y consumibles operativos de taller.'
})
m_inst = [0.0]*12
m_inst[0] = 2110.0 * 1.05
m_inst[7] = (1800.0 + 250.0 + 211.75) * 1.05
items_detail.append({
    'depto': 'Mantenimiento', 'subcat': 'Instalaciones y Edificios', 'concepto': 'Reparación cubierta tejado nave y puerta corredera', 'prov': 'Cuvierbal / Diego Rodrigo Dols',
    'cta': '622000000', 'cta_name': 'Reparaciones y conservación', 'tipo': 'Inmovilizable (Parcial)', 'cta_inm': '211000000 - Construcciones',
    'meses': m_inst, 'tot': sum(m_inst), 'pnl_cta': 17669.40,
    'crit': 'Reparación de tejado (2.110 €) y adecuación nave: activable como mayor valor de Construcciones (211) si supone sustitución de cubierta.'
})
m_tool = [0.0]*12
m_tool[3] = 2000.0 * 1.05
items_detail.append({
    'depto': 'Mantenimiento', 'subcat': 'Equipamiento Taller', 'concepto': 'Herramientas eléctricas de taller', 'prov': 'Recambios Alcora',
    'cta': '629000016', 'cta_name': 'Herramientas taller', 'tipo': 'Inmovilizable (Activable)', 'cta_inm': '214000000 - Utillaje',
    'meses': m_tool, 'tot': sum(m_tool), 'pnl_cta': 6000.36,
    'crit': 'Herramientas duraderas de taller con vida útil plurianual; activable en la cuenta 214 Utillaje según PGC.'
})
m_ext = [0.0]*12
m_ext[8] = (525.0 + 1693.0) * 1.05
m_ext[9] = 1693.0 * 1.05
m_ext[11] = 525.0 * 1.05
items_detail.append({
    'depto': 'Mantenimiento', 'subcat': 'Servicios Externos Especializados', 'concepto': 'Revisión compresores de aire y mantenimiento carretillas elevadoras', 'prov': 'Repcar / Especialista aire comprimido',
    'cta': '622000000', 'cta_name': 'Reparaciones y conservación', 'tipo': 'Gasto Corriente', 'cta_inm': '-',
    'meses': m_ext, 'tot': sum(m_ext), 'pnl_cta': 17669.40,
    'crit': 'Mantenimiento preventivo reglamentario y normativo de compresores y maquinaria de elevación.'
})

# 4. MARKETING
items_detail.append({
    'depto': 'Marketing', 'subcat': 'Ferias y Encuentros', 'concepto': 'Ferias internacionales (Fruit Attraction, Fruit Logistica, Sahara Expo, Growtech, etc.)', 'prov': 'Entidades feriales e instaladores',
    'cta': '627000007', 'cta_name': 'Ferias y congresos', 'tipo': 'Gasto Corriente', 'cta_inm': '-',
    'meses': [round(33770.0 * (mkt_2027_m[i]['importe'] / tot_mkt_27), 2) for i in range(12)],
    'tot': 33770.00, 'pnl_cta': 3855.75,
    'crit': 'Participación ferial, stands, acreditaciones y eventos comerciales internacionales.'
})
items_detail.append({
    'depto': 'Marketing', 'subcat': 'Corporativo y Marca', 'concepto': 'Marketing transversal, diseño corporativo, catálogos, packaging y web', 'prov': 'Agencias diseño y marketing',
    'cta': '627000006', 'cta_name': 'Publicidad, catálogos y folletos', 'tipo': 'Gasto Corriente', 'cta_inm': '-',
    'meses': [round(25630.0 * (mkt_2027_m[i]['importe'] / tot_mkt_27), 2) for i in range(12)],
    'tot': 25630.00, 'pnl_cta': 1236.75,
    'crit': 'Renovación de imagen de producto, fichas técnicas multilingües y soporte institucional.'
})
items_detail.append({
    'depto': 'Marketing', 'subcat': 'Programas Cliente-Producto', 'concepto': 'Planes de activación comercial por cliente clave y producto foco (Familias A/B)', 'prov': 'Distribuidores y material promocional',
    'cta': '627000008', 'cta_name': 'Marketing y campañas comerciales', 'tipo': 'Gasto Corriente', 'cta_inm': '-',
    'meses': [round(196300.0 * (mkt_2027_m[i]['importe'] / tot_mkt_27), 2) for i in range(12)],
    'tot': 196300.00, 'pnl_cta': 0.00,
    'crit': 'Plan específico por clientes estratégicos y productos prioritarios en fase de lanzamiento y crecimiento.'
})
items_detail.append({
    'depto': 'Marketing', 'subcat': 'Reserva Táctica', 'concepto': 'Fondo de maniobra para contingencias y oportunidades comerciales imprevistas', 'prov': 'Fondo de reserva comercial',
    'cta': '627000009', 'cta_name': 'Artículos publicidad / promociones', 'tipo': 'Gasto Corriente', 'cta_inm': '-',
    'meses': [round(10372.40 * (mkt_2027_m[i]['importe'] / tot_mkt_27), 2) for i in range(12)],
    'tot': 10372.40, 'pnl_cta': 26500.00,
    'crit': 'Reserva táctica para promociones de temporada y apoyo a comerciales de zona.'
})

# 5. REGULATORY
for it in reg_items_27:
    cat = it['categoria']
    conc = it['concepto']
    tot = it['total']
    months = it['meses']
    note = it['nota']

    cta = '629000000'
    cta_name = 'Otros servicios'
    tipo = 'Gasto Corriente'
    cta_inm = '-'
    pnl_cta = 0.0

    conc_l = conc.lower()
    if cat == 'medioambiente':
        if 'residuos' in conc_l:
            cta = '629000007'
            cta_name = 'Gastos retirada de residuos'
            pnl_cta = 34000.00
            crit = 'Retirada y tratamiento legal obligatorio de residuos tóxicos, peligrosos y no peligrosos.'
        elif 'regularización nave' in conc_l:
            cta = '622000000'
            cta_name = 'Reparaciones y conservación'
            tipo = 'Inmovilizable (Activable)'
            cta_inm = '211000000 - Construcciones'
            pnl_cta = 0.00
            crit = 'Obra finalizada de acondicionamiento de nave enfrente. Activable como Inmovilizado Material (211).'
    elif cat == 'envases':
        cta = '629000002'
        cta_name = 'Otros gastos de gestión / Envases'
        pnl_cta = 11379.82
        crit = 'Declaración anual y cuotas SCRAP AEVAE por adhesión a sistema de recogida de envases agrícolas.'
    elif cat == 'marcas':
        pnl_cta = 0.00
        if any(k in conc_l for k in ['codiorgan-guatemala', 'trifeno-94', 'arabia', 'sprintium omán', 'sprintium usa', 'codiagro usa']) or 'nueva' in note.lower() or 'extensión' in note.lower():
            cta = '623200000'
            cta_name = 'Gastos marcas y patentes'
            tipo = 'Inmovilizable (Activable)'
            cta_inm = '203000000 - Propiedad Industrial'
            crit = 'Registro de nueva marca con concesión de monopolio legal plurianual. Activable como Inmovilizado Intangible (203).'
        elif 'vigilancia' in conc_l:
            cta = '623200033'
            cta_name = 'Gastos marcas mundial'
            crit = 'Vigilancia de marca y marcas colindantes para prevención de infracciones de propiedad intelectual.'
        elif 'titular' in conc_l:
            cta = '623200000'
            cta_name = 'Gastos marcas y patentes'
            crit = 'Trámites notariales y registrales de cambio de titularidad de expedientes por reestructuración.'
        else:
            cta = '623200000'
            cta_name = 'Gastos marcas y patentes'
            crit = 'Tasas oficiales de renovación decenal de marcas comerciales registradas en el país de destino.'
    elif cat == 'Registros':
        if any(k in conc_l for k in ['azerbaiyán', 'sudáfrica', 'guatemala']):
            cta = '623000003'
            cta_name = 'Servicios registros y notarios'
            tipo = 'Inmovilizable (Activable)'
            cta_inm = '203000000 - Licencias y Registros'
            pnl_cta = 24000.00
            crit = 'Coste del expediente de registro fitosanitario/nutricional oficial con autorización plurianual en mercado exterior.'
        elif 'labo' in conc_l:
            cta = '629000005'
            cta_name = 'Gastos análisis de laboratorios'
            pnl_cta = 11295.24
            crit = 'Ensayos analíticos y de eficacia en laboratorios externos acreditados para expedientes de registro.'
        else:
            cta = '623000003'
            cta_name = 'Servicios registros y notarios'
            pnl_cta = 24000.00
            crit = 'Trámites de registro y tasas administrativas de autorización fitonutricional.'
    elif cat == 'ecológicos':
        cta = '623000004'
        cta_name = 'Gastos certificaciones ecológicas'
        pnl_cta = 15253.68
        crit = 'Auditorías de certificación anual para etiquetado ecológico oficial (BCS-Kiwa, OMRI, FIBL, Ecocert).'
    elif cat == 'calidad' or cat == 'FDS' or cat == 'consejero seguridad' or cat == 'Seguridad industrial':
        cta = '623000001'
        cta_name = 'Servicios profesionales independientes'
        pnl_cta = 74000.00
        crit = 'Servicios técnicos externos de seguridad química, auditoría ISO y fichas de seguridad.'
    elif cat == 'general':
        cta = '629000008'
        cta_name = 'Cuotas asociaciones'
        pnl_cta = 9621.84
        crit = 'Cuota gremial y patronal del sector químico (Quimacova).'
    else:
        cta = '629000002'
        cta_name = 'Otros gastos'
        pnl_cta = 11379.82
        crit = 'Partida menor de gestión regulatoria.'

    items_detail.append({
        'depto': 'Regulatory', 'subcat': cat.capitalize(), 'concepto': conc, 'prov': note if note else '-',
        'cta': cta, 'cta_name': cta_name, 'tipo': tipo, 'cta_inm': cta_inm,
        'meses': months, 'tot': tot, 'pnl_cta': pnl_cta, 'crit': crit
    })

# Write items to Sheet 2
start_d_row = 6
for idx, it in enumerate(items_detail):
    r_idx = start_d_row + idx
    fill_row = FILL_ZEBRA if idx % 2 == 1 else PatternFill(fill_type=None)

    ws2.cell(row=r_idx, column=2, value=it['depto']).font = FONT_BOLD
    ws2.cell(row=r_idx, column=3, value=it['subcat']).font = FONT_DATA
    ws2.cell(row=r_idx, column=4, value=it['concepto']).font = FONT_BOLD if it['tipo'] != 'Gasto Corriente' else FONT_DATA
    ws2.cell(row=r_idx, column=5, value=str(it['prov'])[:30]).font = FONT_DATA
    ws2.cell(row=r_idx, column=6, value=it['cta']).font = FONT_BOLD
    ws2.cell(row=r_idx, column=7, value=it['cta_name']).font = FONT_DATA
    
    c_tipo = ws2.cell(row=r_idx, column=8, value=it['tipo'])
    c_tipo.font = FONT_BOLD if it['tipo'] != 'Gasto Corriente' else FONT_DATA
    if 'Inmovilizable' in it['tipo']:
        c_tipo.fill = FILL_GREEN
        c_tipo.font = FONT_SUCCESS
    elif 'Repercutido' in it['tipo']:
        c_tipo.fill = FILL_ORANGE
        c_tipo.font = FONT_ALERT

    ws2.cell(row=r_idx, column=9, value=it['cta_inm']).font = FONT_BOLD if it['cta_inm'] != '-' else FONT_DATA

    # Months 10 to 21
    for m_i in range(12):
        val = it['meses'][m_i]
        c_m = ws2.cell(row=r_idx, column=10 + m_i, value=val)
        c_m.font = FONT_DATA
        c_m.number_format = FMT_CURR
        c_m.alignment = ALIGN_RIGHT

    # Total 2027 formula
    c_tot = ws2.cell(row=r_idx, column=22, value=f'=SUM(J{r_idx}:U{r_idx})')
    c_tot.font = FONT_BOLD
    c_tot.number_format = FMT_CURR
    c_tot.alignment = ALIGN_RIGHT

    # P&L Actual
    c_pnl = ws2.cell(row=r_idx, column=23, value=it['pnl_cta'])
    c_pnl.font = FONT_DATA
    c_pnl.number_format = FMT_CURR
    c_pnl.alignment = ALIGN_RIGHT

    # Difference formula
    c_diff = ws2.cell(row=r_idx, column=24, value=f'=V{r_idx}-W{r_idx}')
    c_diff.font = FONT_BOLD
    c_diff.number_format = FMT_CURR
    c_diff.alignment = ALIGN_RIGHT

    # Justification
    c_crit = ws2.cell(row=r_idx, column=25, value=it['crit'])
    c_crit.font = FONT_DATA
    c_crit.alignment = ALIGN_LEFT

    for c_col in range(2, 26):
        cell_c = ws2.cell(row=r_idx, column=c_col)
        cell_c.border = BORDER_THIN
        if fill_row.fill_type and c_col not in [8]:
            cell_c.fill = fill_row

# Total Row Sheet 2
tot_d_row = start_d_row + len(items_detail)
ws2.cell(row=tot_d_row, column=2, value='TOTAL GASTOS COMPUTABLES CODIAGRO').font = FONT_TOTAL
for c_col in range(3, 10):
    ws2.cell(row=tot_d_row, column=c_col, value='').font = FONT_TOTAL

for m_i in range(12):
    col_letter = get_column_letter(10 + m_i)
    c_tot = ws2.cell(row=tot_d_row, column=10 + m_i, value=f'=SUM({col_letter}{start_d_row}:{col_letter}{tot_d_row-1})')
    c_tot.font = FONT_TOTAL
    c_tot.number_format = FMT_CURR
    c_tot.alignment = ALIGN_RIGHT

# Total annual formula
c_tot_anual = ws2.cell(row=tot_d_row, column=22, value=f'=SUM(V{start_d_row}:V{tot_d_row-1})')
c_tot_anual.font = FONT_TOTAL
c_tot_anual.number_format = FMT_CURR
c_tot_anual.alignment = ALIGN_RIGHT

ws2.cell(row=tot_d_row, column=23, value=f'=SUM(W{start_d_row}:W{tot_d_row-1})').number_format = FMT_CURR
ws2.cell(row=tot_d_row, column=23).font = FONT_TOTAL
ws2.cell(row=tot_d_row, column=23).alignment = ALIGN_RIGHT

ws2.cell(row=tot_d_row, column=24, value=f'=V{tot_d_row}-W{tot_d_row}').number_format = FMT_CURR
ws2.cell(row=tot_d_row, column=24).font = FONT_TOTAL
ws2.cell(row=tot_d_row, column=24).alignment = ALIGN_RIGHT

ws2.cell(row=tot_d_row, column=25, value='Control consolidado de detalle (excluye partidas repercutidas)').font = FONT_BOLD

for c_col in range(2, 26):
    cell_c = ws2.cell(row=tot_d_row, column=c_col)
    cell_c.border = BORDER_TOTAL
    cell_c.fill = FILL_TOTAL

# Set column widths Sheet 2
col_widths_s2 = {
    'A': 3, 'B': 16, 'C': 22, 'D': 35, 'E': 24, 'F': 14, 'G': 28, 'H': 24, 'I': 24,
    'J': 14, 'K': 14, 'L': 14, 'M': 14, 'N': 14, 'O': 14, 'P': 14, 'Q': 14, 'R': 14, 'S': 14, 'T': 14, 'U': 14,
    'V': 18, 'W': 18, 'X': 18, 'Y': 55
}
for col_l, w in col_widths_s2.items():
    ws2.column_dimensions[col_l].width = w


# -------------------------------------------------------------
# SHEET 3: ANÁLISIS DE INMOVILIZACIONES (CAPITALIZACIONES)
# -------------------------------------------------------------
ws3 = wb.create_sheet(title='Análisis Inmovilizaciones')
ws3.views.sheetView[0].showGridLines = True

ws3.merge_cells('B2:K2')
ws3['B2'] = 'ANÁLISIS DE PARTIDAS SUSCEPTIBLES DE INMOVILIZACIÓN / ACTIVACIÓN'
ws3['B2'].font = FONT_TITLE

ws3.merge_cells('B3:K3')
ws3['B3'] = 'Criterios PGC para activar gastos como Inmovilizado Material e Intangible y su impacto directo en EBITDA y Balance'
ws3['B3'].font = FONT_SUBTITLE

ws3.merge_cells('B5:K5')
ws3['B5'] = '1. DETALLE DE PARTIDAS ACTIVABLES SEGÚN EL PLAN GENERAL CONTABLE (PGC)'
ws3['B5'].font = FONT_SECTION
ws3['B5'].fill = FILL_NAVY
ws3['B5'].alignment = Alignment(vertical='center', indent=1)

headers_inm = [
    'Área / Origen', 'Concepto / Partida', 'Importe (€)', 'Cta Gasto Inicial',
    'Cta Inmovilizado Destino', 'Subgrupo Balance', 'Norma PGC Aplicable',
    'Justificación de Activación', 'Años Amortiz.', 'Cuota Anual Amortiz. (€)', 'Mejora EBITDA (€)'
]
ws3.row_dimensions[6].height = 28
for col_idx, h in enumerate(headers_inm, start=2):
    cell = ws3.cell(row=6, column=col_idx, value=h)
    cell.font = FONT_TH
    cell.fill = FILL_HEADER
    cell.alignment = ALIGN_HEADER
    cell.border = BORDER_HEADER

capitalizable_items = [
    {
        'area': 'Regulatory - Marcas',
        'conc': 'Codiorgan - Guatemala (Registro nueva marca)',
        'imp': 1750.00,
        'cta_gasto': '623200021',
        'cta_inm': '203000000 - Propiedad Industrial',
        'subg': 'Inmovilizado Intangible',
        'norma': 'NRV 5ª y 6ª PGC (Propiedad Industrial)',
        'just': 'Otorga derecho de exclusividad comercial plurianual en Guatemala; generador directo de flujos de caja futuros.',
        'years': 10,
        'cuota': 175.00
    },
    {
        'area': 'Regulatory - Marcas',
        'conc': 'Trifeno-94 - Registro Marca Europea (UE)',
        'imp': 1500.00,
        'cta_gasto': '623200006',
        'cta_inm': '203000000 - Propiedad Industrial',
        'subg': 'Inmovilizado Intangible',
        'norma': 'NRV 5ª y 6ª PGC (Propiedad Industrial)',
        'just': 'Registro de activo intangible comunitario en EUIPO con protección legal y comercial por 10 años.',
        'years': 10,
        'cuota': 150.00
    },
    {
        'area': 'Regulatory - Marcas',
        'conc': 'Trifeno-94 - Registro Marca Guatemala',
        'imp': 1750.00,
        'cta_gasto': '623200021',
        'cta_inm': '203000000 - Propiedad Industrial',
        'subg': 'Inmovilizado Intangible',
        'norma': 'NRV 5ª y 6ª PGC (Propiedad Industrial)',
        'just': 'Activo intangible protegido en mercado centroamericano en expansión.',
        'years': 10,
        'cuota': 175.00
    },
    {
        'area': 'Regulatory - Marcas',
        'conc': 'Salwax, Biorad y Floramec Avance - Arabia Saudí (3 marcas)',
        'imp': 6000.00,
        'cta_gasto': '623200000',
        'cta_inm': '203000000 - Propiedad Industrial',
        'subg': 'Inmovilizado Intangible',
        'norma': 'NRV 5ª y 6ª PGC (Propiedad Industrial)',
        'just': 'Registro de 3 marcas nuevas en Oriente Medio como requisito previo para la distribución exclusiva en Arabia.',
        'years': 10,
        'cuota': 600.00
    },
    {
        'area': 'Regulatory - Marcas',
        'conc': 'Sprintium Omán (Extensión territorial de marca)',
        'imp': 1700.00,
        'cta_gasto': '623200000',
        'cta_inm': '203000000 - Propiedad Industrial',
        'subg': 'Inmovilizado Intangible',
        'norma': 'NRV 5ª y 6ª PGC (Propiedad Industrial)',
        'just': 'Extensión de derechos de marca con vigencia legal plurianual en mercado de Omán.',
        'years': 10,
        'cuota': 170.00
    },
    {
        'area': 'Regulatory - Marcas',
        'conc': 'Sprintium y Codiagro - Estados Unidos (Trámites concesión)',
        'imp': 2600.00,
        'cta_gasto': '623200037',
        'cta_inm': '203000000 - Propiedad Industrial',
        'subg': 'Inmovilizado Intangible',
        'norma': 'NRV 5ª y 6ª PGC (Propiedad Industrial)',
        'just': 'Costes incurridos para el registro oficial de marcas ante la USPTO.',
        'years': 10,
        'cuota': 260.00
    },
    {
        'area': 'Regulatory - Registros',
        'conc': 'Registro Azerbaiyán (6 formulaciones agronómicas)',
        'imp': 13200.00,
        'cta_gasto': '623000003',
        'cta_inm': '203000000 - Concesiones y Registros',
        'subg': 'Inmovilizado Intangible',
        'norma': 'NRV 5ª PGC / Res. ICAC 28-05-2013',
        'just': 'Autorización administrativa oficial imprescindible para vender legalmente en Azerbaiyán por periodo plurianual.',
        'years': 5,
        'cuota': 2640.00
    },
    {
        'area': 'Regulatory - Registros',
        'conc': 'Registro Sudáfrica (8 formulaciones agronómicas)',
        'imp': 6960.00,
        'cta_gasto': '623000003',
        'cta_inm': '203000000 - Concesiones y Registros',
        'subg': 'Inmovilizado Intangible',
        'norma': 'NRV 5ª PGC / Res. ICAC 28-05-2013',
        'just': 'Autorizaciones oficiales de fertilizantes en el Departamento de Agricultura sudafricano con validez plurianual.',
        'years': 5,
        'cuota': 1392.00
    },
    {
        'area': 'Regulatory - Registros',
        'conc': 'Registro Guatemala (Teneduría y registro de catálogo)',
        'imp': 5740.00,
        'cta_gasto': '623000003',
        'cta_inm': '203000000 - Concesiones y Registros',
        'subg': 'Inmovilizado Intangible',
        'norma': 'NRV 5ª PGC / Res. ICAC 28-05-2013',
        'just': 'Teneduría de registros y homologación ministerial de productos.',
        'years': 5,
        'cuota': 1148.00
    },
    {
        'area': 'Regulatory - Medioambiente',
        'conc': 'Regularización y adecuación nave enfrente (Obra finalizada)',
        'imp': 9600.00,
        'cta_gasto': '622000000',
        'cta_inm': '211000000 - Construcciones / Instalaciones',
        'subg': 'Inmovilizado Material',
        'norma': 'NRV 2ª y 3ª PGC (Inmovilizado Material)',
        'just': 'Inversión en obra de nave que genera una infraestructura operativa duradera y prolonga la capacidad productiva/almacén.',
        'years': 20,
        'cuota': 480.00
    },
    {
        'area': 'Mantenimiento - Fábrica',
        'conc': 'Reparación y sustitución de cubierta tejado nave (Cuvierbal)',
        'imp': 2110.00,
        'cta_gasto': '622000000',
        'cta_inm': '211000000 - Construcciones',
        'subg': 'Inmovilizado Material',
        'norma': 'Res. ICAC 01-03-2019 (Grandes reparaciones)',
        'just': 'Sustitución estructural de elementos de cubierta que amplía la vida útil del edificio fabril frente a simple parcheo.',
        'years': 20,
        'cuota': 105.50
    },
    {
        'area': 'Mantenimiento - Fábrica',
        'conc': 'Revisión y reconstrucción mayor Molinos Oliver & Batlle (Maquisel)',
        'imp': 3000.00,
        'cta_gasto': '622010200',
        'cta_inm': '213000000 - Maquinaria',
        'subg': 'Inmovilizado Material',
        'norma': 'Res. ICAC 01-03-2019 (Mejora y overhaul)',
        'just': 'Overhaul integral de 3 molinos EHP50 que restablece y prolonga la capacidad fabril y vida útil del activo clave.',
        'years': 8,
        'cuota': 375.00
    },
    {
        'area': 'Mantenimiento - Fábrica',
        'conc': 'Sustitución y montaje válvulas propulsores (Proymec)',
        'imp': 2300.00,
        'cta_gasto': '622010400',
        'cta_inm': '213000000 - Maquinaria',
        'subg': 'Inmovilizado Material',
        'norma': 'NRV 2ª PGC (Sustitución de componentes)',
        'just': 'Sustitución de componentes motrices con vida superior al ejercicio contable.',
        'years': 5,
        'cuota': 460.00
    },
    {
        'area': 'Mantenimiento - Taller',
        'conc': 'Herramientas eléctricas y utillaje de taller (Recambios Alcora)',
        'imp': 2000.00,
        'cta_gasto': '629000016',
        'cta_inm': '214000000 - Utillaje',
        'subg': 'Inmovilizado Material',
        'norma': 'NRV 2ª y 3ª PGC (Utillaje y herramientas)',
        'just': 'Equipamiento de taller de uso continuado y vida útil superior a un año.',
        'years': 4,
        'cuota': 500.00
    }
]

start_inm_row = 7
for idx, it in enumerate(capitalizable_items):
    r_idx = start_inm_row + idx
    fill_row = FILL_ZEBRA if idx % 2 == 1 else PatternFill(fill_type=None)

    ws3.cell(row=r_idx, column=2, value=it['area']).font = FONT_BOLD
    ws3.cell(row=r_idx, column=3, value=it['conc']).font = FONT_DATA
    
    c_imp = ws3.cell(row=r_idx, column=4, value=it['imp'])
    c_imp.font = FONT_BOLD
    c_imp.number_format = FMT_CURR
    c_imp.alignment = ALIGN_RIGHT

    ws3.cell(row=r_idx, column=5, value=it['cta_gasto']).font = FONT_DATA
    ws3.cell(row=r_idx, column=6, value=it['cta_inm']).font = FONT_BOLD
    ws3.cell(row=r_idx, column=7, value=it['subg']).font = FONT_DATA
    ws3.cell(row=r_idx, column=8, value=it['norma']).font = FONT_DATA
    ws3.cell(row=r_idx, column=9, value=it['just']).font = FONT_DATA
    
    c_yr = ws3.cell(row=r_idx, column=10, value=it['years'])
    c_yr.font = FONT_DATA
    c_yr.alignment = ALIGN_CENTER

    c_cuota = ws3.cell(row=r_idx, column=11, value=it['cuota'])
    c_cuota.font = FONT_DATA
    c_cuota.number_format = FMT_CURR
    c_cuota.alignment = ALIGN_RIGHT

    c_mejora = ws3.cell(row=r_idx, column=12, value=f'=D{r_idx}-K{r_idx}')
    c_mejora.font = FONT_BOLD
    c_mejora.number_format = FMT_CURR
    c_mejora.alignment = ALIGN_RIGHT
    c_mejora.fill = FILL_GREEN

    for c_col in range(2, 13):
        cell_c = ws3.cell(row=r_idx, column=c_col)
        cell_c.border = BORDER_THIN
        if fill_row.fill_type and c_col != 12:
            cell_c.fill = fill_row

# Total Inmovilizaciones
tot_inm_row = start_inm_row + len(capitalizable_items)
ws3.cell(row=tot_inm_row, column=2, value='TOTAL ACTIVABLE').font = FONT_TOTAL
ws3.cell(row=tot_inm_row, column=3, value='').font = FONT_TOTAL

cell_tot_imp = ws3.cell(row=tot_inm_row, column=4, value=f'=SUM(D{start_inm_row}:D{tot_inm_row-1})')
cell_tot_imp.font = FONT_TOTAL
cell_tot_imp.number_format = FMT_CURR
cell_tot_imp.alignment = ALIGN_RIGHT

for c_c in range(5, 11):
    ws3.cell(row=tot_inm_row, column=c_c, value='').font = FONT_TOTAL

cell_tot_cuota = ws3.cell(row=tot_inm_row, column=11, value=f'=SUM(K{start_inm_row}:K{tot_inm_row-1})')
cell_tot_cuota.font = FONT_TOTAL
cell_tot_cuota.number_format = FMT_CURR
cell_tot_cuota.alignment = ALIGN_RIGHT

cell_tot_mejora = ws3.cell(row=tot_inm_row, column=12, value=f'=SUM(L{start_inm_row}:L{tot_inm_row-1})')
cell_tot_mejora.font = FONT_TOTAL
cell_tot_mejora.number_format = FMT_CURR
cell_tot_mejora.alignment = ALIGN_RIGHT
cell_tot_mejora.fill = FILL_GREEN

for c_col in range(2, 13):
    cell_c = ws3.cell(row=tot_inm_row, column=c_col)
    cell_c.border = BORDER_TOTAL
    if c_col != 12: cell_c.fill = FILL_TOTAL

# Section 2: Cuadro de Impacto Financiero
row_curr = tot_inm_row + 3
ws3.merge_cells(f'B{row_curr}:L{row_curr}')
ws3.cell(row=row_curr, column=2, value='2. IMPACTO COMPARATIVO EN EL RESULTADO (P&L): GASTO DIRECTO VS ACTIVACIÓN').font = FONT_SECTION
ws3.cell(row=row_curr, column=2).fill = FILL_NAVY
ws3.cell(row=row_curr, column=2).alignment = Alignment(vertical='center', indent=1)

row_curr += 1
headers_imp = ['Escenario Contable', 'Importe Incurrido (€)', 'Gasto en EBITDA (€)', 'Amortización PGC (€)', 'Impacto Neto en EBIT (€)', 'Impacto en Balance (€)', 'Ventaja Financiera y Dictamen']
for col_idx, h in enumerate(headers_imp, start=2):
    cell = ws3.cell(row=row_curr, column=col_idx, value=h)
    cell.font = FONT_TH
    cell.fill = FILL_HEADER
    cell.alignment = ALIGN_HEADER
    cell.border = BORDER_HEADER
ws3.row_dimensions[row_curr].height = 24

row_curr += 1
ws3.cell(row=row_curr, column=2, value='Escenario A: Imputación a Gasto Directo (Cuentas 622, 623, 629)').font = FONT_BOLD
ws3.cell(row=row_curr, column=3, value=f'=D{tot_inm_row}').number_format = FMT_CURR
ws3.cell(row=row_curr, column=4, value=f'=-D{tot_inm_row}').number_format = FMT_CURR
ws3.cell(row=row_curr, column=4).font = FONT_ALERT
ws3.cell(row=row_curr, column=5, value=0.0).number_format = FMT_CURR
ws3.cell(row=row_curr, column=6, value=f'=-D{tot_inm_row}').number_format = FMT_CURR
ws3.cell(row=row_curr, column=6).font = FONT_ALERT
ws3.cell(row=row_curr, column=7, value=0.0).number_format = FMT_CURR
ws3.cell(row=row_curr, column=8, value='Castiga el EBITDA del ejercicio 2027 en -61.210 € sin reflejar el valor patrimonial creado.').font = FONT_DATA
for c_col in range(2, 9):
    ws3.cell(row=row_curr, column=c_col).border = BORDER_THIN

row_curr += 1
ws3.cell(row=row_curr, column=2, value='Escenario B: Activación como Inmovilizado (Cuentas 203, 211, 213, 214)').font = FONT_BOLD
ws3.cell(row=row_curr, column=3, value=f'=D{tot_inm_row}').number_format = FMT_CURR
ws3.cell(row=row_curr, column=4, value=0.0).number_format = FMT_CURR
ws3.cell(row=row_curr, column=4).font = FONT_SUCCESS
ws3.cell(row=row_curr, column=5, value=f'=-K{tot_inm_row}').number_format = FMT_CURR
ws3.cell(row=row_curr, column=6, value=f'=-K{tot_inm_row}').number_format = FMT_CURR
ws3.cell(row=row_curr, column=7, value=f'=D{tot_inm_row}-K{tot_inm_row}').number_format = FMT_CURR
ws3.cell(row=row_curr, column=7).font = FONT_SUCCESS
ws3.cell(row=row_curr, column=8, value='Mejora el EBITDA en +61.210 € y el EBIT en +52.360 €. Refleja activos reales en el balance.').font = FONT_BOLD
for c_col in range(2, 9):
    c_b = ws3.cell(row=row_curr, column=c_col)
    c_b.border = BORDER_THIN
    c_b.fill = FILL_GREEN

col_widths_s3 = {
    'A': 3, 'B': 24, 'C': 38, 'D': 16, 'E': 16, 'F': 26, 'G': 22,
    'H': 26, 'I': 45, 'J': 14, 'K': 18, 'L': 18
}
for col_l, w in col_widths_s3.items():
    ws3.column_dimensions[col_l].width = w


# -------------------------------------------------------------
# SHEET 4: COMPARATIVA CUENTAS PGC
# -------------------------------------------------------------
ws4 = wb.create_sheet(title='Comparativa Cuentas PGC')
ws4.views.sheetView[0].showGridLines = True

ws4.merge_cells('B2:H2')
ws4['B2'] = 'COMPARATIVA CONSOLIDADA POR CUENTA CONTABLE PGC (2027)'
ws4['B2'].font = FONT_TITLE

ws4.merge_cells('B3:H3')
ws4['B3'] = 'Alineación entre el Presupuesto P&L actual y las partidas asumidas por Codiagro (excluyendo fletes repercutidos)'
ws4['B3'].font = FONT_SUBTITLE

headers_c = [
    'Cuenta PGC', 'Nombre Cuenta Contable', 'Área / Departamento',
    'P&L Actual 2027 (€)', 'Presupuesto Depto 2027 (€)', 'Desviación (€)', 'Desviación (%)', 'Explicación del Ajuste'
]
ws4.row_dimensions[5].height = 28
for col_idx, h in enumerate(headers_c, start=2):
    cell = ws4.cell(row=5, column=col_idx, value=h)
    cell.font = FONT_TH
    cell.fill = FILL_HEADER
    cell.alignment = ALIGN_HEADER
    cell.border = BORDER_HEADER

accounts_summary = [
    ('601000000', 'Compras de materias primas', 'Compras', 4791964.00, sum(mp_m), 'Ajuste fino de necesidades netas de producción (+1,2%)'),
    ('602000001', 'Compra de envases y embalajes', 'Compras', 998826.14, sum(env_m), 'Consistente con volumen de envasado'),
    ('602000002', 'Compra de etiquetas', 'Compras', 101284.89, sum(etiq_m), 'Ajuste por tarifas actualizadas de imprenta'),
    ('602000003', 'Compra de palets', 'Compras', 105364.66, sum(pal_m), 'Incremento acorde a rotación y paletización'),
    ('602000006', 'Compra tapones de envases', 'Compras', 39556.10, sum(tap_m), 'Proporcional a unidades envasadas'),
    ('624000001', 'Transportes nacional', 'Transporte', 212857.93, sum(m['resto_nal'] for m in trans_2027_m), 'Ahorro logístico nacional (-46.4k €) por optimización de rutas'),
    ('624000003', 'Transportes Canarias', 'Transporte', 88690.80, sum(m['canarias'] for m in trans_2027_m), 'Reducción de tarifas insulares negociadas (-9.5k €)'),
    ('624000002', 'Transporte internacional (Exportación)', 'Transporte', 0.00, 0.00, 'Repercutido 100% al cliente en factura (305k € facturados; 0 € coste neto Codiagro)'),
    ('622010000', 'Mantenimiento fábrica (preventivo/correctivo)', 'Mantenimiento', 191422.08, tot_maint_27, 'Ahorro masivo (-100.2k €): la fábrica presupuestó 91k € vs histórico de 191k €'),
    ('629000016', 'Herramientas taller', 'Mantenimiento', 6000.36, 2100.00, 'Revisión a la baja del gasto de ferretería'),
    ('627000007', 'Ferias y congresos', 'Marketing', 3855.75, 33770.00, 'Incorporación del calendario ferial internacional formalizado (+29.9k €)'),
    ('627000008', 'Marketing y campañas comerciales', 'Marketing', 0.00, 196300.00, 'Plan integral de activación comercial por cliente-producto (+196.3k €)'),
    ('627000006', 'Publicidad, catálogos y folletos', 'Marketing', 1236.75, 25630.00, 'Actualización corporativa transversal y diseño institucional (+24.4k €)'),
    ('627000009', 'Artículos publicidad / Promociones', 'Marketing', 26500.00, 10372.40, 'Reasignación de reserva táctica (-16.1k €)'),
    ('627000004', 'Relaciones públicas y comerciales', 'Marketing', 18104.50, 0.00, 'Integrado dentro de los programas de cliente-producto'),
    ('629000007', 'Gastos retirada de residuos', 'Regulatory', 34000.00, 74400.00, 'Gasto medioambiental real de 6.200 €/mes por aumento de producción (+40.4k €)'),
    ('629000002', 'Otros gastos / Envases AEVAE', 'Regulatory', 11379.82, 25115.00, 'Declaración anual y cuotas obligatorias de envases agrícolas SCRAP AEVAE'),
    ('623200000', 'Gastos marcas y patentes', 'Regulatory', 0.00, 75390.00, 'Presupuesto de marcas: 16.3k € son inmovilizables (marcas nuevas) y 59k € gasto corriente'),
    ('623000003', 'Servicios registros y notarios', 'Regulatory', 24000.00, 33050.00, 'Registros de producto: 25.9k € son inmovilizables (Azerbaiyán, Sudáfrica, Guatemala)'),
    ('623000004', 'Gastos certificaciones ecológicas', 'Regulatory', 15253.68, 10193.00, 'Certificaciones BCS-Kiwa, OMRI, FIBL, Ecocert (-5.1k € vs previsión inicial)'),
    ('629000008', 'Cuotas asociaciones (Quimacova)', 'Regulatory', 9621.84, 2400.00, 'Alineado con la cuota real de Quimacova (-7.2k €)'),
    ('629000005', 'Gastos análisis de laboratorios', 'Regulatory', 11295.24, 6600.00, 'Análisis de laboratorio externo para expedientes regulatorios'),
]

start_c_row = 6
for idx, r in enumerate(accounts_summary):
    r_idx = start_c_row + idx
    fill_row = FILL_ZEBRA if idx % 2 == 1 else PatternFill(fill_type=None)

    ws4.cell(row=r_idx, column=2, value=r[0]).font = FONT_BOLD
    ws4.cell(row=r_idx, column=3, value=r[1]).font = FONT_DATA
    ws4.cell(row=r_idx, column=4, value=r[2]).font = FONT_DATA

    c_pnl = ws4.cell(row=r_idx, column=5, value=r[3])
    c_pnl.font = FONT_DATA
    c_pnl.number_format = FMT_CURR
    c_pnl.alignment = ALIGN_RIGHT

    c_dpto = ws4.cell(row=r_idx, column=6, value=r[4])
    c_dpto.font = FONT_BOLD
    c_dpto.number_format = FMT_CURR
    c_dpto.alignment = ALIGN_RIGHT

    c_diff = ws4.cell(row=r_idx, column=7, value=f'=F{r_idx}-E{r_idx}')
    c_diff.font = FONT_BOLD
    c_diff.number_format = FMT_CURR
    c_diff.alignment = ALIGN_RIGHT

    c_pct = ws4.cell(row=r_idx, column=8, value=f'=IF(E{r_idx}=0, 0, G{r_idx}/E{r_idx})')
    c_pct.font = FONT_DATA
    c_pct.number_format = FMT_PCT
    c_pct.alignment = ALIGN_RIGHT

    ws4.cell(row=r_idx, column=9, value=r[5]).font = FONT_DATA

    for c_col in range(2, 10):
        cell_c = ws4.cell(row=r_idx, column=c_col)
        cell_c.border = BORDER_THIN
        if fill_row.fill_type: cell_c.fill = fill_row

# Total Row Sheet 4
tot_c_row = start_c_row + len(accounts_summary)
ws4.cell(row=tot_c_row, column=2, value='TOTAL CUENTAS COMPARADAS').font = FONT_TOTAL
ws4.cell(row=tot_c_row, column=3, value='').font = FONT_TOTAL
ws4.cell(row=tot_c_row, column=4, value='').font = FONT_TOTAL

cell_tot_pnl = ws4.cell(row=tot_c_row, column=5, value=f'=SUM(E{start_c_row}:E{tot_c_row-1})')
cell_tot_pnl.font = FONT_TOTAL
cell_tot_pnl.number_format = FMT_CURR
cell_tot_pnl.alignment = ALIGN_RIGHT

cell_tot_dpto = ws4.cell(row=tot_c_row, column=6, value=f'=SUM(F{start_c_row}:F{tot_c_row-1})')
cell_tot_dpto.font = FONT_TOTAL
cell_tot_dpto.number_format = FMT_CURR
cell_tot_dpto.alignment = ALIGN_RIGHT

cell_tot_diff = ws4.cell(row=tot_c_row, column=7, value=f'=SUM(G{start_c_row}:G{tot_c_row-1})')
cell_tot_diff.font = FONT_TOTAL
cell_tot_diff.number_format = FMT_CURR
cell_tot_diff.alignment = ALIGN_RIGHT

ws4.cell(row=tot_c_row, column=8, value=f'=G{tot_c_row}/E{tot_c_row}').number_format = FMT_PCT
ws4.cell(row=tot_c_row, column=8).font = FONT_TOTAL
ws4.cell(row=tot_c_row, column=8).alignment = ALIGN_RIGHT

ws4.cell(row=tot_c_row, column=9, value='Variación global antes de inmovilizaciones').font = FONT_BOLD

for c_col in range(2, 10):
    cell_c = ws4.cell(row=tot_c_row, column=c_col)
    cell_c.border = BORDER_TOTAL
    cell_c.fill = FILL_TOTAL

col_widths_s4 = {
    'A': 3, 'B': 16, 'C': 36, 'D': 20, 'E': 20, 'F': 22, 'G': 20, 'H': 16, 'I': 55
}
for col_l, w in col_widths_s4.items():
    ws4.column_dimensions[col_l].width = w


# -------------------------------------------------------------
# SHEET 5: CRITERIOS Y NORMATIVA PGC
# -------------------------------------------------------------
ws5 = wb.create_sheet(title='Criterios y Normativa PGC')
ws5.views.sheetView[0].showGridLines = True

ws5.merge_cells('B2:H2')
ws5['B2'] = 'MEMORIA DE CRITERIOS CONTABLES Y REGLAS DE ASIGNACIÓN (PGC RD 1514/2007)'
ws5['B2'].font = FONT_TITLE

ws5.merge_cells('B3:H3')
ws5['B3'] = 'Fundamentos técnicos, normas de valoración aplicadas y directrices para la contabilización y auditoría'
ws5['B3'].font = FONT_SUBTITLE

criterios_textos = [
    ('1. APROVISIONAMIENTOS Y COMPRAS (GRUPO 60)', [
        ('Norma de Registro y Valoración 10ª (Existencias):', 'Las materias primas y otros aprovisionamientos se valoran al precio de adquisición, que incluye el importe facturado por el vendedor deducido cualquier descuento comercial y añadiendo todos los gastos adicionales que se produzcan hasta que los bienes se hallen en almacén (fletes, seguros, aranceles).'),
        ('Tratamiento de la Reserva por Mínimos de Pedido (3%):', 'La reserva monetaria de 182.159,08 € contemplada en el presupuesto de compras obedece a exigencias comerciales de los proveedores (lotes mínimos de compra / MOQ). Contablemente, esta compra adicional generará existencias en almacén (activo corriente) a cierre de ejercicio, neutralizando el impacto en la cuenta de resultados a través de la cuenta 610/611 (Variación de existencias).'),
        ('Fusión de Cuentas de Envases:', 'Se mantiene la unificación de la cuenta 602000000 en la 602000001 (Envases y embalajes), centralizando todas las adquisiciones de bidones, garrafas y botellas bajo una única métrica de consumo.')
    ]),
    ('2. TRANSPORTES Y LOGÍSTICA: REPERCUSIÓN DE FLETES (CUENTAS 624 Y 700001624)', [
        ('Tratamiento del Transporte de Exportación Repercutido:', 'El transporte internacional de exportación (305.223,30 € en 2027 y 97.236,00 € en Q4 2026) se factura y repercute íntegramente al cliente según las condiciones contractuales de venta (Incoterms CIF/CFR). En aplicación del principio contable de correlación de ingresos y gastos y la neutralidad del flete suplido, dicho coste se compensa 100% con la facturación emitida al cliente (cuenta 700001624 Venta por transporte repercutido), por lo que NO DEBE COMPUTAR COMO GASTO OPERATIVO de Codiagro ni absorber margen de la cuenta de explotación.'),
        ('Gasto Logístico Real Asumido por Codiagro (245.670,44 €):', 'El presupuesto operativo real que asume la empresa corresponde exclusivamente al Transporte Nacional peninsular (166.472,15 € en Campillo) y Transporte a Canarias (79.198,29 € en Sealine). Al comparar este importe con los 301.548,73 € que figuraban en el P&L actual para estas dos líneas, Codiagro obtiene un AHORRO LOGÍSTICO NETO DE -55.878,29 € (-18,5%).'),
        ('Transporte de Importación de Materias Primas:', 'El transporte de aprovisionamientos importados debe incorporarse contablemente como mayor valor de adquisición de las existencias (Grupo 601), salvo que por materialidad y practicidad se registre en la cuenta 624000006.')
    ]),
    ('3. MANTENIMIENTO: REPARACIÓN ORDINARIA VS ACTIVACIÓN (CUENTAS 622 VS 21X)', [
        ('Principio General (Resolución ICAC de 5 de marzo de 2019):', 'Los costes de mantenimiento, entretenimiento y reparación corriente que tengan por objeto conservar el inmovilizado en condiciones normales de funcionamiento se imputan directamente a la cuenta de pérdidas y ganancias (cuenta 622).'),
        ('Criterio de Inmovilización (Ampliación, Mejora y Gran Reparación):', 'Para poder activar un gasto de mantenimiento como Inmovilizado Material (Grupo 21) deben cumplirse simultáneamente tres requisitos:\n1. Aumento de la capacidad de producción o del rendimiento operativo.\n2. Alargamiento sustancial de la vida útil estimada del bien.\n3. Posibilidad de identificar y valorar de forma fiable el coste incurrido.'),
        ('Aplicación a las Partidas de Codiagro:', '• Cubierta tejado fábrica (2.110 €): Si sustituye placas dañadas de forma permanente, es activable en 211000000 (Construcciones).\n• Molinos Oliver & Batlle (3.000 €): La reconstrucción mayor (overhaul) de los tres molinos es activable en 213000000 (Maquinaria) amortizable a 8 años.\n• Herramientas de taller (2.000 €): Activable en 214000000 (Utillaje) al superar el umbral de fungibilidad.\n• Resto de partidas (correctivo mensual de 3.000 €, válvulas ordinarias, consumibles): Gasto directo del ejercicio en subcuentas 622.')
    ]),
    ('4. REGULATORY: MARCAS Y REGISTROS (CUENTAS 623/629 VS 203)', [
        ('Resolución del ICAC de 28 de mayo de 2013 (Inmovilizado Intangible):', 'Los costes asociados a la obtención de propiedad industrial y registros oficiales se activan en la cuenta 203000000 si confieren derechos exclusivos de explotación protegidos jurídicamente y son capaces de generar rendimientos económicos futuros.'),
        ('Marcas Nuevas (16.300 €) -> ACTIVABLES EN 203000000:', 'El registro inicial de marcas (Codiorgan Guatemala, Trifeno-94 Europa/Guatemala, Salwax Arabia, Biorad Arabia, Sprintium Omán/USA) confiere el monopolio de explotación durante 10 años. Se activan como Inmovilizado Intangible y se amortizan al 10% anual.'),
        ('Mantenimiento y Vigilancia de Marcas (59.090 €) -> GASTO CORRIENTE (6232):', 'Las tasas de renovación periódica, la vigilancia mundial anual (2.550 €) y los cambios de titularidad por mero trámite administrativo no incrementan el valor del activo y se consideran gasto corriente del ejercicio.'),
        ('Registros Oficiales de Producto (25.900 €) -> ACTIVABLES EN 203000000:', 'Los expedientes de registro de formulaciones en Azerbaiyán (13.200 €), Sudáfrica (6.960 €) y Guatemala (5.740 €) constituyen licencias comerciales habilitantes obligatorias sin las cuales la empresa no puede operar. Amortización recomendada en 5 años (periodo estándar de vigencia ministerial).'),
        ('Residuos y Medioambiente (629000007):', 'El servicio continuado de retirada y tratamiento de residuos industriales peligrosos y no peligrosos (74.400 €) es un gasto de explotación ordinario insoslayable.'),
        ('Obra Nave Enfrente (9.600 €):', 'Gasto de regularización y obra finalizada que incrementa la infraestructura patrimonial de la empresa, susceptible de activación en la cuenta 211000000.')
    ]),
    ('5. MARKETING Y PROMOCIÓN (CUENTA 627)', [
        ('Imputación Temporal de Ferias y Campañas:', 'Según el principio de devengo, los gastos de stands, derechos de inscripción y diseño de ferias se devengan en el mes de celebración del evento (octubre para Fruit Attraction, febrero para Fruit Logistica).'),
        ('Prohibición de Activación de Gastos Publicitarios:', 'El PGC prohíbe taxativamente la activación de gastos de publicidad y promoción como activos intangibles (marcas generadas internamente o campañas comerciales), debiendo imputarse íntegramente a pérdidas y ganancias en el ejercicio en que se reciben los servicios.')
    ])
]

row_s5 = 5
for sec_title, sec_items in criterios_textos:
    ws5.merge_cells(f'B{row_s5}:H{row_s5}')
    ws5.cell(row=row_s5, column=2, value=sec_title).font = FONT_SECTION
    ws5.cell(row=row_s5, column=2).fill = FILL_NAVY
    ws5.cell(row=row_s5, column=2).alignment = Alignment(vertical='center', indent=1)
    ws5.row_dimensions[row_s5].height = 24
    row_s5 += 1

    for title_it, desc_it in sec_items:
        ws5.cell(row=row_s5, column=2, value=title_it).font = FONT_BOLD
        ws5.cell(row=row_s5, column=2).alignment = Alignment(vertical='top')
        
        ws5.merge_cells(f'C{row_s5}:H{row_s5}')
        c_desc = ws5.cell(row=row_s5, column=3, value=desc_it)
        c_desc.font = FONT_DATA
        c_desc.alignment = Alignment(vertical='top', wrap_text=True)

        for col_i in range(2, 9):
            ws5.cell(row=row_s5, column=col_i).border = BORDER_THIN
        
        ws5.row_dimensions[row_s5].height = 55 if len(desc_it) > 140 else 30
        row_s5 += 1
    row_s5 += 1

ws5.column_dimensions['A'].width = 3
ws5.column_dimensions['B'].width = 38
ws5.column_dimensions['C'].width = 16
ws5.column_dimensions['D'].width = 16
ws5.column_dimensions['E'].width = 16
ws5.column_dimensions['F'].width = 16
ws5.column_dimensions['G'].width = 16
ws5.column_dimensions['H'].width = 32

wb.save(EXCEL_OUT)
print(f"Workbook actualizado con la salvedad de exportación guardado en: {EXCEL_OUT}")
