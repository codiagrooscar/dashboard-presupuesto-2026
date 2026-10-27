import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os, sys

sys.stdout.reconfigure(encoding='utf-8')

print("1. Cargando datos y precios de referencia...")

# 1. Cargar precios de referencia
prices_map = {}
if os.path.exists('Consolidado_Ventas_y_Previsiones_2026.xlsx'):
    try:
        wb_c = openpyxl.load_workbook('Consolidado_Ventas_y_Previsiones_2026.xlsx', data_only=True)
        ws_det = wb_c['Detalle Lneas 2026'] if 'Detalle Lneas 2026' in wb_c.sheetnames else wb_c[wb_c.sheetnames[3]]
        for r in range(2, ws_det.max_row + 1):
            sku1 = str(ws_det.cell(r, 5).value or '').strip().upper()
            sku2 = str(ws_det.cell(r, 6).value or '').strip().upper()
            pm = ws_det.cell(r, 10).value
            if isinstance(pm, (int, float)) and pm > 0:
                if sku1 and sku1 not in prices_map: prices_map[sku1] = pm
                if sku2 and sku2 not in prices_map: prices_map[sku2] = pm
    except Exception as e:
        print("Aviso al cargar precios:", e)

# 2. Cargar datos de los archivos originales y copias
folder = 'Revision_Previsiones_Q4_2026'
comerciales = ['Alfonso', 'Garcia', 'Irene', 'Javier', 'Pedro', 'Ricardo']

data_by_com = {}
all_modified_lines = []
all_comments_log = []

summary_rows = []

for com in comerciales:
    f_orig = os.path.join(folder, f'Revision_Previsiones_Q4_2026_{com}.xlsx')
    f_copy = os.path.join(folder, f'Copia de Revision_Previsiones_Q4_2026_{com}.xlsx')
    
    if not os.path.exists(f_copy):
        continue
        
    wb_orig = openpyxl.load_workbook(f_orig, data_only=True)
    wb_copy = openpyxl.load_workbook(f_copy, data_only=True)
    
    ws_orig = wb_orig.active
    ws_copy = wb_copy.active
    
    com_records = []
    
    com_stats = {
        'comercial': com,
        'lineas_totales': 0,
        'lineas_modificadas': 0,
        'comentarios_sep': 0,
        'comentarios_q4': 0,
        'oct_ant': 0.0, 'oct_rev': 0.0,
        'nov_ant': 0.0, 'nov_rev': 0.0,
        'dic_ant': 0.0, 'dic_rev': 0.0,
        'eur_ant': 0.0, 'eur_rev': 0.0
    }
    
    for r in range(10, ws_copy.max_row + 1):
        pais = ws_copy.cell(r, 2).value
        cliente = ws_copy.cell(r, 3).value
        sku = ws_copy.cell(r, 4).value
        desc = ws_copy.cell(r, 5).value
        
        if not sku or 'TOTAL' in str(cliente).upper() or 'TOTAL' in str(sku).upper():
            continue
            
        com_stats['lineas_totales'] += 1
        
        sep_prev = float(ws_copy.cell(r, 6).value or 0)
        sep_real = float(ws_copy.cell(r, 7).value or 0)
        com_sep = ws_copy.cell(r, 10).value
        com_sep_o = ws_orig.cell(r, 10).value
        
        oct_ant = float(ws_copy.cell(r, 11).value or 0)
        oct_rev = float(ws_copy.cell(r, 12).value or 0)
        nov_ant = float(ws_copy.cell(r, 13).value or 0)
        nov_rev = float(ws_copy.cell(r, 14).value or 0)
        dic_ant = float(ws_copy.cell(r, 15).value or 0)
        dic_rev = float(ws_copy.cell(r, 16).value or 0)
        
        com_q4 = ws_copy.cell(r, 20).value
        com_q4_o = ws_orig.cell(r, 20).value
        
        has_qty_change = (oct_ant != oct_rev) or (nov_ant != nov_rev) or (dic_ant != dic_rev)
        has_sep_com = bool(com_sep and com_sep != com_sep_o)
        has_q4_com = bool(com_q4 and com_q4 != com_q4_o)
        
        price = prices_map.get(str(sku).strip().upper(), 2.59)
        
        q4_ant = oct_ant + nov_ant + dic_ant
        q4_rev = oct_rev + nov_rev + dic_rev
        q4_dif = q4_rev - q4_ant
        
        eur_ant_line = q4_ant * price
        eur_rev_line = q4_rev * price
        
        com_stats['oct_ant'] += oct_ant
        com_stats['oct_rev'] += oct_rev
        com_stats['nov_ant'] += nov_ant
        com_stats['nov_rev'] += nov_rev
        com_stats['dic_ant'] += dic_ant
        com_stats['dic_rev'] += dic_rev
        com_stats['eur_ant'] += eur_ant_line
        com_stats['eur_rev'] += eur_rev_line
        
        if has_qty_change:
            com_stats['lineas_modificadas'] += 1
        if has_sep_com or (com_sep and str(com_sep).strip()):
            com_stats['comentarios_sep'] += 1
        if has_q4_com or (com_q4 and str(com_q4).strip()):
            com_stats['comentarios_q4'] += 1
            
        record = {
            'comercial': com,
            'fila': r,
            'pais': pais,
            'cliente': cliente,
            'sku': sku,
            'desc': desc,
            'precio': price,
            'sep_prev': sep_prev,
            'sep_real': sep_real,
            'com_sep': com_sep,
            'oct_ant': oct_ant, 'oct_rev': oct_rev, 'oct_dif': oct_rev - oct_ant,
            'nov_ant': nov_ant, 'nov_rev': nov_rev, 'nov_dif': nov_rev - nov_ant,
            'dic_ant': dic_ant, 'dic_rev': dic_rev, 'dic_dif': dic_rev - dic_ant,
            'q4_ant': q4_ant, 'q4_rev': q4_rev, 'q4_dif': q4_dif,
            'eur_ant': eur_ant_line, 'eur_rev': eur_rev_line, 'eur_dif': eur_rev_line - eur_ant_line,
            'com_q4': com_q4,
            'has_qty_change': has_qty_change
        }
        com_records.append(record)
        
        if has_qty_change:
            all_modified_lines.append(record)
            
        if (com_sep and str(com_sep).strip()) or (com_q4 and str(com_q4).strip()):
            all_comments_log.append(record)
            
    data_by_com[com] = com_records
    summary_rows.append(com_stats)

print(f"Líneas con cambios numéricos: {len(all_modified_lines)}")
print(f"Líneas con comentarios: {len(all_comments_log)}")

# -------------------------------------------------------------------------
# CONSTRUCCIÓN DEL LIBRO EXCEL
# -------------------------------------------------------------------------
wb = openpyxl.Workbook()
wb.remove(wb.active) # Quitar hoja default

# Paleta Corporativa Codiagro
FONT_TITLE = Font(name='Segoe UI', size=16, bold=True, color='FFFFFF')
FONT_SUBTITLE = Font(name='Segoe UI', size=10, italic=True, color='94A3B8')
FONT_SEC_HDR = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
FONT_TH = Font(name='Segoe UI', size=9, bold=True, color='FFFFFF')
FONT_DATA = Font(name='Segoe UI', size=9, color='0F172A')
FONT_BOLD = Font(name='Segoe UI', size=9, bold=True, color='0F172A')
FONT_TOTAL = Font(name='Segoe UI', size=10, bold=True, color='0F172A')
FONT_POS = Font(name='Segoe UI', size=9, bold=True, color='047857')
FONT_NEG = Font(name='Segoe UI', size=9, bold=True, color='B91C1C')
FONT_NEU = Font(name='Segoe UI', size=9, color='475569')

FONT_CARD_TITLE = Font(name='Segoe UI', size=8, bold=True, color='475569')
FONT_CARD_VAL = Font(name='Segoe UI', size=16, bold=True, color='0F172A')
FONT_CARD_SUB = Font(name='Segoe UI', size=8, italic=True, color='64748B')

FILL_NAVY = PatternFill(start_color='0F172A', end_color='0F172A', fill_type='solid')
FILL_HEADER = PatternFill(start_color='1E3A8A', end_color='1E3A8A', fill_type='solid')
FILL_SUBHEADER = PatternFill(start_color='334155', end_color='334155', fill_type='solid')
FILL_ZEBRA = PatternFill(start_color='F8FAFC', end_color='F8FAFC', fill_type='solid')
FILL_TOTAL = PatternFill(start_color='E2E8F0', end_color='E2E8F0', fill_type='solid')
FILL_POS_LIGHT = PatternFill(start_color='ECFDF5', end_color='ECFDF5', fill_type='solid')
FILL_NEG_LIGHT = PatternFill(start_color='FEF2F2', end_color='FEF2F2', fill_type='solid')
FILL_CARD_BG = PatternFill(start_color='F1F5F9', end_color='F1F5F9', fill_type='solid')

BORDER_THIN = Border(
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1'),
    top=Side(style='thin', color='CBD5E1'),
    bottom=Side(style='thin', color='CBD5E1')
)
BORDER_HEADER = Border(
    left=Side(style='thin', color='475569'),
    right=Side(style='thin', color='475569'),
    top=Side(style='thin', color='1E3A8A'),
    bottom=Side(style='medium', color='0F172A')
)
BORDER_TOTAL = Border(
    top=Side(style='thin', color='0F172A'),
    bottom=Side(style='double', color='0F172A'),
    left=Side(style='thin', color='CBD5E1'),
    right=Side(style='thin', color='CBD5E1')
)
BORDER_CARD = Border(
    left=Side(style='thin', color='94A3B8'),
    right=Side(style='thin', color='94A3B8'),
    top=Side(style='medium', color='1E3A8A'),
    bottom=Side(style='thin', color='94A3B8')
)

FMT_INT = '#,##0'
FMT_INT_DIFF = '+#,##0;-#,##0;"-"'
FMT_CURR = '#,##0.00 €;[Red]-#,##0.00 €;"-"'
FMT_CURR_DIFF = '+#,##0.00 €;-#,##0.00 €;"-"'
FMT_PCT = '+0.0%;-0.0%;"0.0%"'

ALIGN_LEFT = Alignment(horizontal='left', vertical='center')
ALIGN_RIGHT = Alignment(horizontal='right', vertical='center')
ALIGN_CENTER = Alignment(horizontal='center', vertical='center')
ALIGN_HDR = Alignment(horizontal='center', vertical='center', wrap_text=True)

# -------------------------------------------------------------------------
# HOJA 1: RESUMEN EJECUTIVO
# -------------------------------------------------------------------------
ws1 = wb.create_sheet(title='Resumen Ejecutivo')
ws1.views.sheetView[0].showGridLines = True

# Banner
ws1.merge_cells('B2:R2')
ws1['B2'] = 'CODIAGRO · INFORME CONSOLIDADO DE REVISIÓN DE PREVISIONES Q4 2026'
ws1['B2'].font = FONT_TITLE
ws1['B2'].fill = FILL_NAVY
ws1['B2'].alignment = Alignment(horizontal='left', vertical='center', indent=1)
ws1.row_dimensions[2].height = 35

ws1.merge_cells('B3:R3')
ws1['B3'] = 'Consolidación de plantillas revisadas devueltas por los delegados comerciales (Octubre - Diciembre 2026) | Fecha: 02/10/2026'
ws1['B3'].font = FONT_SUBTITLE
ws1['B3'].fill = PatternFill(start_color='1E293B', end_color='1E293B', fill_type='solid')
ws1['B3'].alignment = Alignment(horizontal='left', vertical='center', indent=1)
ws1.row_dimensions[3].height = 20

# KPI Cards
# Card 1: Líneas Modificadas
ws1.merge_cells('B5:D5')
ws1['B5'] = 'LÍNEAS CON CAMBIOS Q4'
ws1['B5'].font = FONT_CARD_TITLE
ws1['B5'].alignment = ALIGN_CENTER
ws1.merge_cells('B6:D6')
ws1['B6'] = len(all_modified_lines)
ws1['B6'].font = FONT_CARD_VAL
ws1['B6'].alignment = ALIGN_CENTER
ws1.merge_cells('B7:D7')
ws1['B7'] = '27 líneas modificadas en Q4'
ws1['B7'].font = FONT_CARD_SUB
ws1['B7'].alignment = ALIGN_CENTER

# Card 2: Previsión Inicial Q4
tot_q4_ant = sum(s['oct_ant'] + s['nov_ant'] + s['dic_ant'] for s in summary_rows)
ws1.merge_cells('E5:G5')
ws1['E5'] = 'PREVISIÓN INICIAL Q4'
ws1['E5'].font = FONT_CARD_TITLE
ws1['E5'].alignment = ALIGN_CENTER
ws1.merge_cells('E6:G6')
ws1['E6'] = tot_q4_ant
ws1['E6'].font = FONT_CARD_VAL
ws1['E6'].number_format = FMT_INT
ws1['E6'].alignment = ALIGN_CENTER
ws1.merge_cells('E7:G7')
ws1['E7'] = 'Total unidades planificadas inicial'
ws1['E7'].font = FONT_CARD_SUB
ws1['E7'].alignment = ALIGN_CENTER

# Card 3: Previsión Revisada Q4
tot_q4_rev = sum(s['oct_rev'] + s['nov_rev'] + s['dic_rev'] for s in summary_rows)
ws1.merge_cells('H5:J5')
ws1['H5'] = 'PREVISIÓN REVISADA Q4'
ws1['H5'].font = FONT_CARD_TITLE
ws1['H5'].alignment = ALIGN_CENTER
ws1.merge_cells('H6:J6')
ws1['H6'] = tot_q4_rev
ws1['H6'].font = FONT_CARD_VAL
ws1['H6'].number_format = FMT_INT
ws1['H6'].alignment = ALIGN_CENTER
ws1.merge_cells('H7:J7')
ws1['H7'] = 'Unidades tras revisión comercial'
ws1['H7'].font = FONT_CARD_SUB
ws1['H7'].alignment = ALIGN_CENTER

# Card 4: Variación Neta Q4
tot_q4_dif = tot_q4_rev - tot_q4_ant
ws1.merge_cells('K5:M5')
ws1['K5'] = 'VARIACIÓN NETA Q4'
ws1['K5'].font = FONT_CARD_TITLE
ws1['K5'].alignment = ALIGN_CENTER
ws1.merge_cells('K6:M6')
ws1['K6'] = tot_q4_dif
ws1['K6'].font = Font(name='Segoe UI', size=16, bold=True, color='B91C1C' if tot_q4_dif < 0 else '047857')
ws1['K6'].number_format = FMT_INT_DIFF
ws1['K6'].alignment = ALIGN_CENTER
ws1.merge_cells('K7:M7')
ws1['K7'] = f"{tot_q4_dif / tot_q4_ant:+.2%} sobre el inicial"
ws1['K7'].font = FONT_CARD_SUB
ws1['K7'].alignment = ALIGN_CENTER

# Card 5: Desplazamiento Mensual
tot_oct_dif = sum(s['oct_rev'] - s['oct_ant'] for s in summary_rows)
tot_nov_dif = sum(s['nov_rev'] - s['nov_ant'] for s in summary_rows)
tot_dic_dif = sum(s['dic_rev'] - s['dic_ant'] for s in summary_rows)
ws1.merge_cells('N5:P5')
ws1['N5'] = 'DESPLAZAMIENTO MENSUAL'
ws1['N5'].font = FONT_CARD_TITLE
ws1['N5'].alignment = ALIGN_CENTER
ws1.merge_cells('N6:P6')
ws1['N6'] = f"Oct {tot_oct_dif:+,.0f} | Nov {tot_nov_dif:+,.0f} | Dic {tot_dic_dif:+,.0f}"
ws1['N6'].font = Font(name='Segoe UI', size=11, bold=True, color='1E3A8A')
ws1['N6'].alignment = ALIGN_CENTER
ws1.merge_cells('N7:P7')
ws1['N7'] = 'Trasvase relevante de Noviembre hacia Diciembre'
ws1['N7'].font = FONT_CARD_SUB
ws1['N7'].alignment = ALIGN_CENTER

# Card 6: Comentarios Cualitativos
tot_comms = len(all_comments_log)
ws1.merge_cells('Q5:R5')
ws1['Q5'] = 'JUSTIFICACIONES'
ws1['Q5'].font = FONT_CARD_TITLE
ws1['Q5'].alignment = ALIGN_CENTER
ws1.merge_cells('Q6:R6')
ws1['Q6'] = tot_comms
ws1['Q6'].font = FONT_CARD_VAL
ws1['Q6'].alignment = ALIGN_CENTER
ws1.merge_cells('Q7:R7')
ws1['Q7'] = 'Notas de cierre y campaña'
ws1['Q7'].font = FONT_CARD_SUB
ws1['Q7'].alignment = ALIGN_CENTER

# Aplicar fondo y bordes a cards
for col_start, col_end in [(2,4), (5,7), (8,10), (11,13), (14,16), (17,18)]:
    for r in range(5, 8):
        for c in range(col_start, col_end + 1):
            cell = ws1.cell(r, c)
            cell.fill = FILL_CARD_BG
            # Borde exterior de card
            top_side = Side(style='medium', color='1E3A8A') if r == 5 else Side(style='thin', color='E2E8F0')
            bot_side = Side(style='thin', color='94A3B8') if r == 7 else Side(style='thin', color='E2E8F0')
            left_side = Side(style='thin', color='94A3B8') if c == col_start else None
            right_side = Side(style='thin', color='94A3B8') if c == col_end else None
            cell.border = Border(top=top_side, bottom=bot_side, left=left_side, right=right_side)

ws1.row_dimensions[5].height = 18
ws1.row_dimensions[6].height = 28
ws1.row_dimensions[7].height = 18

# Sección 1: Tabla Resumen por Comercial
ws1.merge_cells('B9:R9')
ws1['B9'] = '1. BALANCE COMPARATIVO DE PREVISIONES POR DELEGADO COMERCIAL (Q4 2026)'
ws1['B9'].font = FONT_SEC_HDR
ws1['B9'].fill = FILL_HEADER
ws1['B9'].alignment = Alignment(horizontal='left', vertical='center', indent=1)
ws1.row_dimensions[9].height = 24

headers_s1 = [
    'Comercial', 'Líneas Rev.', 'Coment.',
    'Oct Ant (u)', 'Oct Rev (u)', 'Dif Oct (u)',
    'Nov Ant (u)', 'Nov Rev (u)', 'Dif Nov (u)',
    'Dic Ant (u)', 'Dic Rev (u)', 'Dic Dif (u)',
    'Total Q4 Ant (u)', 'Total Q4 Rev (u)', 'Dif Q4 (u)', 'Var (%)',
    'Patrón Principal / Análisis de los Cambios'
]
ws1.row_dimensions[10].height = 28
for idx, h in enumerate(headers_s1, start=2):
    cell = ws1.cell(10, idx, h)
    cell.font = FONT_TH
    cell.fill = FILL_SUBHEADER
    cell.alignment = ALIGN_HDR
    cell.border = BORDER_HEADER

com_descriptions = {
    'Alfonso': 'Ajuste fino en La Veguilla (-1.560 u en Biorad y Brotamec) por adelanto de pedidos al cierre de Septiembre.',
    'Garcia': 'Dotra Chemicals reduce 8.000 u en Noviembre (+1.000 u en Oct); Marcoser sustituye formatos. Aporta 24 notas de campo.',
    'Irene': 'Reajustes por adelanto de compras y auditorías en Campoejido y CASI; refuerzo en Floramec Avance 20L (+320 u), PGR IV (+1.920 u) y Brotamec (+600 u). Aporta 27 comentarios de campo.',
    'Javier': 'Reducción en Octubre por pedidos adelantados a Septiembre en Fitosur y Cabrera Morales; Drago traslada 5.840 u a Diciembre.',
    'Pedro': 'Adelanto e incremento en Asesoramiento Técnico (+1.580 u en Octubre); ajuste a la baja en Fitocarthago (-3.400 u).',
    'Ricardo': 'Fuerte incremento en Lekkerbio (+22.000 u en Octubre). Trasvase masivo de Noviembre a Diciembre en El Llano y Campo Abierto.'
}

curr_r = 11
for s in summary_rows:
    com = s['comercial']
    q4_ant = s['oct_ant'] + s['nov_ant'] + s['dic_ant']
    q4_rev = s['oct_rev'] + s['nov_rev'] + s['dic_rev']
    q4_dif = q4_rev - q4_ant
    pct_var = (q4_dif / q4_ant) if q4_ant != 0 else 0.0
    
    vals = [
        com, s['lineas_modificadas'], s['comentarios_sep'] + s['comentarios_q4'],
        s['oct_ant'], s['oct_rev'], s['oct_rev'] - s['oct_ant'],
        s['nov_ant'], s['nov_rev'], s['nov_rev'] - s['nov_ant'],
        s['dic_ant'], s['dic_rev'], s['dic_rev'] - s['dic_ant'],
        q4_ant, q4_rev, q4_dif, pct_var,
        com_descriptions.get(com, '')
    ]
    
    for idx, v in enumerate(vals, start=2):
        cell = ws1.cell(curr_r, idx, v)
        cell.font = FONT_DATA
        cell.border = BORDER_THIN
        cell.fill = FILL_ZEBRA if curr_r % 2 == 0 else PatternFill(fill_type=None)
        
        if idx == 2:
            cell.alignment = ALIGN_LEFT
            cell.font = FONT_BOLD
        elif idx in [3, 4]:
            cell.alignment = ALIGN_CENTER
            cell.number_format = FMT_INT
        elif idx in [5, 6, 8, 9, 11, 12, 14, 15]:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = FMT_INT
        elif idx in [7, 10, 13, 16]:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = FMT_INT_DIFF
            if isinstance(v, (int, float)):
                if v > 0: cell.font = FONT_POS; cell.fill = FILL_POS_LIGHT
                elif v < 0: cell.font = FONT_NEG; cell.fill = FILL_NEG_LIGHT
        elif idx == 17:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = FMT_PCT
            if isinstance(v, (int, float)):
                if v > 0: cell.font = FONT_POS
                elif v < 0: cell.font = FONT_NEG
        elif idx == 18:
            cell.alignment = ALIGN_LEFT
            
    ws1.row_dimensions[curr_r].height = 22
    curr_r += 1

# Total Fila Resumen
total_lineas_mod = len(all_modified_lines)
total_comms_count = len(all_comments_log)
tot_vals = [
    'TOTAL CODIAGRO', total_lineas_mod, total_comms_count,
    sum(s['oct_ant'] for s in summary_rows), sum(s['oct_rev'] for s in summary_rows), tot_oct_dif,
    sum(s['nov_ant'] for s in summary_rows), sum(s['nov_rev'] for s in summary_rows), tot_nov_dif,
    sum(s['dic_ant'] for s in summary_rows), sum(s['dic_rev'] for s in summary_rows), tot_dic_dif,
    tot_q4_ant, tot_q4_rev, tot_q4_dif, (tot_q4_dif / tot_q4_ant) if tot_q4_ant else 0.0,
    'Variación neta moderada (-1,29%), con fuerte retraso de pedidos de Noviembre a Diciembre.'
]

for idx, v in enumerate(tot_vals, start=2):
    cell = ws1.cell(curr_r, idx, v)
    cell.font = FONT_TOTAL
    cell.fill = FILL_TOTAL
    cell.border = BORDER_TOTAL
    if idx == 2:
        cell.alignment = ALIGN_LEFT
    elif idx in [3, 4]:
        cell.alignment = ALIGN_CENTER
        cell.number_format = FMT_INT
    elif idx in [5, 6, 8, 9, 11, 12, 14, 15]:
        cell.alignment = ALIGN_RIGHT
        cell.number_format = FMT_INT
    elif idx in [7, 10, 13, 16]:
        cell.alignment = ALIGN_RIGHT
        cell.number_format = FMT_INT_DIFF
        if isinstance(v, (int, float)):
            if v > 0: cell.font = FONT_POS
            elif v < 0: cell.font = FONT_NEG
    elif idx == 17:
        cell.alignment = ALIGN_RIGHT
        cell.number_format = FMT_PCT
    elif idx == 18:
        cell.alignment = ALIGN_LEFT

ws1.row_dimensions[curr_r].height = 25
curr_r += 2

# Sección 2: Análisis Cualitativo y Hallazgos
ws1.merge_cells(f'B{curr_r}:R{curr_r}')
ws1[f'B{curr_r}'] = '2. DICTAMEN DE CONTROLLING Y RECOMENDACIONES PARA EL CIERRE DE EJERCICIO'
ws1[f'B{curr_r}'].font = FONT_SEC_HDR
ws1[f'B{curr_r}'].fill = FILL_HEADER
ws1[f'B{curr_r}'].alignment = Alignment(horizontal='left', vertical='center', indent=1)
ws1.row_dimensions[curr_r].height = 24
curr_r += 1

dictamen_puntos = [
    ("Efecto Calendario (Noviembre a Diciembre):", "Se detecta un trasvase masivo de 54.820 unidades que salen de Noviembre y se posponen a Diciembre (+38.680 u), protagonizado principalmente por Ricardo en Servicios Agropecuarios El Llano (16.000 u) y Campo Abierto (16.000 u), así como Javier en Drago (5.840 u). Esto concentrará la actividad de expediciones y facturación al final de año."),
    ("Compensación por Adelantos a Septiembre:", "Alfonso, Javier y Pedro reflejan que parte de las caídas de Octubre obedecen a pedidos de Biorad, Brotamec y Fitosanitarios que los clientes solicitaron adelantar al cierre de Septiembre (aprovechando condiciones de campaña o stock)."),
    ("Potencial de Crecimiento Internacional:", "La revisión de Lekkerbio B.V. (Ricardo) añade +22.000 unidades en Octubre (duplicando previsión), lo cual compensa casi en su totalidad las reducciones netas del resto de zonas."),
    ("Estabilidad Global del Presupuesto:", "La variación neta final sobre el volumen total de Q4 es de solo -13.600 unidades (-1,29%), manteniendo prácticamente inalterada la previsión de cierre de ejercicio 2026 en torno a 5,83 - 5,84 M de unidades.")
]

for tit, desc in dictamen_puntos:
    ws1.merge_cells(f'B{curr_r}:D{curr_r}')
    ws1[f'B{curr_r}'] = tit
    ws1[f'B{curr_r}'].font = FONT_BOLD
    ws1[f'B{curr_r}'].alignment = Alignment(horizontal='left', vertical='top')
    
    ws1.merge_cells(f'E{curr_r}:R{curr_r}')
    ws1[f'E{curr_r}'] = desc
    ws1[f'E{curr_r}'].font = FONT_DATA
    ws1[f'E{curr_r}'].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    ws1.row_dimensions[curr_r].height = 28
    curr_r += 1

# -------------------------------------------------------------------------
# HOJA 2: DETALLE CAMBIOS CANTIDADES Q4
# -------------------------------------------------------------------------
ws2 = wb.create_sheet(title='Detalle Cambios Q4')
ws2.views.sheetView[0].showGridLines = True

ws2.merge_cells('B2:T2')
ws2['B2'] = 'DETALLE LÍNEA A LÍNEA DE LOS CAMBIOS DE CANTIDADES EN Q4 2026'
ws2['B2'].font = FONT_TITLE
ws2['B2'].fill = FILL_NAVY
ws2['B2'].alignment = Alignment(horizontal='left', vertical='center', indent=1)
ws2.row_dimensions[2].height = 32

ws2.merge_cells('B3:T3')
ws2['B3'] = 'Registro exhaustivo de las 27 líneas con diferencias numéricas entre la previsión inicial y la devuelta por los comerciales'
ws2['B3'].font = FONT_SUBTITLE
ws2['B3'].fill = PatternFill(start_color='1E293B', end_color='1E293B', fill_type='solid')
ws2['B3'].alignment = Alignment(horizontal='left', vertical='center', indent=1)
ws2.row_dimensions[3].height = 20

headers_s2 = [
    'Comercial', 'País', 'Cliente', 'Cód. SKU', 'Descripción Artículo', 'P. Medio (€/u)',
    'Sep Prev (u)', 'Sep Real (u)', 'Comentario Cierre Sep',
    'Oct Ant (u)', 'Oct Rev (u)', 'Dif Oct (u)',
    'Nov Ant (u)', 'Nov Rev (u)', 'Dif Nov (u)',
    'Dic Ant (u)', 'Dic Rev (u)', 'Dic Dif (u)',
    'Total Q4 Ant', 'Total Q4 Rev', 'Dif Q4 Neta (u)', 'Dif Q4 (€ Est.)', 'Comentario / Justificación Q4'
]

ws2.row_dimensions[5].height = 26
for idx, h in enumerate(headers_s2, start=2):
    cell = ws2.cell(5, idx, h)
    cell.font = FONT_TH
    cell.fill = FILL_HEADER
    cell.alignment = ALIGN_HDR
    cell.border = BORDER_HEADER

r_idx = 6
for ch in sorted(all_modified_lines, key=lambda x: (x['comercial'], x['cliente'], x['sku'])):
    vals = [
        ch['comercial'], ch['pais'], ch['cliente'], ch['sku'], ch['desc'], ch['precio'],
        ch['sep_prev'], ch['sep_real'], ch['com_sep'] or '-',
        ch['oct_ant'], ch['oct_rev'], ch['oct_dif'],
        ch['nov_ant'], ch['nov_rev'], ch['nov_dif'],
        ch['dic_ant'], ch['dic_rev'], ch['dic_dif'],
        ch['q4_ant'], ch['q4_rev'], ch['q4_dif'], ch['eur_dif'],
        ch['com_q4'] or '-'
    ]
    
    for c_idx, v in enumerate(vals, start=2):
        cell = ws2.cell(r_idx, c_idx, v)
        cell.font = FONT_DATA
        cell.border = BORDER_THIN
        cell.fill = FILL_ZEBRA if r_idx % 2 == 0 else PatternFill(fill_type=None)
        
        if c_idx in [2, 3, 4, 5, 6]:
            cell.alignment = ALIGN_LEFT
        elif c_idx == 7: # Precio
            cell.alignment = ALIGN_RIGHT
            cell.number_format = FMT_CURR
        elif c_idx in [8, 9]:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = FMT_INT
        elif c_idx == 10:
            cell.alignment = ALIGN_LEFT
        elif c_idx in [11, 12, 14, 15, 17, 18, 20, 21]:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = FMT_INT
        elif c_idx in [13, 16, 19, 22]:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = FMT_INT_DIFF
            if isinstance(v, (int, float)):
                if v > 0: cell.font = FONT_POS; cell.fill = FILL_POS_LIGHT
                elif v < 0: cell.font = FONT_NEG; cell.fill = FILL_NEG_LIGHT
        elif c_idx == 23:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = FMT_CURR_DIFF
            if isinstance(v, (int, float)):
                if v > 0: cell.font = FONT_POS
                elif v < 0: cell.font = FONT_NEG
        elif c_idx == 24:
            cell.alignment = ALIGN_LEFT
            
    ws2.row_dimensions[r_idx].height = 20
    r_idx += 1

# Total Fila Detalle
ws2.cell(r_idx, 2, 'TOTALES').font = FONT_TOTAL
ws2.cell(r_idx, 2).alignment = ALIGN_LEFT
for c_idx in range(2, len(headers_s2) + 2):
    cell = ws2.cell(r_idx, c_idx)
    cell.fill = FILL_TOTAL
    cell.border = BORDER_TOTAL
    
ws2.cell(r_idx, 11, sum(x['oct_ant'] for x in all_modified_lines)).number_format = FMT_INT
ws2.cell(r_idx, 12, sum(x['oct_rev'] for x in all_modified_lines)).number_format = FMT_INT
ws2.cell(r_idx, 13, sum(x['oct_dif'] for x in all_modified_lines)).number_format = FMT_INT_DIFF
ws2.cell(r_idx, 14, sum(x['nov_ant'] for x in all_modified_lines)).number_format = FMT_INT
ws2.cell(r_idx, 15, sum(x['nov_rev'] for x in all_modified_lines)).number_format = FMT_INT
ws2.cell(r_idx, 16, sum(x['nov_dif'] for x in all_modified_lines)).number_format = FMT_INT_DIFF
ws2.cell(r_idx, 17, sum(x['dic_ant'] for x in all_modified_lines)).number_format = FMT_INT
ws2.cell(r_idx, 18, sum(x['dic_rev'] for x in all_modified_lines)).number_format = FMT_INT
ws2.cell(r_idx, 19, sum(x['dic_dif'] for x in all_modified_lines)).number_format = FMT_INT_DIFF
ws2.cell(r_idx, 20, sum(x['q4_ant'] for x in all_modified_lines)).number_format = FMT_INT
ws2.cell(r_idx, 21, sum(x['q4_rev'] for x in all_modified_lines)).number_format = FMT_INT
ws2.cell(r_idx, 22, sum(x['q4_dif'] for x in all_modified_lines)).number_format = FMT_INT_DIFF
ws2.cell(r_idx, 23, sum(x['eur_dif'] for x in all_modified_lines)).number_format = FMT_CURR_DIFF
ws2.row_dimensions[r_idx].height = 24

# -------------------------------------------------------------------------
# HOJA 3: REGISTRO TODOS LOS COMENTARIOS
# -------------------------------------------------------------------------
ws3 = wb.create_sheet(title='Todos los Comentarios')
ws3.views.sheetView[0].showGridLines = True

ws3.merge_cells('B2:K2')
ws3['B2'] = 'REGISTRO COMPLETO DE JUSTIFICACIONES Y COMENTARIOS DE LA RED COMERCIAL'
ws3['B2'].font = FONT_TITLE
ws3['B2'].fill = FILL_NAVY
ws3['B2'].alignment = Alignment(horizontal='left', vertical='center', indent=1)
ws3.row_dimensions[2].height = 32

ws3.merge_cells('B3:K3')
ws3['B3'] = 'Notas explicativas aportadas sobre el cierre de Septiembre (desviaciones) y la planificación de Q4'
ws3['B3'].font = FONT_SUBTITLE
ws3['B3'].fill = PatternFill(start_color='1E293B', end_color='1E293B', fill_type='solid')
ws3['B3'].alignment = Alignment(horizontal='left', vertical='center', indent=1)
ws3.row_dimensions[3].height = 20

headers_s3 = [
    'Comercial', 'País', 'Cliente', 'Cód. SKU', 'Descripción Artículo',
    'Sep Prev (u)', 'Sep Real (u)', 'Motivo Desviación Septiembre',
    'Total Q4 Rev (u)', 'Dif Q4 (u)', 'Comentario / Justificación Q4'
]

ws3.row_dimensions[5].height = 26
for idx, h in enumerate(headers_s3, start=2):
    cell = ws3.cell(5, idx, h)
    cell.font = FONT_TH
    cell.fill = FILL_HEADER
    cell.alignment = ALIGN_HDR
    cell.border = BORDER_HEADER

r3_idx = 6
for item in sorted(all_comments_log, key=lambda x: (x['comercial'], x['cliente'])):
    vals = [
        item['comercial'], item['pais'], item['cliente'], item['sku'], item['desc'],
        item['sep_prev'], item['sep_real'], item['com_sep'] or '-',
        item['q4_rev'], item['q4_dif'], item['com_q4'] or '-'
    ]
    for c_idx, v in enumerate(vals, start=2):
        cell = ws3.cell(r3_idx, c_idx, v)
        cell.font = FONT_DATA
        cell.border = BORDER_THIN
        cell.fill = FILL_ZEBRA if r3_idx % 2 == 0 else PatternFill(fill_type=None)
        
        if c_idx in [2, 3, 4, 5, 6]:
            cell.alignment = ALIGN_LEFT
        elif c_idx in [7, 8, 10]:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = FMT_INT
        elif c_idx == 11:
            cell.alignment = ALIGN_RIGHT
            cell.number_format = FMT_INT_DIFF
            if isinstance(v, (int, float)):
                if v > 0: cell.font = FONT_POS
                elif v < 0: cell.font = FONT_NEG
        elif c_idx in [9, 12]:
            cell.alignment = ALIGN_LEFT
            if v != '-':
                cell.font = FONT_BOLD
                
    ws3.row_dimensions[r3_idx].height = 20
    r3_idx += 1

# -------------------------------------------------------------------------
# HOJAS POR COMERCIAL (Alfonso, Garcia, Javier, Pedro, Ricardo)
# -------------------------------------------------------------------------
for com in comerciales:
    ws_com = wb.create_sheet(title=f'Detalle_{com}')
    ws_com.views.sheetView[0].showGridLines = True
    
    ws_com.merge_cells('B2:S2')
    ws_com['B2'] = f'CODIAGRO · REVISIÓN DE PREVISIONES Q4 - DELEGADO: {com.upper()}'
    ws_com['B2'].font = FONT_TITLE
    ws_com['B2'].fill = FILL_NAVY
    ws_com['B2'].alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws_com.row_dimensions[2].height = 32
    
    records = data_by_com.get(com, [])
    mod_count = sum(1 for r in records if r['has_qty_change'])
    q4_ant_c = sum(r['q4_ant'] for r in records)
    q4_rev_c = sum(r['q4_rev'] for r in records)
    q4_dif_c = q4_rev_c - q4_ant_c
    
    ws_com.merge_cells('B3:S3')
    ws_com['B3'] = f'Líneas Totales: {len(records)} | Líneas Modificadas: {mod_count} | Previsión Inicial Q4: {q4_ant_c:,.0f} u | Previsión Revisada Q4: {q4_rev_c:,.0f} u | Dif Neta: {q4_dif_c:+,.0f} u ({q4_dif_c/q4_ant_c:+.2%})'
    ws_com['B3'].font = FONT_SUBTITLE
    ws_com['B3'].fill = PatternFill(start_color='1E293B', end_color='1E293B', fill_type='solid')
    ws_com['B3'].alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws_com.row_dimensions[3].height = 20
    
    headers_com = [
        'País', 'Cliente', 'Cód. SKU', 'Descripción Artículo',
        'Sep Prev (u)', 'Sep Real (u)', 'Comentario Sep',
        'Oct Ant (u)', 'Oct Rev (u)', 'Dif Oct (u)',
        'Nov Ant (u)', 'Nov Rev (u)', 'Dif Nov (u)',
        'Dic Ant (u)', 'Dic Rev (u)', 'Dic Dif (u)',
        'Total Q4 Ant', 'Total Q4 Rev', 'Dif Q4 (u)', 'Comentario / Justificación Q4'
    ]
    
    ws_com.row_dimensions[5].height = 26
    for idx, h in enumerate(headers_com, start=2):
        cell = ws_com.cell(5, idx, h)
        cell.font = FONT_TH
        cell.fill = FILL_HEADER
        cell.alignment = ALIGN_HDR
        cell.border = BORDER_HEADER
        
    c_r = 6
    for rec in records:
        is_mod = rec['has_qty_change']
        vals = [
            rec['pais'], rec['cliente'], rec['sku'], rec['desc'],
            rec['sep_prev'], rec['sep_real'], rec['com_sep'] or '-',
            rec['oct_ant'], rec['oct_rev'], rec['oct_dif'],
            rec['nov_ant'], rec['nov_rev'], rec['nov_dif'],
            rec['dic_ant'], rec['dic_rev'], rec['dic_dif'],
            rec['q4_ant'], rec['q4_rev'], rec['q4_dif'], rec['com_q4'] or '-'
        ]
        
        for c_idx, v in enumerate(vals, start=2):
            cell = ws_com.cell(c_r, c_idx, v)
            cell.font = FONT_BOLD if is_mod else FONT_DATA
            cell.border = BORDER_THIN
            
            # Fondo resaltado si está modificado
            if is_mod:
                cell.fill = PatternFill(start_color='FEF9C3', end_color='FEF9C3', fill_type='solid') # Amarillo suave
            else:
                cell.fill = FILL_ZEBRA if c_r % 2 == 0 else PatternFill(fill_type=None)
                
            if c_idx in [2, 3, 4, 5]:
                cell.alignment = ALIGN_LEFT
            elif c_idx in [6, 7]:
                cell.alignment = ALIGN_RIGHT
                cell.number_format = FMT_INT
            elif c_idx == 8:
                cell.alignment = ALIGN_LEFT
            elif c_idx in [9, 10, 12, 13, 15, 16, 18, 19]:
                cell.alignment = ALIGN_RIGHT
                cell.number_format = FMT_INT
            elif c_idx in [11, 14, 17, 20]:
                cell.alignment = ALIGN_RIGHT
                cell.number_format = FMT_INT_DIFF
                if isinstance(v, (int, float)):
                    if v > 0: cell.font = FONT_POS
                    elif v < 0: cell.font = FONT_NEG
            elif c_idx == 21:
                cell.alignment = ALIGN_LEFT
                
        ws_com.row_dimensions[c_r].height = 20
        c_r += 1
        
    # Total comercial
    ws_com.cell(c_r, 2, 'TOTAL').font = FONT_TOTAL
    for c_idx in range(2, len(headers_com) + 2):
        cell = ws_com.cell(c_r, c_idx)
        cell.fill = FILL_TOTAL
        cell.border = BORDER_TOTAL
    ws_com.cell(c_r, 9, sum(r['oct_ant'] for r in records)).number_format = FMT_INT
    ws_com.cell(c_r, 10, sum(r['oct_rev'] for r in records)).number_format = FMT_INT
    ws_com.cell(c_r, 11, sum(r['oct_dif'] for r in records)).number_format = FMT_INT_DIFF
    ws_com.cell(c_r, 12, sum(r['nov_ant'] for r in records)).number_format = FMT_INT
    ws_com.cell(c_r, 13, sum(r['nov_rev'] for r in records)).number_format = FMT_INT
    ws_com.cell(c_r, 14, sum(r['nov_dif'] for r in records)).number_format = FMT_INT_DIFF
    ws_com.cell(c_r, 15, sum(r['dic_ant'] for r in records)).number_format = FMT_INT
    ws_com.cell(c_r, 16, sum(r['dic_rev'] for r in records)).number_format = FMT_INT
    ws_com.cell(c_r, 17, sum(r['dic_dif'] for r in records)).number_format = FMT_INT_DIFF
    ws_com.cell(c_r, 18, q4_ant_c).number_format = FMT_INT
    ws_com.cell(c_r, 19, q4_rev_c).number_format = FMT_INT
    ws_com.cell(c_r, 20, q4_dif_c).number_format = FMT_INT_DIFF
    ws_com.row_dimensions[c_r].height = 24

# Autoajustar anchos de columna para todas las hojas
for ws in wb.worksheets:
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            # Ignorar filas de títulos banner
            if cell.row in [2, 3]: continue
            val = str(cell.value or '')
            if len(val) > max_len:
                max_len = len(val)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 10)
    # Columna A margen estrecho
    ws.column_dimensions['A'].width = 3

# Guardar en ambas ubicaciones
out_p1 = 'Informe_Cambios_Previsiones_Q4_Comerciales.xlsx'
out_p2 = 'Revision_Previsiones_Q4_2026/Informe_Cambios_Previsiones_Q4_Comerciales.xlsx'

wb.save(out_p1)
wb.save(out_p2)

print(f"¡Libro Excel generado y guardado con éxito!")
print(f"  -> {out_p1}")
print(f"  -> {out_p2}")
