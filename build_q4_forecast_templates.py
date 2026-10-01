import os
import re
import shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
import pandas as pd
import unicodedata
from datetime import datetime

base_dir = r"c:\Users\oscar.ocampo\OneDrive - Agroquímica Codiagro S.L\Escritorio\Presupuesto 2026\Calsif Comerciales"
os.chdir(base_dir)

def norm(s):
    if not isinstance(s, str): return ''
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn').upper().strip()

def map_client_name(c):
    if not isinstance(c, str): return c
    cn = norm(c)
    if 'DOGA TARIM' in cn:
        return 'ÇİTAR ÇİÇEK TARIM TIC. LTD. STI'
    return c

print("1. Cargando datos de presupuestos y pedidos reales...")

# 1. Cargar Budget original sin aplanar
src_budget = 'Presupuesto_Ventas_2027_Original_Sin_Aplanar.xlsx' if os.path.exists('Presupuesto_Ventas_2027_Original_Sin_Aplanar.xlsx') else 'Presupuesto_Ventas_2027.xlsx'
shutil.copy2(src_budget, 'temp_budget_input.xlsx')
df_bud = pd.read_excel('temp_budget_input.xlsx', sheet_name='Previsión Matriz Horizontal')
df_bud = df_bud[df_bud['Comercial'].notna() & (~df_bud['Comercial'].astype(str).str.contains('TOTAL', case=False))].copy()
df_bud['Cliente'] = df_bud['Cliente'].apply(map_client_name)
df_bud['CLIENTE_NORM'] = df_bud['Cliente'].apply(norm)

def clean_budget_sku(row):
    sku = str(row['Código Artículo']).strip().upper()
    pais = str(row['País']).strip().upper()
    if pais != 'ESPAÑA' and len(sku) > 2:
        return sku[:-2]
    return sku

df_bud['SKU_CLEAN'] = df_bud.apply(clean_budget_sku, axis=1)

client_map = df_bud.groupby('CLIENTE_NORM')['Comercial'].first().to_dict()
client_map[norm('FINCA DOÑA ANA C.B.')] = 'Javier'
client_map[norm('FITOSANITARIOS CARCAIXENT, S.L.')] = 'Javier'
client_map[norm('ALMENDRALIA IBÉRICA, S.L.U.')] = 'Javier'
client_map[norm('TÉCNICAS AGRÍCOLAS, S.A.')] = 'Ricardo'

# 2. Cargar Pedidos reales de septiembre
def get_latest_pedidos_file():
    candidates = [
        f for f in os.listdir('.')
        if f.startswith('Pedidos') and f.endswith('.xlsx') and not f.startswith(('temp_', '~$'))
    ]
    if not candidates:
        return 'Pedidos 01.10 comerciales.xlsx'
    def sort_key(f):
        m = re.search(r'(\d{1,2})\.(\d{1,2})', f)
        if m:
            return (int(m.group(2)), int(m.group(1)), os.path.getmtime(f))
        return (0, 0, os.path.getmtime(f))
    candidates.sort(key=sort_key, reverse=True)
    return candidates[0]

pedidos_file = get_latest_pedidos_file()
print(f"Usando pedidos de cierre: {pedidos_file}")
shutil.copy2(pedidos_file, 'temp_pedidos_input.xlsx')
df_ped = pd.read_excel('temp_pedidos_input.xlsx')
df_ped['Cliente'] = df_ped['Cliente'].apply(map_client_name)
df_ped['CLIENTE_NORM'] = df_ped['Cliente'].apply(norm)
df_ped['SKU_CLEAN'] = df_ped['CodigoArticulo'].astype(str).str.strip().str.upper()
df_ped = df_ped[~df_ped['CLIENTE_NORM'].str.contains('SUSTAINABLE', case=False, na=False)].copy()
df_ped['Comercial'] = df_ped['CLIENTE_NORM'].map(client_map).fillna('Sin Asignar')

comerciales = ['Alfonso', 'García', 'Irene', 'Javier', 'Mehmet', 'Pedro', 'Ricardo']

# Estilos Openpyxl
font_title = Font(name='Segoe UI', size=15, bold=True, color='FFFFFF')
font_subtitle = Font(name='Segoe UI', size=10, italic=True, color='94A3B8')
font_instructions = Font(name='Segoe UI', size=9, bold=True, color='1E293B')

font_kpi_lbl = Font(name='Segoe UI', size=9, bold=True, color='64748B')
font_group_hdr = Font(name='Segoe UI', size=10, bold=True, color='FFFFFF')
font_col_hdr = Font(name='Segoe UI', size=9, bold=True, color='FFFFFF')
font_regular = Font(name='Segoe UI', size=9, color='1E293B')
font_bold = Font(name='Segoe UI', size=9, bold=True, color='0F172A')
font_editable = Font(name='Segoe UI', size=9, bold=True, color='065F46')
font_total = Font(name='Segoe UI', size=10, bold=True, color='0F172A')

fill_dark = PatternFill(start_color='0F172A', end_color='0F172A', fill_type='solid')
fill_subbanner = PatternFill(start_color='1E293B', end_color='1E293B', fill_type='solid')
fill_instructions = PatternFill(start_color='EFF6FF', end_color='EFF6FF', fill_type='solid')

fill_grp_info = PatternFill(start_color='334155', end_color='334155', fill_type='solid')
fill_grp_sep = PatternFill(start_color='1D4ED8', end_color='1D4ED8', fill_type='solid')
fill_grp_q4 = PatternFill(start_color='047857', end_color='047857', fill_type='solid')
fill_grp_tot = PatternFill(start_color='475569', end_color='475569', fill_type='solid')

fill_col_hdr = PatternFill(start_color='1E293B', end_color='1E293B', fill_type='solid')
fill_col_sep = PatternFill(start_color='2563EB', end_color='2563EB', fill_type='solid')
fill_col_q4 = PatternFill(start_color='059669', end_color='059669', fill_type='solid')

fill_card_kpi = PatternFill(start_color='F1F5F9', end_color='F1F5F9', fill_type='solid')
fill_card_accent = PatternFill(start_color='ECFDF5', end_color='ECFDF5', fill_type='solid')
fill_card_blue = PatternFill(start_color='EFF6FF', end_color='EFF6FF', fill_type='solid')
fill_card_warn = PatternFill(start_color='FEF3C7', end_color='FEF3C7', fill_type='solid')

fill_prev_hist = PatternFill(start_color='F8FAFC', end_color='F8FAFC', fill_type='solid')
fill_editable = PatternFill(start_color='ECFDF5', end_color='ECFDF5', fill_type='solid')
fill_comment_sep = PatternFill(start_color='FEF9C3', end_color='FEF9C3', fill_type='solid')
fill_comment_q4 = PatternFill(start_color='F8FAFC', end_color='F8FAFC', fill_type='solid')
fill_total_row = PatternFill(start_color='E2E8F0', end_color='E2E8F0', fill_type='solid')

border_thin = Border(
    left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
    top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1')
)
border_editable = Border(
    left=Side(style='thin', color='10B981'), right=Side(style='thin', color='10B981'),
    top=Side(style='thin', color='10B981'), bottom=Side(style='thin', color='10B981')
)
border_double = Border(
    top=Side(style='thin', color='94A3B8'), bottom=Side(style='double', color='0F172A'),
    left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1')
)

align_center = Alignment(horizontal='center', vertical='center')
align_left = Alignment(horizontal='left', vertical='center')
align_right = Alignment(horizontal='right', vertical='center')

# Validación de Comentarios Sep
dv_coment_sep = DataValidation(
    type="list",
    formula1='"Adelanto de pedidos,Atraso de pedidos,Pedido extra no previsto,Campaña habitual,Cancelado / Sin siembra,Otro motivo"',
    allow_blank=True
)
dv_coment_sep.error = 'Por favor seleccione un motivo de la lista o escriba su comentario'
dv_coment_sep.errorTitle = 'Motivo de Desviación'

out_dir = "Revision_Previsiones_Q4_2026"
os.makedirs(out_dir, exist_ok=True)

def populate_sheet_commercial(ws, c, is_individual=False):
    ws.views.sheetView[0].showGridLines = True
    
    sub_ped = df_ped[df_ped['Comercial'] == c]
    sub_bud = df_bud[df_bud['Comercial'] == c]
    
    # 1. Título y Banner
    ws.merge_cells('B2:T2')
    ws['B2'] = f"CODIAGRO · ACTUALIZACIÓN DE PREVISIONES COMERCIALES Q4 2026 (OCTUBRE - DICIEMBRE)"
    ws['B2'].font = font_title
    ws['B2'].fill = fill_dark
    ws['B2'].alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.row_dimensions[2].height = 36
    
    ws.merge_cells('B3:T3')
    ws['B3'] = f"DELEGADO COMERCIAL: {c.upper()} | Cierre Septiembre Oficial (01/10/2026) vs Previsiones a Final de Año | Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    ws['B3'].font = font_subtitle
    ws['B3'].fill = fill_subbanner
    ws['B3'].alignment = Alignment(horizontal='left', vertical='center', indent=1)
    ws.row_dimensions[3].height = 20
    
    # 2. Caja de Instrucciones
    ws.merge_cells('B4:T4')
    ws['B4'] = "📌 INSTRUCCIONES: 1) Revise el bloque de SEPTIEMBRE (Previsto, Real, Desviación en Valor Absoluto y % Consecución) e indique en col. J el motivo de la desviación (adelanto, atraso, extra). 2) En las columnas verdes (L, N, P), revise y actualice las cantidades estimadas de Octubre, Noviembre y Diciembre si prevé cambios. 3) Comente en col. T cualquier nota de campaña."
    ws['B4'].font = font_instructions
    ws['B4'].fill = fill_instructions
    ws['B4'].alignment = Alignment(horizontal='left', vertical='center', indent=1, wrap_text=True)
    ws.row_dimensions[4].height = 26
    
    ws.row_dimensions[5].height = 16
    ws.row_dimensions[6].height = 26
    
    # Construcción de datos por línea
    keys_set = set()
    for _, r in sub_bud.iterrows():
        keys_set.add((r['Cliente'], r['SKU_CLEAN']))
    for _, r in sub_ped.iterrows():
        keys_set.add((r['Cliente'], r['SKU_CLEAN']))
        
    keys_sorted = sorted(list(keys_set), key=lambda x: (x[0], x[1]))
    
    # Cabeceras Agrupadas (Fila 8)
    ws.merge_cells('B8:E8')
    ws['B8'] = "DATOS CLIENTE Y ARTÍCULO"
    ws['B8'].font = font_group_hdr
    ws['B8'].fill = fill_grp_info
    ws['B8'].alignment = align_center
    
    ws.merge_cells('F8:J8')
    ws['F8'] = "CIERRE SEPTIEMBRE 2026 (CALIBRACIÓN Y ANÁLISIS)"
    ws['F8'].font = font_group_hdr
    ws['F8'].fill = fill_grp_sep
    ws['F8'].alignment = align_center
    
    ws.merge_cells('K8:P8')
    ws['K8'] = "REVISIÓN PREVISIONES MES A MES (HASTA FINAL DE AÑO 2026)"
    ws['K8'].font = font_group_hdr
    ws['K8'].fill = fill_grp_q4
    ws['K8'].alignment = align_center
    
    ws.merge_cells('Q8:T8')
    ws['Q8'] = "IMPACTO Q4 Y JUSTIFICACIÓN"
    ws['Q8'].font = font_group_hdr
    ws['Q8'].fill = fill_grp_tot
    ws['Q8'].alignment = align_center
    
    ws.row_dimensions[8].height = 22
    
    # Cabeceras Columnas (Fila 9)
    col_headers = [
        ('B', 'País', fill_col_hdr),
        ('C', 'Cliente', fill_col_hdr),
        ('D', 'Cód. SKU', fill_col_hdr),
        ('E', 'Descripción Artículo', fill_col_hdr),
        ('F', 'Sep Previsto (u)', fill_col_sep),
        ('G', 'Sep Real (u)', fill_col_sep),
        ('H', 'Desv. Sep (u)', fill_col_sep),
        ('I', '% Consec.', fill_col_sep),
        ('J', 'Comentario Sep (¿Adelanto/Atraso/Extra?)', fill_col_sep),
        ('K', 'Oct Anterior (u)', fill_col_q4),
        ('L', 'Oct Revisado (u)', fill_col_q4),
        ('M', 'Nov Anterior (u)', fill_col_q4),
        ('N', 'Nov Revisado (u)', fill_col_q4),
        ('O', 'Dic Anterior (u)', fill_col_q4),
        ('P', 'Dic Revisado (u)', fill_col_q4),
        ('Q', 'Total Q4 Ant (u)', fill_col_hdr),
        ('R', 'Total Q4 Rev (u)', fill_col_hdr),
        ('S', 'Dif Q4 (u)', fill_col_hdr),
        ('T', 'Comentarios / Justificación Q4', fill_col_hdr)
    ]
    
    for c_letter, h_text, fill_c in col_headers:
        c_cell = ws[f"{c_letter}9"]
        c_cell.value = h_text
        c_cell.font = font_col_hdr
        c_cell.fill = fill_c
        c_cell.alignment = align_center
        c_cell.border = border_thin
    ws.row_dimensions[9].height = 28
    
    row_num = 10
    start_data_row = 10
    
    for cli, sku in keys_sorted:
        bm = sub_bud[(sub_bud['Cliente'] == cli) & (sub_bud['SKU_CLEAN'] == sku)]
        pm = sub_ped[(sub_ped['Cliente'] == cli) & (sub_ped['SKU_CLEAN'] == sku)]
        
        desc = ""
        pais = "ESPAÑA"
        if len(bm) > 0:
            if pd.notna(bm['Descripción Artículo'].iloc[0]): desc = str(bm['Descripción Artículo'].iloc[0])
            if pd.notna(bm['País'].iloc[0]): pais = str(bm['País'].iloc[0])
        elif len(pm) > 0 and pd.notna(pm['Articulo'].iloc[0]):
            desc = str(pm['Articulo'].iloc[0])
            
        s_b = float(bm['Sep-26 (u)'].sum()) if len(bm) > 0 else 0.0
        s_r = float(pm['Pedidas'].sum()) if len(pm) > 0 else 0.0
        
        oct_ant = float(bm['Oct-26 (u)'].sum()) if len(bm) > 0 else 0.0
        nov_ant = float(bm['Nov-26 (u)'].sum()) if len(bm) > 0 else 0.0
        dic_ant = float(bm['Dic-26 (u)'].sum()) if len(bm) > 0 else 0.0
        
        # Omitir líneas con todo a cero
        if s_b == 0 and s_r == 0 and oct_ant == 0 and nov_ant == 0 and dic_ant == 0:
            continue
            
        r_str = str(row_num)
        
        # B: País
        ws[f"B{r_str}"] = pais
        ws[f"B{r_str}"].font = font_regular
        ws[f"B{r_str}"].alignment = align_center
        ws[f"B{r_str}"].border = border_thin
        
        # C: Cliente
        ws[f"C{r_str}"] = cli
        ws[f"C{r_str}"].font = font_regular
        ws[f"C{r_str}"].alignment = align_left
        ws[f"C{r_str}"].border = border_thin
        
        # D: SKU
        ws[f"D{r_str}"] = sku
        ws[f"D{r_str}"].font = font_regular
        ws[f"D{r_str}"].alignment = align_center
        ws[f"D{r_str}"].border = border_thin
        
        # E: Desc
        ws[f"E{r_str}"] = desc
        ws[f"E{r_str}"].font = font_regular
        ws[f"E{r_str}"].alignment = align_left
        ws[f"E{r_str}"].border = border_thin
        
        # F: Sep Previsto (u)
        ws[f"F{r_str}"] = s_b
        ws[f"F{r_str}"].font = font_regular
        ws[f"F{r_str}"].alignment = align_right
        ws[f"F{r_str}"].number_format = '#,##0'
        ws[f"F{r_str}"].border = border_thin
        
        # G: Sep Real (u)
        ws[f"G{r_str}"] = s_r
        ws[f"G{r_str}"].font = font_bold
        ws[f"G{r_str}"].alignment = align_right
        ws[f"G{r_str}"].number_format = '#,##0'
        ws[f"G{r_str}"].border = border_thin
        
        # H: Desv Sep (u) en valor absoluto = |Real - Previsto|
        ws[f"H{r_str}"] = f"=ABS(G{r_str}-F{r_str})"
        ws[f"H{r_str}"].font = font_regular
        ws[f"H{r_str}"].alignment = align_right
        ws[f"H{r_str}"].number_format = '#,##0'
        ws[f"H{r_str}"].border = border_thin

        # I: % Consecución = Real / Previsto
        ws[f"I{r_str}"] = f'=IF(F{r_str}>0, G{r_str}/F{r_str}, IF(G{r_str}>0, 1.0, 0.0))'
        ws[f"I{r_str}"].font = font_regular
        ws[f"I{r_str}"].alignment = align_right
        ws[f"I{r_str}"].number_format = '0.0%'
        ws[f"I{r_str}"].border = border_thin
        
        # J: Comentario Sep
        ws[f"J{r_str}"] = "Pedido extra no previsto" if (s_b == 0 and s_r > 0) else ""
        ws[f"J{r_str}"].font = font_regular
        ws[f"J{r_str}"].fill = fill_comment_sep
        ws[f"J{r_str}"].alignment = align_left
        ws[f"J{r_str}"].border = border_thin
        
        # K: Oct Ant
        ws[f"K{r_str}"] = oct_ant
        ws[f"K{r_str}"].font = font_regular
        ws[f"K{r_str}"].fill = fill_prev_hist
        ws[f"K{r_str}"].alignment = align_right
        ws[f"K{r_str}"].number_format = '#,##0'
        ws[f"K{r_str}"].border = border_thin
        
        # L: Oct Rev (precargado con oct_ant)
        ws[f"L{r_str}"] = oct_ant
        ws[f"L{r_str}"].font = font_editable
        ws[f"L{r_str}"].fill = fill_editable
        ws[f"L{r_str}"].alignment = align_right
        ws[f"L{r_str}"].number_format = '#,##0'
        ws[f"L{r_str}"].border = border_editable
        
        # M: Nov Ant
        ws[f"M{r_str}"] = nov_ant
        ws[f"M{r_str}"].font = font_regular
        ws[f"M{r_str}"].fill = fill_prev_hist
        ws[f"M{r_str}"].alignment = align_right
        ws[f"M{r_str}"].number_format = '#,##0'
        ws[f"M{r_str}"].border = border_thin
        
        # N: Nov Rev
        ws[f"N{r_str}"] = nov_ant
        ws[f"N{r_str}"].font = font_editable
        ws[f"N{r_str}"].fill = fill_editable
        ws[f"N{r_str}"].alignment = align_right
        ws[f"N{r_str}"].number_format = '#,##0'
        ws[f"N{r_str}"].border = border_editable
        
        # O: Dic Ant
        ws[f"O{r_str}"] = dic_ant
        ws[f"O{r_str}"].font = font_regular
        ws[f"O{r_str}"].fill = fill_prev_hist
        ws[f"O{r_str}"].alignment = align_right
        ws[f"O{r_str}"].number_format = '#,##0'
        ws[f"O{r_str}"].border = border_thin
        
        # P: Dic Rev
        ws[f"P{r_str}"] = dic_ant
        ws[f"P{r_str}"].font = font_editable
        ws[f"P{r_str}"].fill = fill_editable
        ws[f"P{r_str}"].alignment = align_right
        ws[f"P{r_str}"].number_format = '#,##0'
        ws[f"P{r_str}"].border = border_editable
        
        # Q: Total Q4 Ant = K + M + O
        ws[f"Q{r_str}"] = f"=K{r_str}+M{r_str}+O{r_str}"
        ws[f"Q{r_str}"].font = font_regular
        ws[f"Q{r_str}"].alignment = align_right
        ws[f"Q{r_str}"].number_format = '#,##0'
        ws[f"Q{r_str}"].border = border_thin
        
        # R: Total Q4 Rev = L + N + P
        ws[f"R{r_str}"] = f"=L{r_str}+N{r_str}+P{r_str}"
        ws[f"R{r_str}"].font = font_bold
        ws[f"R{r_str}"].alignment = align_right
        ws[f"R{r_str}"].number_format = '#,##0'
        ws[f"R{r_str}"].border = border_thin
        
        # S: Dif Q4 = R - Q
        ws[f"S{r_str}"] = f"=R{r_str}-Q{r_str}"
        ws[f"S{r_str}"].font = font_regular
        ws[f"S{r_str}"].alignment = align_right
        ws[f"S{r_str}"].number_format = '+#,##0;-#,##0;0'
        ws[f"S{r_str}"].border = border_thin
        
        # T: Comentario Q4
        ws[f"T{r_str}"] = ""
        ws[f"T{r_str}"].font = font_regular
        ws[f"T{r_str}"].fill = fill_comment_q4
        ws[f"T{r_str}"].alignment = align_left
        ws[f"T{r_str}"].border = border_thin
        
        ws.row_dimensions[row_num].height = 20
        row_num += 1
        
    last_data_row = row_num - 1
    
    # Fila de Totales
    tot_row = row_num
    tr_str = str(tot_row)
    
    ws.merge_cells(f"B{tr_str}:E{tr_str}")
    ws[f"B{tr_str}"] = f"TOTAL CONSOLIDADO - {c.upper()} (UNIDADES)"
    ws[f"B{tr_str}"].font = font_total
    ws[f"B{tr_str}"].fill = fill_total_row
    ws[f"B{tr_str}"].alignment = Alignment(horizontal='right', vertical='center', indent=1)
    
    for c_let in ['B', 'C', 'D', 'E']:
        ws[f"{c_let}{tr_str}"].border = border_double
        
    # Sumas de cantidades (F: Prev, G: Real, H: Desv en valor absoluto, K, L, M, N, O, P, Q, R, S)
    for c_let in ['F', 'G', 'H', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S']:
        ws[f"{c_let}{tr_str}"] = f"=SUM({c_let}{start_data_row}:{c_let}{last_data_row})"
        ws[f"{c_let}{tr_str}"].font = font_total
        ws[f"{c_let}{tr_str}"].fill = fill_total_row
        ws[f"{c_let}{tr_str}"].alignment = align_right
        ws[f"{c_let}{tr_str}"].border = border_double
        ws[f"{c_let}{tr_str}"].number_format = '+#,##0;-#,##0;0' if c_let == 'S' else '#,##0'
        
    # Consecución Total
    ws[f"I{tr_str}"] = f'=IF(F{tr_str}>0, G{tr_str}/F{tr_str}, 1.0)'
    ws[f"I{tr_str}"].font = font_total
    ws[f"I{tr_str}"].fill = fill_total_row
    ws[f"I{tr_str}"].alignment = align_right
    ws[f"I{tr_str}"].number_format = '0.0%'
    ws[f"I{tr_str}"].border = border_double
    
    # Vacíos para comentarios en total
    for c_let in ['J', 'T']:
        ws[f"{c_let}{tr_str}"].border = border_double
        ws[f"{c_let}{tr_str}"].fill = fill_total_row
        
    ws.row_dimensions[tot_row].height = 25
    
    # Llenar fórmulas en KPI Cards de cabecera
    # B:D -> F_tot
    ws.merge_cells("B5:D5")
    ws.merge_cells("B6:D6")
    ws['B5'] = "PREVISTO SEP 2026 (U)"
    ws['B5'].font = font_kpi_lbl
    ws['B5'].fill = fill_card_blue
    ws['B5'].alignment = align_center
    ws['B6'] = f"=F{tr_str}"
    ws['B6'].font = Font(name='Segoe UI', size=13, bold=True, color='1D4ED8')
    ws['B6'].fill = fill_card_blue
    ws['B6'].alignment = align_center
    ws['B6'].number_format = '#,##0 "u"'
    
    # E:G -> G_tot
    ws.merge_cells("E5:G5")
    ws.merge_cells("E6:G6")
    ws['E5'] = "REAL PEDIDOS SEP 2026 (U)"
    ws['E5'].font = font_kpi_lbl
    ws['E5'].fill = fill_card_accent
    ws['E5'].alignment = align_center
    ws['E6'] = f"=G{tr_str}"
    ws['E6'].font = Font(name='Segoe UI', size=13, bold=True, color='047857')
    ws['E6'].fill = fill_card_accent
    ws['E6'].alignment = align_center
    ws['E6'].number_format = '#,##0 "u"'
    
    # H:J -> I_tot (% Consecución)
    ws.merge_cells("H5:J5")
    ws.merge_cells("H6:J6")
    ws['H5'] = "% CONSECUCIÓN SEP"
    ws['H5'].font = font_kpi_lbl
    ws['H5'].fill = fill_card_kpi
    ws['H5'].alignment = align_center
    ws['H6'] = f"=I{tr_str}"
    ws['H6'].font = Font(name='Segoe UI', size=13, bold=True, color='0F172A')
    ws['H6'].fill = fill_card_kpi
    ws['H6'].alignment = align_center
    ws['H6'].number_format = '0.0%'
    
    # K:M -> Q_tot (Total Q4 Ant)
    ws.merge_cells("K5:M5")
    ws.merge_cells("K6:M6")
    ws['K5'] = "PREVISIÓN INICIAL Q4 (U)"
    ws['K5'].font = font_kpi_lbl
    ws['K5'].fill = fill_card_blue
    ws['K5'].alignment = align_center
    ws['K6'] = f"=Q{tr_str}"
    ws['K6'].font = Font(name='Segoe UI', size=13, bold=True, color='1D4ED8')
    ws['K6'].fill = fill_card_blue
    ws['K6'].alignment = align_center
    ws['K6'].number_format = '#,##0 "u"'
    
    # N:P -> R_tot (Total Q4 Rev)
    ws.merge_cells("N5:P5")
    ws.merge_cells("N6:P6")
    ws['N5'] = "PREVISIÓN REVISADA Q4 (U)"
    ws['N5'].font = font_kpi_lbl
    ws['N5'].fill = fill_card_accent
    ws['N5'].alignment = align_center
    ws['N6'] = f"=R{tr_str}"
    ws['N6'].font = Font(name='Segoe UI', size=13, bold=True, color='047857')
    ws['N6'].fill = fill_card_accent
    ws['N6'].alignment = align_center
    ws['N6'].number_format = '#,##0 "u"'
    
    # Q:T -> S_tot (Diferencia Q4)
    ws.merge_cells("Q5:T5")
    ws.merge_cells("Q6:T6")
    ws['Q5'] = "VARIACIÓN NETA Q4 (U)"
    ws['Q5'].font = font_kpi_lbl
    ws['Q5'].fill = fill_card_warn
    ws['Q5'].alignment = align_center
    ws['Q6'] = f"=S{tr_str}"
    ws['Q6'].font = Font(name='Segoe UI', size=13, bold=True, color='B45309')
    ws['Q6'].fill = fill_card_warn
    ws['Q6'].alignment = align_center
    ws['Q6'].number_format = '+#,##0 "u";-#,##0 "u";0 "u"'
    
    for c1, c2 in [('B','D'), ('E','G'), ('H','J'), ('K','M'), ('N','P'), ('Q','T')]:
        start_c = ord(c1)
        end_c = ord(c2)
        for cc in range(start_c, end_c + 1):
            letter = chr(cc)
            ws[f"{letter}5"].border = border_thin
            ws[f"{letter}6"].border = border_thin
            
    # Añadir Data Validation a columna J (Comentario Sep)
    ws.add_data_validation(dv_coment_sep)
    dv_coment_sep.add(f"J{start_data_row}:J{last_data_row}")
    
    # Anchos de columna optimizados para visualización perfecta
    col_widths = {
        'A': 3,
        'B': 14, # País
        'C': 34, # Cliente
        'D': 14, # SKU
        'E': 32, # Desc
        'F': 16, # Sep Prev
        'G': 15, # Sep Real
        'H': 14, # Desv Sep
        'I': 13, # % Consec
        'J': 32, # Coment Sep
        'K': 15, # Oct Ant
        'L': 16, # Oct Rev
        'M': 15, # Nov Ant
        'N': 16, # Nov Rev
        'O': 15, # Dic Ant
        'P': 16, # Dic Rev
        'Q': 17, # Total Q4 Ant
        'R': 17, # Total Q4 Rev
        'S': 15, # Dif Q4
        'T': 34  # Coment Q4
    }
    for col_l, w in col_widths.items():
        ws.column_dimensions[col_l].width = w

print("2. Generando archivos individuales para cada comercial...")

for c in comerciales:
    wb_ind = openpyxl.Workbook()
    ws_ind = wb_ind.active
    ws_ind.title = f"Previsiones_{c}"
    populate_sheet_commercial(ws_ind, c, is_individual=True)
    
    safe_name = c.replace('í', 'i')
    filepath = os.path.join(out_dir, f"Revision_Previsiones_Q4_2026_{safe_name}.xlsx")
    wb_ind.save(filepath)
    print(f" -> Guardado: {filepath}")

print("3. Generando archivo CONSOLIDADO Master para Dirección...")

wb_con = openpyxl.Workbook()
wb_con.remove(wb_con.active)

# Hoja 1: Resumen General Consolidado
ws_resumen = wb_con.create_sheet(title="Resumen Consolidado Dirección")
ws_resumen.views.sheetView[0].showGridLines = True

ws_resumen.merge_cells('B2:M2')
ws_resumen['B2'] = "CODIAGRO · RESUMEN DIRECCIÓN: REVISIÓN DE PREVISIONES Q4 2026"
ws_resumen['B2'].font = font_title
ws_resumen['B2'].fill = fill_dark
ws_resumen['B2'].alignment = Alignment(horizontal='left', vertical='center', indent=1)
ws_resumen.row_dimensions[2].height = 36

ws_resumen.merge_cells('B2:L2')
ws_resumen['B2'] = "CODIAGRO · RESUMEN DIRECCIÓN: REVISIÓN DE PREVISIONES Q4 2026"
ws_resumen['B2'].font = font_title
ws_resumen['B2'].fill = fill_dark
ws_resumen['B2'].alignment = Alignment(horizontal='left', vertical='center', indent=1)
ws_resumen.row_dimensions[2].height = 36

ws_resumen.merge_cells('B3:L3')
ws_resumen['B3'] = f"Consolidado de Desviaciones de Septiembre y Previsiones hasta Final de Año (Octubre a Diciembre) | Cierre Sep: 01/10/2026"
ws_resumen['B3'].font = font_subtitle
ws_resumen['B3'].fill = fill_subbanner
ws_resumen['B3'].alignment = Alignment(horizontal='left', vertical='center', indent=1)
ws_resumen.row_dimensions[3].height = 20

# Resumen tabla cabeceras (11 columnas: B a L)
res_headers = [
    ('B', 'Comercial', fill_col_hdr),
    ('C', 'Clientes Activos', fill_col_hdr),
    ('D', 'Sep Previsto (u)', fill_col_sep),
    ('E', 'Sep Real (u)', fill_col_sep),
    ('F', 'Desv. Sep (u)', fill_col_sep),
    ('G', '% Consec.', fill_col_sep),
    ('H', 'Oct Previsto (u)', fill_col_q4),
    ('I', 'Nov Previsto (u)', fill_col_q4),
    ('J', 'Dic Previsto (u)', fill_col_q4),
    ('K', 'Total Q4 Previsto (u)', fill_col_q4),
    ('L', 'Estado Cierre Sep', fill_col_hdr)
]

ws_resumen.merge_cells('B5:C5')
ws_resumen['B5'] = "INFORMACIÓN COMERCIAL"
ws_resumen['B5'].font = font_group_hdr
ws_resumen['B5'].fill = fill_grp_info
ws_resumen['B5'].alignment = align_center

ws_resumen.merge_cells('D5:G5')
ws_resumen['D5'] = "CIERRE SEPTIEMBRE 2026 (REAL vs PREVISTO)"
ws_resumen['D5'].font = font_group_hdr
ws_resumen['D5'].fill = fill_grp_sep
ws_resumen['D5'].alignment = align_center

ws_resumen.merge_cells('H5:K5')
ws_resumen['H5'] = "PREVISIÓN CARTERA HASTA FINAL DE AÑO (Q4 2026)"
ws_resumen['H5'].font = font_group_hdr
ws_resumen['H5'].fill = fill_grp_q4
ws_resumen['H5'].alignment = align_center

ws_resumen.cell(5, 12).value = "VALIDACIÓN"
ws_resumen.cell(5, 12).font = font_group_hdr
ws_resumen.cell(5, 12).fill = fill_grp_tot
ws_resumen.cell(5, 12).alignment = align_center
ws_resumen.row_dimensions[5].height = 22

for c_l, h_t, f_c in res_headers:
    cell = ws_resumen[f"{c_l}6"]
    cell.value = h_t
    cell.font = font_col_hdr
    cell.fill = f_c
    cell.alignment = align_center
    cell.border = border_thin
ws_resumen.row_dimensions[6].height = 26

r_res = 7
for c in comerciales:
    s_c = f"='{c}'!"
    r_str = str(r_res)
    
    ws_resumen[f"B{r_str}"] = c
    ws_resumen[f"B{r_str}"].font = font_bold
    ws_resumen[f"B{r_str}"].alignment = align_left
    ws_resumen[f"B{r_str}"].border = border_thin
    
    # Clientes únicos
    n_cli = df_bud[df_bud['Comercial'] == c]['Cliente'].nunique()
    ws_resumen[f"C{r_str}"] = n_cli
    ws_resumen[f"C{r_str}"].font = font_regular
    ws_resumen[f"C{r_str}"].alignment = align_center
    ws_resumen[f"C{r_str}"].border = border_thin
    
    # Links a fórmulas del total de la pestaña
    sub_ped = df_ped[df_ped['Comercial'] == c]
    sub_bud = df_bud[df_bud['Comercial'] == c]
    keys_c = set()
    for _, r in sub_bud.iterrows(): keys_c.add((r['Cliente'], r['SKU_CLEAN']))
    for _, r in sub_ped.iterrows(): keys_c.add((r['Cliente'], r['SKU_CLEAN']))
    count_lines = 0
    for cli, sku in keys_c:
        bm = sub_bud[(sub_bud['Cliente'] == cli) & (sub_bud['SKU_CLEAN'] == sku)]
        pm = sub_ped[(sub_ped['Cliente'] == cli) & (sub_ped['SKU_CLEAN'] == sku)]
        sb = float(bm['Sep-26 (u)'].sum()) if len(bm) > 0 else 0
        sr = float(pm['Pedidas'].sum()) if len(pm) > 0 else 0
        o_ant = float(bm['Oct-26 (u)'].sum()) if len(bm) > 0 else 0
        n_ant = float(bm['Nov-26 (u)'].sum()) if len(bm) > 0 else 0
        d_ant = float(bm['Dic-26 (u)'].sum()) if len(bm) > 0 else 0
        if sb > 0 or sr > 0 or o_ant > 0 or n_ant > 0 or d_ant > 0:
            count_lines += 1
    t_row_sheet = 10 + count_lines
    
    # D: Sep Previsto (col F en hoja individual)
    ws_resumen[f"D{r_str}"] = f"{s_c}F{t_row_sheet}"
    ws_resumen[f"D{r_str}"].font = font_regular
    ws_resumen[f"D{r_str}"].alignment = align_right
    ws_resumen[f"D{r_str}"].number_format = '#,##0'
    ws_resumen[f"D{r_str}"].border = border_thin
    
    # E: Sep Real (col G en hoja individual)
    ws_resumen[f"E{r_str}"] = f"{s_c}G{t_row_sheet}"
    ws_resumen[f"E{r_str}"].font = font_bold
    ws_resumen[f"E{r_str}"].alignment = align_right
    ws_resumen[f"E{r_str}"].number_format = '#,##0'
    ws_resumen[f"E{r_str}"].border = border_thin
    
    # F: Desv Sep (u) en valor absoluto (col H en hoja individual)
    ws_resumen[f"F{r_str}"] = f"{s_c}H{t_row_sheet}"
    ws_resumen[f"F{r_str}"].font = font_regular
    ws_resumen[f"F{r_str}"].alignment = align_right
    ws_resumen[f"F{r_str}"].number_format = '#,##0'
    ws_resumen[f"F{r_str}"].border = border_thin
    
    # G: % Consec (col I en hoja individual)
    ws_resumen[f"G{r_str}"] = f"{s_c}I{t_row_sheet}"
    ws_resumen[f"G{r_str}"].font = font_bold
    ws_resumen[f"G{r_str}"].alignment = align_right
    ws_resumen[f"G{r_str}"].number_format = '0.0%'
    ws_resumen[f"G{r_str}"].border = border_thin
    
    # H: Oct Previsto (col L en hoja individual)
    ws_resumen[f"H{r_str}"] = f"{s_c}L{t_row_sheet}"
    ws_resumen[f"H{r_str}"].font = font_editable
    ws_resumen[f"H{r_str}"].alignment = align_right
    ws_resumen[f"H{r_str}"].number_format = '#,##0'
    ws_resumen[f"H{r_str}"].border = border_thin
    
    # I: Nov Previsto (col N en hoja individual)
    ws_resumen[f"I{r_str}"] = f"{s_c}N{t_row_sheet}"
    ws_resumen[f"I{r_str}"].font = font_editable
    ws_resumen[f"I{r_str}"].alignment = align_right
    ws_resumen[f"I{r_str}"].number_format = '#,##0'
    ws_resumen[f"I{r_str}"].border = border_thin
    
    # J: Dic Previsto (col P en hoja individual)
    ws_resumen[f"J{r_str}"] = f"{s_c}P{t_row_sheet}"
    ws_resumen[f"J{r_str}"].font = font_editable
    ws_resumen[f"J{r_str}"].alignment = align_right
    ws_resumen[f"J{r_str}"].number_format = '#,##0'
    ws_resumen[f"J{r_str}"].border = border_thin
    
    # K: Total Q4 Previsto (col R en hoja individual)
    ws_resumen[f"K{r_str}"] = f"{s_c}R{t_row_sheet}"
    ws_resumen[f"K{r_str}"].font = font_bold
    ws_resumen[f"K{r_str}"].alignment = align_right
    ws_resumen[f"K{r_str}"].number_format = '#,##0'
    ws_resumen[f"K{r_str}"].border = border_thin
    
    # L: Estado
    ws_resumen[f"L{r_str}"] = f'=IF(G{r_str}>=1.0, "Superado (+)", IF(G{r_str}>=0.7, "En Rango", "Revisar"))'
    ws_resumen[f"L{r_str}"].font = font_bold
    ws_resumen[f"L{r_str}"].alignment = align_center
    ws_resumen[f"L{r_str}"].border = border_thin
    
    ws_resumen.row_dimensions[r_res].height = 22
    r_res += 1

tot_r_res = str(r_res)
ws_resumen[f"B{tot_r_res}"] = "TOTAL CODIAGRO"
ws_resumen[f"B{tot_r_res}"].font = font_total
ws_resumen[f"B{tot_r_res}"].fill = fill_total_row
ws_resumen[f"B{tot_r_res}"].alignment = Alignment(horizontal='left', vertical='center', indent=1)

ws_resumen[f"C{tot_r_res}"] = f"=SUM(C7:C{r_res-1})"
ws_resumen[f"C{tot_r_res}"].font = font_total
ws_resumen[f"C{tot_r_res}"].fill = fill_total_row
ws_resumen[f"C{tot_r_res}"].alignment = align_center
ws_resumen[f"C{tot_r_res}"].border = border_double

# Sumas de Previsto (D), Real (E), Desv (F), Oct (H), Nov (I), Dic (J), Total Q4 (K)
for cl in ['D', 'E', 'F', 'H', 'I', 'J', 'K']:
    ws_resumen[f"{cl}{tot_r_res}"] = f"=SUM({cl}7:{cl}{r_res-1})"
    ws_resumen[f"{cl}{tot_r_res}"].font = font_total
    ws_resumen[f"{cl}{tot_r_res}"].fill = fill_total_row
    ws_resumen[f"{cl}{tot_r_res}"].alignment = align_right
    ws_resumen[f"{cl}{tot_r_res}"].border = border_double
    ws_resumen[f"{cl}{tot_r_res}"].number_format = '#,##0'

# Total Consecución = Real / Previsto
ws_resumen[f"G{tot_r_res}"] = f'=E{tot_r_res}/D{tot_r_res}'
ws_resumen[f"G{tot_r_res}"].font = font_total
ws_resumen[f"G{tot_r_res}"].fill = fill_total_row
ws_resumen[f"G{tot_r_res}"].alignment = align_right
ws_resumen[f"G{tot_r_res}"].border = border_double
ws_resumen[f"G{tot_r_res}"].number_format = '0.0%'

ws_resumen[f"L{tot_r_res}"] = "CONSOLIDADO"
ws_resumen[f"L{tot_r_res}"].font = font_total
ws_resumen[f"L{tot_r_res}"].fill = fill_total_row
ws_resumen[f"L{tot_r_res}"].alignment = align_center
ws_resumen[f"L{tot_r_res}"].border = border_double
ws_resumen[f"B{tot_r_res}"].border = border_double

ws_resumen.row_dimensions[r_res].height = 25

res_widths = {'A': 3, 'B': 18, 'C': 16, 'D': 17, 'E': 17, 'F': 16, 'G': 14, 'H': 16, 'I': 16, 'J': 16, 'K': 19, 'L': 18}
for c_l, w in res_widths.items():
    ws_resumen.column_dimensions[c_l].width = w

# Añadir hojas de cada comercial al consolidado
for c in comerciales:
    ws_c = wb_con.create_sheet(title=c)
    populate_sheet_commercial(ws_c, c, is_individual=False)

consolidado_path = os.path.join(out_dir, "Revision_Previsiones_Q4_2026_CONSOLIDADO.xlsx")
wb_con.save(consolidado_path)
print(f" -> Guardado Consolidado Master: {consolidado_path}")
print("¡Proceso completado con éxito!")
