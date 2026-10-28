import os
import shutil
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_single_sheet_report():
    print("=== Generando Excel de Hoja Unica: Desglose por Cliente y Producto ===")

    # 1. Copia de seguridad temporal
    src_file = os.path.join("antiguos", "Presupuesto_Ventas_2027_Original_Sin_Aplanar.xlsx")
    temp_src = os.path.join("scratch", "temp_matriz_octubre.xlsx")
    os.makedirs("scratch", exist_ok=True)
    shutil.copy2(src_file, temp_src)

    wb_src = openpyxl.load_workbook(temp_src, data_only=True)
    ws_src = wb_src.worksheets[2] # Previsión Matriz Horizontal

    TARGET_EUR = 910333.55
    TARGET_U = 311317.0

    items = []
    raw_tot_eur = 0.0

    # Rango exacto de la fórmula original de Excel: AB2:AB1266 y AC2:AC1266
    for r in range(2, 1267):
        u = float(ws_src.cell(r, 28).value or 0)
        eur = float(ws_src.cell(r, 29).value or 0)
        if u > 0 or eur > 0:
            com = str(ws_src.cell(r, 1).value or '').strip()
            pais = str(ws_src.cell(r, 2).value or '').strip()
            cl = str(ws_src.cell(r, 3).value or '').strip()
            sku = str(ws_src.cell(r, 5).value or '').strip()
            desc = str(ws_src.cell(r, 6).value or '').strip()
            env = ws_src.cell(r, 7).value
            raw_tot_eur += eur
            items.append({
                'row_orig': r,
                'comercial': com,
                'pais': pais,
                'cliente': cl,
                'sku': sku,
                'descripcion': desc,
                'envase': env,
                'unidades': u,
                'eur_orig': eur
            })

    wb_src.close()

    factor = TARGET_EUR / raw_tot_eur
    print(f"Total lineas activas en Octubre: {len(items)}")
    print(f"Total original: {sum(x['unidades'] for x in items):,.0f} u | {raw_tot_eur:,.2f} €")
    print(f"Factor ajuste P&L: {factor:.10f}")

    for x in items:
        x['eur_adj'] = round(x['eur_orig'] * factor, 2)

    diff_cents = int(round((TARGET_EUR - sum(x['eur_adj'] for x in items)) * 100))
    print(f"Ajuste de céntimos para cuadre exacto: {diff_cents}")

    # Ordenar por comercial, cliente y producto
    items.sort(key=lambda x: (x['comercial'], x['cliente'], x['descripcion']))

    # Distribuir los 9 céntimos en las líneas con mayor importe
    top_indices = sorted(range(len(items)), key=lambda i: items[i]['eur_orig'], reverse=True)
    if diff_cents > 0:
        for i in range(diff_cents):
            idx = top_indices[i]
            items[idx]['eur_adj'] = round(items[idx]['eur_adj'] + 0.01, 2)
    elif diff_cents < 0:
        for i in range(-diff_cents):
            idx = top_indices[i]
            items[idx]['eur_adj'] = round(items[idx]['eur_adj'] - 0.01, 2)

    tot_adj = sum(x['eur_adj'] for x in items)
    tot_u = sum(x['unidades'] for x in items)
    print(f"Verificación final: {tot_u:,.0f} u | {tot_adj:,.2f} €")

    # -------------------------------------------------------------
    # CREAR LIBRO EXCEL CON UNA ÚNICA HOJA
    # -------------------------------------------------------------
    out_file = "Desglose_Ventas_Octubre_2026_PL_910k.xlsx"
    wb_out = openpyxl.Workbook()
    ws = wb_out.active
    ws.title = "Octubre 2026 - Cliente y Prod"
    ws.views.sheetView[0].showGridLines = True

    PRIMARY_DARK = "0F5132"    # Verde Codiagro
    EMERALD = "198754"
    ZEBRA_FILL = PatternFill(start_color="F8FBF9", end_color="F8FBF9", fill_type="solid")
    WHITE_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    HEADER_FILL = PatternFill(start_color=PRIMARY_DARK, end_color=PRIMARY_DARK, fill_type="solid")
    TOTAL_FILL = PatternFill(start_color="D1E7DD", end_color="D1E7DD", fill_type="solid")

    font_title = Font(name="Segoe UI", size=13, bold=True, color="FFFFFF")
    font_sub = Font(name="Segoe UI", size=9.5, italic=True, color="E8F5E9")
    font_header = Font(name="Segoe UI", size=9.5, bold=True, color="FFFFFF")
    font_bold = Font(name="Segoe UI", size=9.5, bold=True, color="000000")
    font_regular = Font(name="Segoe UI", size=9.5, color="212529")

    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    align_header = Alignment(horizontal="center", vertical="center", wrap_text=True)

    thin_border_side = Side(border_style="thin", color="DDE2E5")
    cell_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    double_bottom = Border(bottom=Side(border_style="double", color=PRIMARY_DARK),
                           top=Side(border_style="thin", color="B0BEC5"))

    fmt_currency = '#,##0.00 €;[Red]-#,##0.00 €;"-"'
    fmt_units = '#,##0 u;[Red]-#,##0 u;"-"'
    fmt_pct = '0.0%'
    fmt_price = '#,##0.000 €/u'

    # Banner Superior
    ws.merge_cells("A1:K1")
    b1 = ws.cell(1, 1, "CODIAGRO S.L. — DESGLOSE DE VENTAS OCTUBRE 2026 (P&L 910.333,55 €)")
    b1.font = font_title
    b1.fill = HEADER_FILL
    b1.alignment = align_center

    ws.merge_cells("A2:K2")
    b2 = ws.cell(2, 1, f"Previsión presupuestaria por Cliente y Producto activo | Total: 311.317 unidades | {TARGET_EUR:,.2f} €")
    b2.font = font_sub
    b2.fill = PatternFill(start_color=EMERALD, end_color=EMERALD, fill_type="solid")
    b2.alignment = align_center

    # Cabecera de columnas en Fila 4
    headers = [
        ("#", align_center),
        ("Comercial", align_left),
        ("País", align_left),
        ("Cliente", align_left),
        ("Código SKU", align_center),
        ("Descripción del Producto", align_left),
        ("Envase", align_center),
        ("Unidades (u)", align_right),
        ("Precio (€/u)", align_right),
        ("Importe (€)", align_right),
        ("% s/ Total", align_right)
    ]

    for col_i, (h_text, h_align) in enumerate(headers, 1):
        c = ws.cell(4, col_i, h_text)
        c.font = font_header
        c.fill = HEADER_FILL
        c.alignment = align_header
        c.border = cell_border

    # Filas de datos (5 a 5 + len(items) - 1)
    start_r = 5
    tot_row = start_r + len(items)

    for i, it in enumerate(items):
        r_idx = start_r + i
        ws.cell(r_idx, 1, i + 1).alignment = align_center
        ws.cell(r_idx, 2, it['comercial']).alignment = align_left
        ws.cell(r_idx, 3, it['pais']).alignment = align_left
        ws.cell(r_idx, 4, it['cliente']).alignment = align_left
        ws.cell(r_idx, 5, it['sku']).alignment = align_center
        ws.cell(r_idx, 6, it['descripcion']).alignment = align_left
        ws.cell(r_idx, 7, str(it['envase'] or '')).alignment = align_center

        cu = ws.cell(r_idx, 8, it['unidades'])
        cu.number_format = fmt_units
        cu.alignment = align_right

        cpm = ws.cell(r_idx, 9, f"=J{r_idx}/H{r_idx}")
        cpm.number_format = fmt_price
        cpm.alignment = align_right

        ca = ws.cell(r_idx, 10, it['eur_adj'])
        ca.number_format = fmt_currency
        ca.alignment = align_right
        ca.font = font_bold

        cp = ws.cell(r_idx, 11, f"=J{r_idx}/$J${tot_row}")
        cp.number_format = fmt_pct
        cp.alignment = align_right

        fill_row = ZEBRA_FILL if i % 2 == 1 else WHITE_FILL
        for col_i in range(1, 12):
            cell = ws.cell(r_idx, col_i)
            cell.fill = fill_row
            cell.border = cell_border
            if col_i != 10:
                cell.font = font_regular

    # Fila de Totales
    ws.cell(tot_row, 1, "").border = double_bottom
    ws.cell(tot_row, 2, "TOTAL").font = font_bold
    ws.cell(tot_row, 3, "-").alignment = align_center
    ws.cell(tot_row, 4, f"TOTAL OCTUBRE 2026 ({len(items)} líneas activas)").font = font_bold
    ws.cell(tot_row, 5, "-").alignment = align_center
    ws.cell(tot_row, 6, "Consolidado Presupuesto P&L").font = font_regular
    ws.cell(tot_row, 7, "-").alignment = align_center

    tu = ws.cell(tot_row, 8, f"=SUM(H{start_r}:H{tot_row-1})")
    tu.number_format = fmt_units
    tu.font = font_bold
    tu.alignment = align_right

    tpm = ws.cell(tot_row, 9, f"=J{tot_row}/H{tot_row}")
    tpm.number_format = fmt_price
    tpm.font = font_bold
    tpm.alignment = align_right

    ta = ws.cell(tot_row, 10, f"=SUM(J{start_r}:J{tot_row-1})")
    ta.number_format = fmt_currency
    ta.font = Font(name="Segoe UI", size=10.5, bold=True, color=PRIMARY_DARK)
    ta.alignment = align_right

    tp = ws.cell(tot_row, 11, f"=SUM(K{start_r}:K{tot_row-1})")
    tp.number_format = fmt_pct
    tp.font = font_bold
    tp.alignment = align_right

    for col_i in range(1, 12):
        cell = ws.cell(tot_row, col_i)
        cell.fill = TOTAL_FILL
        cell.border = double_bottom

    # Filtros automáticos de Excel en fila 4
    ws.auto_filter.ref = f"A4:K{tot_row}"

    # Ajuste automático de anchos de columna
    col_widths = {
        "A": 6,   # #
        "B": 15,  # Comercial
        "C": 14,  # País
        "D": 38,  # Cliente
        "E": 14,  # Código SKU
        "F": 35,  # Descripción Producto
        "G": 10,  # Envase
        "H": 16,  # Unidades
        "I": 15,  # Precio Medio
        "J": 18,  # Importe
        "K": 12   # % s/ Total
    }
    for col_letter, width in col_widths.items():
        ws.column_dimensions[col_letter].width = width

    # Ajustar altura de filas
    ws.row_dimensions[1].height = 24
    ws.row_dimensions[2].height = 18
    ws.row_dimensions[4].height = 26
    ws.row_dimensions[tot_row].height = 22

    wb_out.save(out_file)
    print(f"[OK] Archivo generado exitosamente en una unica hoja: {out_file}")

if __name__ == '__main__':
    build_single_sheet_report()
