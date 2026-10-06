"""
Generador Profesional de Orden de Fabricación y Planificación MRP en Excel (XLSX).
Produce un archivo con diseño ejecutivo, tarjetas KPI, jerarquía visual de columnas,
formato de números con separadores de miles, semáforos condicionales y pestañas desglosadas.
"""

import json
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

def generar_excel_orden_fabricacion(
    json_path="web_dashboard/production_planning_data.json",
    output_path="Orden_Fabricacion_MRP_Codiagro.xlsx",
    scope="ALL",
    horizon_preset="N2",
    batch_size=0,
    format_filter="ALL"
):
    if not os.path.exists(json_path):
        json_path = "production_planning_data.json"
    
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    meta = data.get("metadata", {})
    presets = meta.get("presets", {})
    selected_months = presets.get(horizon_preset, presets.get("N2", ["2026-10", "2026-11", "2026-12"]))
    
    months_dict = {m["key"]: m.get("label_es", m["key"]) for m in meta.get("months", [])}
    horizon_label = " + ".join([months_dict.get(m, m) for m in selected_months])

    wb = openpyxl.Workbook()
    # Eliminar hoja por defecto
    ws_main = wb.active
    ws_main.title = "Orden de Fabricación"
    ws_main.views.sheetView[0].showGridLines = True

    # ---------------- PALETA DE COLORES CORPORATIVOS ----------------
    COLOR_HEADER_BG = "0F172A"       # Azul marino muy oscuro (Slate 900)
    COLOR_TITLE_TEXT = "FFFFFF"
    
    COLOR_GRP_ID = "1E3A8A"          # Azul cobalto oscuro (Identificación)
    COLOR_GRP_DEMAND = "334155"      # Pizarra oscuro (Stock & Demanda)
    COLOR_GRP_PLAN = "065F46"        # Verde bosque oscuro (Planificación MRP)
    
    COLOR_KPI_BG = "F8FAFC"          # Gris muy claro para KPIs
    COLOR_KPI_BORDER = "CBD5E1"
    
    COLOR_ROW_ALT = "F8FAFC"         # Fila par alterna suave
    COLOR_ROW_WHITE = "FFFFFF"
    
    COLOR_CRITICAL_BG = "FEE2E2"     # Rojo suave
    COLOR_CRITICAL_FG = "991B1B"
    COLOR_WARNING_BG = "FEF3C7"      # Ámbar suave
    COLOR_WARNING_FG = "92400E"
    COLOR_COVERED_BG = "D1FAE5"      # Verde suave
    COLOR_COVERED_FG = "065F46"

    # Fuentes
    font_main_title = Font(name="Calibri", size=15, bold=True, color=COLOR_TITLE_TEXT)
    font_sub_title = Font(name="Calibri", size=9, italic=True, color="94A3B8")
    font_kpi_label = Font(name="Calibri", size=8, bold=True, color="64748B")
    font_kpi_val = Font(name="Calibri", size=13, bold=True, color="0F172A")
    font_kpi_val_red = Font(name="Calibri", size=13, bold=True, color="DC2626")
    font_kpi_val_amber = Font(name="Calibri", size=13, bold=True, color="D97706")
    font_kpi_val_green = Font(name="Calibri", size=13, bold=True, color="059669")
    
    font_grp_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_col_header = Font(name="Calibri", size=9.5, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=9.5, color="1E293B")
    font_data_bold = Font(name="Calibri", size=9.5, bold=True, color="0F172A")
    font_total = Font(name="Calibri", size=10, bold=True, color="0F172A")

    # Bordes
    border_thin = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='thin', color='E2E8F0')
    )
    border_kpi = Border(
        left=Side(style='thin', color=COLOR_KPI_BORDER),
        right=Side(style='thin', color=COLOR_KPI_BORDER),
        top=Side(style='thin', color=COLOR_KPI_BORDER),
        bottom=Side(style='thin', color=COLOR_KPI_BORDER)
    )
    border_total = Border(
        top=Side(style='thin', color='0F172A'),
        bottom=Side(style='double', color='0F172A')
    )

    # 1. BANNER PRINCIPAL (Filas 1 y 2)
    ws_main.merge_cells("A1:P1")
    cell_title = ws_main["A1"]
    cell_title.value = "  CODIAGRO S.L.  |  ORDEN DE FABRICACIÓN Y PLANIFICACIÓN DE PRODUCCIÓN (MRP)"
    cell_title.font = font_main_title
    cell_title.fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
    cell_title.alignment = Alignment(vertical="center", horizontal="left")
    ws_main.row_dimensions[1].height = 28

    ws_main.merge_cells("A2:P2")
    cell_sub = ws_main["A2"]
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
    scope_str = "Global (Nacional + Exportación)" if scope == "ALL" else scope
    moq_str = f"Lote Mínimo {batch_size} u" if batch_size > 0 else "Exacto (Sin redondeo a lote)"
    cell_sub.value = f"  Horizonte: {horizon_label} ({len(selected_months)} meses)  |  Ámbito: {scope_str}  |  Criterio Lotes: {moq_str}  |  Fecha Emisión: {now_str}"
    cell_sub.font = font_sub_title
    cell_sub.fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
    cell_sub.alignment = Alignment(vertical="center", horizontal="left")
    ws_main.row_dimensions[2].height = 18

    ws_main.row_dimensions[3].height = 8

    # 2. PROCESAR FILAS
    n2_months = ['2026-10', '2026-11', '2026-12']
    n5_months = ['2026-10', '2026-11', '2026-12', '2027-01', '2027-02', '2027-03']
    
    rows_data = []
    tot_unidades_fabricar = 0
    tot_lotes_fabricar = 0
    count_critical = 0
    count_warning = 0
    count_covered = 0
    tot_pedidos_cartera = 0
    tot_stock_envasado = 0

    for item in data.get("base_skus", []):
        # Filtro ámbito
        if scope == "NACIONAL" and not item.get("has_nacional"):
            continue
        if scope == "EXPORTACION" and not item.get("has_export"):
            continue

        # Demanda horizonte
        forecast_period = sum(item.get("monthly_forecast", {}).get(m, 0) for m in selected_months)
        stock = float(item.get("stock_actual", 0) or 0)
        pending = float(item.get("pedidos_pendientes", 0) or 0)
        granel_stock = float(item.get("granel_stock", 0) or 0)
        
        total_demand = pending + forecast_period
        net_balance = stock - total_demand
        falta_envasar = max(0.0, -net_balance)
        
        falta_fabricar_granel = 0.0
        if falta_envasar > 0:
            if granel_stock >= falta_envasar:
                falta_fabricar_granel = 0.0
            else:
                falta_fabricar_granel = falta_envasar - granel_stock

        if falta_envasar > 0:
            if falta_fabricar_granel == 0.0:
                status = "WARNING"
                action_label = "PROGRAMAR LOTE"
                count_warning += 1
            else:
                status = "CRITICAL"
                action_label = "FABRICACIÓN URGENTE"
                count_critical += 1
        else:
            status = "COVERED"
            action_label = "CUBIERTO"
            count_covered += 1

        coverage_pct = 100.0
        if total_demand > 0:
            coverage_pct = ((stock + min(granel_stock, falta_envasar)) / total_demand) * 100.0
        elif stock == 0 and granel_stock == 0:
            coverage_pct = 0.0

        falta_para_lote = falta_fabricar_granel if falta_fabricar_granel > 0 else falta_envasar
        num_batches = 0
        sug_fabricar = falta_para_lote
        if batch_size > 0 and falta_para_lote > 0:
            import math
            num_batches = math.ceil(falta_para_lote / batch_size)
            sug_fabricar = num_batches * batch_size

        # Mes agotamiento
        run_stock = stock - pending
        depleted_month = "Cubierto >15 meses"
        if run_stock < 0:
            depleted_month = "Inmediato (Cartera)"
        else:
            for m_meta in meta.get("months", []):
                m_k = m_meta["key"]
                dem_m = float(item.get("monthly_forecast", {}).get(m_k, 0) or 0)
                run_stock -= dem_m
                if run_stock < 0:
                    depleted_month = m_meta.get("label_es", m_k)
                    break

        # Formato envase
        envase = str(item.get("envase", ""))
        fmt_label = "Garrafas 20L" if "20" in envase else ("Garrafas 5L" if "5" in envase else ("Botellas 1L" if "1" in envase else ("GRG 1000L" if "1000" in envase else "Otros")))

        # Mercado
        if item.get("has_nacional") and item.get("has_export"):
            mercado = "Nacional+Export"
        elif item.get("has_nacional"):
            mercado = "Nacional"
        else:
            mercado = "Exportación"

        tot_unidades_fabricar += sug_fabricar
        tot_lotes_fabricar += num_batches
        tot_pedidos_cartera += pending
        tot_stock_envasado += stock

        rows_data.append({
            "sku": item.get("base_sku", ""),
            "desc": item.get("desc", ""),
            "envase": envase,
            "formato": fmt_label,
            "mercado": mercado,
            "stock": stock,
            "pending": pending,
            "forecast": forecast_period,
            "demand": total_demand,
            "balance": net_balance,
            "falta_neta": falta_para_lote,
            "lotes": num_batches,
            "sug_fabricar": sug_fabricar,
            "cobertura": coverage_pct / 100.0, # Para formato 0.0%
            "agotamiento": depleted_month,
            "status": status,
            "action": action_label
        })

    # Ordenar: primero los críticos, luego warning, luego cubiertos; secundario por sugerencia de fabricar descendente
    priority_order = {"CRITICAL": 0, "WARNING": 1, "COVERED": 2}
    rows_data.sort(key=lambda r: (priority_order.get(r["status"], 9), -r["sug_fabricar"]))

    # 3. TARJETAS KPI (Filas 4 y 5)
    kpis = [
        ("A4", "C4", "A5", "C5", "TOTAL A FABRICAR", f"{tot_unidades_fabricar:,.0f} u".replace(",", "."), font_kpi_val_red),
        ("D4", "F4", "D5", "F5", "LOTES SUGERIDOS", f"{tot_lotes_fabricar:,.0f} lotes".replace(",", "."), font_kpi_val_amber),
        ("G4", "I4", "G5", "I5", "ARTÍCULOS URGENTES", f"{count_critical} SKUs (Déficit Granel)", font_kpi_val_red),
        ("J4", "L4", "J5", "L5", "SOLO ENVASAR", f"{count_warning} SKUs (Granel OK)", font_kpi_val_amber),
        ("M4", "N4", "M5", "N5", "PEDIDOS CARTERA", f"{tot_pedidos_cartera:,.0f} u".replace(",", "."), font_kpi_val),
        ("O4", "P4", "O5", "P5", "STOCK ENVASADO", f"{tot_stock_envasado:,.0f} u".replace(",", "."), font_kpi_val_green)
    ]

    fill_kpi = PatternFill(start_color=COLOR_KPI_BG, end_color=COLOR_KPI_BG, fill_type="solid")
    for top_l, top_r, bot_l, bot_r, label, val_text, val_font in kpis:
        ws_main.merge_cells(f"{top_l}:{top_r}")
        ws_main.merge_cells(f"{bot_l}:{bot_r}")
        
        c_lbl = ws_main[top_l]
        c_lbl.value = label
        c_lbl.font = font_kpi_label
        c_lbl.alignment = Alignment(horizontal="center", vertical="center")
        
        c_val = ws_main[bot_l]
        c_val.value = val_text
        c_val.font = val_font
        c_val.alignment = Alignment(horizontal="center", vertical="center")

        # Aplicar bordes y relleno a las celdas del bloque KPI
        start_col = openpyxl.utils.column_index_from_string(top_l[0])
        end_col = openpyxl.utils.column_index_from_string(top_r[0]) if len(top_r) == 2 else openpyxl.utils.column_index_from_string(top_r[:len(top_r)-1])
        row_top = int(top_l[1:])
        row_bot = int(bot_l[1:])
        
        for r_k in range(row_top, row_bot + 1):
            for c_k in range(start_col, end_col + 1):
                cell_k = ws_main.cell(r_k, c_k)
                cell_k.fill = fill_kpi
                cell_k.border = border_kpi

    ws_main.row_dimensions[4].height = 16
    ws_main.row_dimensions[5].height = 24
    ws_main.row_dimensions[6].height = 10

    # 4. ENCABEZADOS DE TABLA (Filas 7 y 8)
    # Fila 7: Grupos funcionales
    ws_main.merge_cells("A7:E7")
    c_g1 = ws_main["A7"]
    c_g1.value = "1. IDENTIFICACIÓN DEL ARTÍCULO"
    c_g1.font = font_grp_header
    c_g1.fill = PatternFill(start_color=COLOR_GRP_ID, end_color=COLOR_GRP_ID, fill_type="solid")
    c_g1.alignment = Alignment(horizontal="center", vertical="center")

    ws_main.merge_cells("F7:J7")
    c_g2 = ws_main["F7"]
    c_g2.value = "2. SITUACIÓN ACTUAL DE STOCK Y DEMANDA"
    c_g2.font = font_grp_header
    c_g2.fill = PatternFill(start_color=COLOR_GRP_DEMAND, end_color=COLOR_GRP_DEMAND, fill_type="solid")
    c_g2.alignment = Alignment(horizontal="center", vertical="center")

    ws_main.merge_cells("K7:P7")
    c_g3 = ws_main["K7"]
    c_g3.value = "3. PLAN DE FABRICACIÓN Y ACCIÓN MRP"
    c_g3.font = font_grp_header
    c_g3.fill = PatternFill(start_color=COLOR_GRP_PLAN, end_color=COLOR_GRP_PLAN, fill_type="solid")
    c_g3.alignment = Alignment(horizontal="center", vertical="center")

    ws_main.row_dimensions[7].height = 20

    # Fila 8: Columnas individuales claras y legibles
    headers = [
        ("A", "Código SKU", COLOR_GRP_ID, Alignment(horizontal="center")),
        ("B", "Producto / Descripción", COLOR_GRP_ID, Alignment(horizontal="left")),
        ("C", "Envase", COLOR_GRP_ID, Alignment(horizontal="center")),
        ("D", "Línea Envasado", COLOR_GRP_ID, Alignment(horizontal="center")),
        ("E", "Mercado", COLOR_GRP_ID, Alignment(horizontal="center")),
        
        ("F", "Stock Envasado (u)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("G", "Pedidos Cartera (u)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("H", "Previsión Período (u)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("I", "Demanda Total (u)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("J", "Balance Neto (u)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        
        ("K", "Déficit Neto (u)", COLOR_GRP_PLAN, Alignment(horizontal="right")),
        ("L", "Lotes", COLOR_GRP_PLAN, Alignment(horizontal="center")),
        ("M", "A Fabricar Sugerido (u)", COLOR_GRP_PLAN, Alignment(horizontal="right")),
        ("N", "Cobertura", COLOR_GRP_PLAN, Alignment(horizontal="center")),
        ("O", "Agotamiento Stock", COLOR_GRP_PLAN, Alignment(horizontal="center")),
        ("P", "Prioridad / Acción", COLOR_GRP_PLAN, Alignment(horizontal="center"))
    ]

    for col_letter, text, bg_color, align in headers:
        cell_h = ws_main[f"{col_letter}8"]
        cell_h.value = text
        cell_h.font = font_col_header
        cell_h.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        cell_h.alignment = align
        cell_h.border = border_thin

    ws_main.row_dimensions[8].height = 26

    # 5. FILAS DE DATOS
    start_row = 9
    current_row = start_row

    fill_alt = PatternFill(start_color=COLOR_ROW_ALT, end_color=COLOR_ROW_ALT, fill_type="solid")
    fill_white = PatternFill(start_color=COLOR_ROW_WHITE, end_color=COLOR_ROW_WHITE, fill_type="solid")

    fill_crit = PatternFill(start_color=COLOR_CRITICAL_BG, end_color=COLOR_CRITICAL_BG, fill_type="solid")
    font_crit = Font(name="Calibri", size=9.5, bold=True, color=COLOR_CRITICAL_FG)

    fill_warn = PatternFill(start_color=COLOR_WARNING_BG, end_color=COLOR_WARNING_BG, fill_type="solid")
    font_warn = Font(name="Calibri", size=9.5, bold=True, color=COLOR_WARNING_FG)

    fill_cov = PatternFill(start_color=COLOR_COVERED_BG, end_color=COLOR_COVERED_BG, fill_type="solid")
    font_cov = Font(name="Calibri", size=9.5, bold=True, color=COLOR_COVERED_FG)

    for r in rows_data:
        fill_row = fill_alt if (current_row % 2 == 0) else fill_white
        
        # A: SKU
        ws_main.cell(current_row, 1, r["sku"]).alignment = Alignment(horizontal="center")
        ws_main.cell(current_row, 1).font = font_data_bold
        
        # B: Desc
        ws_main.cell(current_row, 2, r["desc"]).alignment = Alignment(horizontal="left")
        ws_main.cell(current_row, 2).font = font_data
        
        # C: Envase
        ws_main.cell(current_row, 3, r["envase"]).alignment = Alignment(horizontal="center")
        ws_main.cell(current_row, 3).font = font_data
        
        # D: Formato
        ws_main.cell(current_row, 4, r["formato"]).alignment = Alignment(horizontal="center")
        ws_main.cell(current_row, 4).font = font_data
        
        # E: Mercado
        ws_main.cell(current_row, 5, r["mercado"]).alignment = Alignment(horizontal="center")
        ws_main.cell(current_row, 5).font = font_data
        
        # F: Stock
        c_f = ws_main.cell(current_row, 6, r["stock"])
        c_f.number_format = '#,##0'
        c_f.alignment = Alignment(horizontal="right")
        c_f.font = font_data
        
        # G: Pedidos
        c_g = ws_main.cell(current_row, 7, r["pending"])
        c_g.number_format = '#,##0'
        c_g.alignment = Alignment(horizontal="right")
        c_g.font = font_data_bold if r["pending"] > 0 else font_data
        
        # H: Prevision
        c_h = ws_main.cell(current_row, 8, r["forecast"])
        c_h.number_format = '#,##0'
        c_h.alignment = Alignment(horizontal="right")
        c_h.font = font_data
        
        # I: Demanda Total (=G+H)
        c_i = ws_main.cell(current_row, 9, f"=G{current_row}+H{current_row}")
        c_i.number_format = '#,##0'
        c_i.alignment = Alignment(horizontal="right")
        c_i.font = font_data_bold
        
        # J: Balance Neto (=F-I)
        c_j = ws_main.cell(current_row, 10, f"=F{current_row}-I{current_row}")
        c_j.number_format = '#,##0;[Red]-#,##0;0'
        c_j.alignment = Alignment(horizontal="right")
        c_j.font = font_data_bold
        
        # K: Falta Neta
        c_k = ws_main.cell(current_row, 11, r["falta_neta"])
        c_k.number_format = '#,##0'
        c_k.alignment = Alignment(horizontal="right")
        c_k.font = font_data_bold if r["falta_neta"] > 0 else font_data
        
        # L: Lotes
        c_l = ws_main.cell(current_row, 12, r["lotes"])
        c_l.number_format = '#,##0'
        c_l.alignment = Alignment(horizontal="center")
        c_l.font = font_data
        
        # M: Sugerencia Producción
        c_m = ws_main.cell(current_row, 13, r["sug_fabricar"])
        c_m.number_format = '#,##0'
        c_m.alignment = Alignment(horizontal="right")
        c_m.font = Font(name="Calibri", size=10, bold=True, color="991B1B" if r["sug_fabricar"] > 0 else "0F172A")
        
        # N: Cobertura
        c_n = ws_main.cell(current_row, 14, r["cobertura"])
        c_n.number_format = '0.0%'
        c_n.alignment = Alignment(horizontal="center")
        c_n.font = font_data
        
        # O: Agotamiento
        c_o = ws_main.cell(current_row, 15, r["agotamiento"])
        c_o.alignment = Alignment(horizontal="center")
        c_o.font = font_data_bold if "Inmediato" in r["agotamiento"] else font_data
        
        # P: Prioridad / Acción con Badge Coloreado
        c_p = ws_main.cell(current_row, 16, r["action"])
        c_p.alignment = Alignment(horizontal="center")
        if r["status"] == "CRITICAL":
            c_p.fill = fill_crit
            c_p.font = font_crit
        elif r["status"] == "WARNING":
            c_p.fill = fill_warn
            c_p.font = font_warn
        else:
            c_p.fill = fill_cov
            c_p.font = font_cov

        # Aplicar bordes y relleno base (columnas A a O)
        for col_idx in range(1, 16):
            cell_x = ws_main.cell(current_row, col_idx)
            cell_x.fill = fill_row
            cell_x.border = border_thin
        ws_main.cell(current_row, 16).border = border_thin

        ws_main.row_dimensions[current_row].height = 20
        current_row += 1

    # 6. FILA DE TOTALES
    tot_row = current_row
    ws_main.merge_cells(f"A{tot_row}:E{tot_row}")
    c_tot_lbl = ws_main[f"A{tot_row}"]
    c_tot_lbl.value = f"TOTAL GENERAL ({len(rows_data)} SKUs)"
    c_tot_lbl.font = font_total
    c_tot_lbl.alignment = Alignment(horizontal="right", vertical="center")

    sum_cols = [
        (6, "F"), (7, "G"), (8, "H"), (9, "I"), (10, "J"),
        (11, "K"), (12, "L"), (13, "M")
    ]
    for c_idx, letter in sum_cols:
        cell_s = ws_main.cell(tot_row, c_idx)
        cell_s.value = f"=SUM({letter}{start_row}:{letter}{tot_row - 1})"
        cell_s.number_format = '#,##0'
        cell_s.font = font_total
        cell_s.alignment = Alignment(horizontal="right", vertical="center")

    # Cobertura promedio ponderada o en blanco
    ws_main.cell(tot_row, 14, f"=IF(I{tot_row}>0, F{tot_row}/I{tot_row}, 1)").number_format = '0.0%'
    ws_main.cell(tot_row, 14).font = font_total
    ws_main.cell(tot_row, 14).alignment = Alignment(horizontal="center", vertical="center")

    ws_main.cell(tot_row, 15, "").alignment = Alignment(horizontal="center")
    ws_main.cell(tot_row, 16, "").alignment = Alignment(horizontal="center")

    fill_total = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    for col_idx in range(1, 17):
        cell_t = ws_main.cell(tot_row, col_idx)
        cell_t.fill = fill_total
        cell_t.border = border_total

    ws_main.row_dimensions[tot_row].height = 26

    # 7. ANCHOS DE COLUMNAS HOLGADOS (Evitar cortes de texto y ########)
    col_widths = {
        'A': 15,   # Código SKU
        'B': 38,   # Producto / Descripción
        'C': 12,   # Envase
        'D': 18,   # Línea Envasado
        'E': 16,   # Mercado
        'F': 16,   # Stock Envasado
        'G': 16,   # Pedidos Cartera
        'H': 17,   # Previsión Período
        'I': 17,   # Demanda Total
        'J': 16,   # Balance Neto
        'K': 16,   # Déficit Neto
        'L': 11,   # Lotes
        'M': 20,   # A Fabricar Sugerido
        'N': 13,   # Cobertura
        'O': 20,   # Agotamiento Stock
        'P': 22    # Prioridad / Acción
    }
    for col_l, width in col_widths.items():
        ws_main.column_dimensions[col_l].width = width

    # Inmovilizar paneles en fila 9 (para que los encabezados permanezcan fijos al hacer scroll)
    ws_main.freeze_panes = "A9"

    # Activar autofiltro en la fila de encabezados 8
    ws_main.auto_filter.ref = f"A8:P{tot_row - 1}"

    # ---------------- HOJA 2: RESUMEN POR LÍNEA DE ENVASADO ----------------
    ws_env = wb.create_sheet(title="Resumen por Envase")
    ws_env.views.sheetView[0].showGridLines = True
    
    ws_env.merge_cells("A1:G1")
    c_env_title = ws_env["A1"]
    c_env_title.value = "  CODIAGRO S.L.  |  RESUMEN DE NECESIDADES POR LÍNEA DE ENVASADO"
    c_env_title.font = font_main_title
    c_env_title.fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
    c_env_title.alignment = Alignment(vertical="center", horizontal="left")
    ws_env.row_dimensions[1].height = 28

    env_headers = [
        ("A", "Línea / Tipo Envase", 24, Alignment(horizontal="left")),
        ("B", "Nº SKUs", 12, Alignment(horizontal="center")),
        ("C", "Stock Actual (u)", 18, Alignment(horizontal="right")),
        ("D", "Pedidos Cartera (u)", 18, Alignment(horizontal="right")),
        ("E", "Previsión (u)", 18, Alignment(horizontal="right")),
        ("F", "Demanda Total (u)", 18, Alignment(horizontal="right")),
        ("G", "A Fabricar Sugerido (u)", 22, Alignment(horizontal="right"))
    ]

    for col_l, text, width, align in env_headers:
        cell_eh = ws_env[f"{col_l}3"]
        cell_eh.value = text
        cell_eh.font = font_col_header
        cell_eh.fill = PatternFill(start_color=COLOR_GRP_ID, end_color=COLOR_GRP_ID, fill_type="solid")
        cell_eh.alignment = align
        cell_eh.border = border_thin
        ws_env.column_dimensions[col_l].width = width
    ws_env.row_dimensions[3].height = 24

    # Agrupar por línea de envasado
    from collections import defaultdict
    format_agg = defaultdict(lambda: {"count": 0, "stock": 0, "pending": 0, "forecast": 0, "demand": 0, "sug": 0})
    for r in rows_data:
        fmt = r["formato"]
        format_agg[fmt]["count"] += 1
        format_agg[fmt]["stock"] += r["stock"]
        format_agg[fmt]["pending"] += r["pending"]
        format_agg[fmt]["forecast"] += r["forecast"]
        format_agg[fmt]["demand"] += r["demand"]
        format_agg[fmt]["sug"] += r["sug_fabricar"]

    cur_env_row = 4
    for fmt_name in sorted(format_agg.keys()):
        agg = format_agg[fmt_name]
        ws_env.cell(cur_env_row, 1, fmt_name).font = font_data_bold
        ws_env.cell(cur_env_row, 1).alignment = Alignment(horizontal="left")
        
        ws_env.cell(cur_env_row, 2, agg["count"]).font = font_data
        ws_env.cell(cur_env_row, 2).alignment = Alignment(horizontal="center")
        
        for c_i, k_val in [(3, "stock"), (4, "pending"), (5, "forecast"), (6, "demand"), (7, "sug")]:
            c_val = ws_env.cell(cur_env_row, c_i, agg[k_val])
            c_val.number_format = '#,##0'
            c_val.font = font_data_bold if c_i == 7 else font_data
            c_val.alignment = Alignment(horizontal="right")

        for col_i in range(1, 8):
            ws_env.cell(cur_env_row, col_i).border = border_thin
            ws_env.cell(cur_env_row, col_i).fill = fill_alt if (cur_env_row % 2 == 0) else fill_white

        ws_env.row_dimensions[cur_env_row].height = 20
        cur_env_row += 1

    # Total de la hoja 2
    ws_env.cell(cur_env_row, 1, "TOTAL PLANTA").font = font_total
    ws_env.cell(cur_env_row, 1).alignment = Alignment(horizontal="left")
    ws_env.cell(cur_env_row, 2, f"=SUM(B4:B{cur_env_row - 1})").font = font_total
    ws_env.cell(cur_env_row, 2).alignment = Alignment(horizontal="center")
    
    for c_i, letter in [(3, "C"), (4, "D"), (5, "E"), (6, "F"), (7, "G")]:
        c_val = ws_env.cell(cur_env_row, c_i, f"=SUM({letter}4:{letter}{cur_env_row - 1})")
        c_val.number_format = '#,##0'
        c_val.font = font_total
        c_val.alignment = Alignment(horizontal="right")

    for col_i in range(1, 8):
        ws_env.cell(cur_env_row, col_i).fill = fill_total
        ws_env.cell(cur_env_row, col_i).border = border_total
    ws_env.row_dimensions[cur_env_row].height = 24

    # Guardar
    wb.save(output_path)
    print(f"¡Orden de Fabricación MRP generada con éxito en: {output_path}!")
    return output_path

if __name__ == "__main__":
    generar_excel_orden_fabricacion()
