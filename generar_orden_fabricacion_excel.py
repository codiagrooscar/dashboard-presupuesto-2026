"""
Generador Profesional de Orden de Fabricación y Planificación MRP en Excel (XLSX).
Produce un archivo con diseño corporativo Codiagro:
1. Hoja "Orden de Fabricación": Desglose individual de cada artículo y país/mercado con su 'Nuevo SKU'
   específico (sin concatenar), litros y unidades por país, con columnas mensuales Oct-26, Nov-26, Dic-26.
2. Hoja "Plan Graneles por Mes": Resumen maestro de productos químicos a fabricar en reactores (L/Kg),
   con los Graneles en filas y los Meses en columnas.
3. Hoja "Resumen por Envase": Clasificación estricta de líneas de envasado (Depósito 1000 L, Depósito 200 L,
   Garrafas 20L, Garrafas 5L, Botellas 1L, etc.).
"""

import json
import os
import re
import math
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

def parse_envase_litros(env_val, desc=""):
    try:
        val = float(str(env_val).replace(",", "."))
        if val > 0:
            return val
    except (ValueError, TypeError):
        pass
    m = re.search(r'(\d+[\.,]?\d*)\s*(?:L|KG|LITROS|KILOS)', str(desc), re.I)
    if m:
        try:
            return float(m.group(1).replace(",", "."))
        except (ValueError, TypeError):
            pass
    return 1.0

def get_packaging_format(env_val, desc=""):
    l_val = parse_envase_litros(env_val, desc)
    if l_val >= 900:
        return "Depósito 1000 L"
    elif l_val >= 150:
        return "Depósito 200 L"
    elif l_val >= 15:
        return "Garrafas 20L"
    elif l_val >= 4:
        return "Garrafas 5L"
    elif l_val >= 0.9:
        return "Botellas 1L"
    elif l_val > 0:
        return f"Botellas {l_val}L"
    return "Otros Formatos"

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
    
    # ---------------- PALETA CORPORATIVA CODIAGRO ----------------
    COLOR_HEADER_BG = "0F172A"       # Slate 900
    COLOR_TITLE_TEXT = "FFFFFF"
    
    COLOR_GRP_ID = "1E3A8A"          # Azul cobalto oscuro (Identificación)
    COLOR_GRP_DEMAND = "334155"      # Pizarra oscuro (Demanda & Fechas)
    COLOR_GRP_PLAN = "065F46"        # Verde bosque oscuro (Planificación MRP)
    
    COLOR_KPI_BG = "F8FAFC"
    COLOR_KPI_BORDER = "CBD5E1"
    
    COLOR_ROW_ALT = "F8FAFC"
    COLOR_ROW_WHITE = "FFFFFF"
    
    COLOR_CRITICAL_BG = "FEE2E2"     # Rojo suave
    COLOR_CRITICAL_FG = "991B1B"
    COLOR_WARNING_BG = "FEF3C7"      # Ámbar suave
    COLOR_WARNING_FG = "92400E"
    COLOR_COVERED_BG = "D1FAE5"      # Verde suave
    COLOR_COVERED_FG = "065F46"

    # Fuentes
    font_main_title = Font(name="Calibri", size=13, bold=True, color=COLOR_TITLE_TEXT)
    font_sub_title = Font(name="Calibri", size=9, italic=True, color="94A3B8")
    font_kpi_label = Font(name="Calibri", size=8, bold=True, color="64748B")
    font_kpi_val = Font(name="Calibri", size=12, bold=True, color="0F172A")
    font_kpi_val_red = Font(name="Calibri", size=12, bold=True, color="DC2626")
    font_kpi_val_amber = Font(name="Calibri", size=12, bold=True, color="D97706")
    font_kpi_val_green = Font(name="Calibri", size=12, bold=True, color="059669")
    
    font_grp_header = Font(name="Calibri", size=9.5, bold=True, color="FFFFFF")
    font_col_header = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=9, color="1E293B")
    font_data_bold = Font(name="Calibri", size=9, bold=True, color="0F172A")
    font_data_purple = Font(name="Calibri", size=9, bold=True, color="6B21A8")
    font_total = Font(name="Calibri", size=9.5, bold=True, color="0F172A")

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

    fill_alt = PatternFill(start_color=COLOR_ROW_ALT, end_color=COLOR_ROW_ALT, fill_type="solid")
    fill_white = PatternFill(start_color=COLOR_ROW_WHITE, end_color=COLOR_ROW_WHITE, fill_type="solid")
    fill_crit = PatternFill(start_color=COLOR_CRITICAL_BG, end_color=COLOR_CRITICAL_BG, fill_type="solid")
    fill_warn = PatternFill(start_color=COLOR_WARNING_BG, end_color=COLOR_WARNING_BG, fill_type="solid")
    fill_cov = PatternFill(start_color=COLOR_COVERED_BG, end_color=COLOR_COVERED_BG, fill_type="solid")

    font_crit = Font(name="Calibri", size=8.5, bold=True, color=COLOR_CRITICAL_FG)
    font_warn = Font(name="Calibri", size=8.5, bold=True, color=COLOR_WARNING_FG)
    font_cov = Font(name="Calibri", size=8.5, bold=True, color=COLOR_COVERED_FG)

    now_str = datetime.now().strftime("%d/%m/%Y %H:%M")
    scope_str = "Global (Nacional + Exportación)" if scope == "ALL" else scope
    moq_str = f"Lote Mínimo {batch_size} u" if batch_size > 0 else "Exacto"

    # =========================================================================
    # HOJA 1: ORDEN DE FABRICACIÓN (Desglosada por SKU / País)
    # =========================================================================
    ws_main = wb.active
    ws_main.title = "Orden de Fabricación"
    ws_main.views.sheetView[0].showGridLines = True

    # Banner Superior (A1:P2)
    ws_main.merge_cells("A1:P1")
    c_title = ws_main["A1"]
    c_title.value = "  CODIAGRO S.L.  |  ORDEN DE FABRICACIÓN Y ENVASADO MRP (DESGLOSE POR PAÍS Y FORMATO)"
    c_title.font = font_main_title
    c_title.fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
    c_title.alignment = Alignment(vertical="center", horizontal="left")
    ws_main.row_dimensions[1].height = 28

    ws_main.merge_cells("A2:P2")
    c_sub = ws_main["A2"]
    c_sub.value = f"  Horizonte: {horizon_label} ({len(selected_months)} meses)  |  Ámbito: {scope_str}  |  Lotes: {moq_str}  |  Fecha Emisión: {now_str}"
    c_sub.font = font_sub_title
    c_sub.fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
    c_sub.alignment = Alignment(vertical="center", horizontal="left")
    ws_main.row_dimensions[2].height = 18

    ws_main.row_dimensions[3].height = 8

    # Procesamiento de Filas Desglosadas por Destino / País
    rows_data = []
    tot_unidades_fabricar = 0
    tot_litros_fabricar = 0
    count_critical = 0
    count_warning = 0
    count_covered = 0
    tot_pedidos_cartera = 0
    tot_stock_envasado = 0

    for b in data.get("base_skus", []):
        if scope == "NACIONAL" and not b.get("has_nacional"):
            continue
        if scope == "EXPORTACION" and not b.get("has_export"):
            continue

        base_sku = b.get("base_sku", "")
        desc = b.get("desc", "")
        env_str = str(b.get("envase", ""))
        env_l = parse_envase_litros(env_str, desc)
        fmt_label = get_packaging_format(env_str, desc)
        granel_stock = float(b.get("granel_stock", 0) or 0)
        stock_total_b = float(b.get("stock_actual", 0) or 0)

        # Destinos del SKU base
        destinos = b.get("destinos", [])
        if not destinos:
            destinos = [{
                "sku": base_sku,
                "desc": desc,
                "pais": "ESPAÑA" if b.get("has_nacional") else "EXPORTACIÓN",
                "ambito": "Nacional" if b.get("has_nacional") else "Exportación",
                "pendientes_uds": b.get("pedidos_pendientes", 0),
                "monthly_forecast": b.get("monthly_forecast", {})
            }]

        # Filtrar solo destinos con demanda o cartera en el ámbito solicitado
        active_destinos = []
        for d in destinos:
            d_ambito = d.get("ambito", "Nacional")
            if scope == "NACIONAL" and d_ambito != "Nacional":
                continue
            if scope == "EXPORTACION" and d_ambito != "Exportación":
                continue
            
            p_uds = float(d.get("pendientes_uds", 0) or 0)
            f_uds = sum(float(d.get("monthly_forecast", {}).get(m, 0) or 0) for m in selected_months)
            if p_uds > 0 or f_uds > 0:
                active_destinos.append(d)

        # Si ningún destino tiene demanda pero hay stock, mostramos al menos 1 fila base
        if not active_destinos and stock_total_b > 0:
            active_destinos = [destinos[0]]

        # Asignación secuencial de stock disponible a los destinos
        rem_stock = stock_total_b
        tot_stock_envasado += stock_total_b

        for d in active_destinos:
            p_uds = float(d.get("pendientes_uds", 0) or 0)
            oct_u = float(d.get("monthly_forecast", {}).get("2026-10", 0) or 0)
            nov_u = float(d.get("monthly_forecast", {}).get("2026-11", 0) or 0)
            dic_u = float(d.get("monthly_forecast", {}).get("2026-12", 0) or 0)
            
            dem_period_u = sum(float(d.get("monthly_forecast", {}).get(m, 0) or 0) for m in selected_months)
            total_dem_u = p_uds + dem_period_u
            total_dem_l = total_dem_u * env_l

            allocated_stock = min(rem_stock, total_dem_u)
            rem_stock -= allocated_stock
            
            falta_envasar_u = max(0.0, total_dem_u - allocated_stock)
            falta_envasar_l = falta_envasar_u * env_l

            sug_fab_u = falta_envasar_u
            if batch_size > 0 and falta_envasar_u > 0:
                num_batches = math.ceil(falta_envasar_u / batch_size)
                sug_fab_u = num_batches * batch_size
            sug_fab_l = sug_fab_u * env_l

            # Estado
            if falta_envasar_u > 0:
                if granel_stock >= falta_envasar_l:
                    status = "WARNING"
                    action_label = "SOLO ENVASAR"
                    count_warning += 1
                else:
                    status = "CRITICAL"
                    action_label = "FABRICACIÓN URGENTE"
                    count_critical += 1
            else:
                status = "COVERED"
                action_label = "CUBIERTO"
                count_covered += 1

            cov_pct = 1.0
            if total_dem_u > 0:
                cov_pct = min(1.0, allocated_stock / total_dem_u)
            elif allocated_stock == 0:
                cov_pct = 0.0

            pais_nombre = d.get("pais", "ESPAÑA")
            nuevo_sku = d.get("sku", base_sku)

            tot_unidades_fabricar += sug_fab_u
            tot_litros_fabricar += sug_fab_l
            tot_pedidos_cartera += p_uds

            rows_data.append({
                "base_sku": base_sku,
                "nuevo_sku": nuevo_sku,
                "desc": desc,
                "envase": env_l,
                "formato": fmt_label,
                "pais": pais_nombre,
                "stock_asignado": allocated_stock,
                "pending": p_uds,
                "oct_u": oct_u,
                "nov_u": nov_u,
                "dic_u": dic_u,
                "demand_u": total_dem_u,
                "demand_l": total_dem_l,
                "sug_u": sug_fab_u,
                "sug_l": sug_fab_l,
                "cobertura": cov_pct,
                "status": status,
                "action": action_label
            })

    priority_order = {"CRITICAL": 0, "WARNING": 1, "COVERED": 2}
    rows_data.sort(key=lambda r: (priority_order.get(r["status"], 9), -r["sug_l"]))

    # Tarjetas KPI (Filas 4 y 5)
    kpis = [
        ("A4", "B4", "A5", "B5", "TOTAL A FABRICAR (L/Kg)", f"{tot_litros_fabricar:,.0f} L".replace(",", "."), font_kpi_val_red),
        ("C4", "D4", "C5", "D5", "TOTAL UNIDADES (u)", f"{tot_unidades_fabricar:,.0f} u".replace(",", "."), font_kpi_val_red),
        ("E4", "F4", "E5", "F5", "DÉFICIT URGENTE", f"{count_critical} Artículos", font_kpi_val_red),
        ("G4", "H4", "G5", "H5", "SOLO ENVASAR (Granel OK)", f"{count_warning} Artículos", font_kpi_val_amber),
        ("I4", "L4", "I5", "L5", "PEDIDOS CARTERA INMEDIATOS", f"{tot_pedidos_cartera:,.0f} u".replace(",", "."), font_kpi_val),
        ("M4", "P4", "M5", "P5", "STOCK ENVASADO ALMACÉN", f"{tot_stock_envasado:,.0f} u".replace(",", "."), font_kpi_val_green)
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

    # Encabezados de Tabla (Filas 7 y 8)
    ws_main.merge_cells("A7:F7")
    c_g1 = ws_main["A7"]
    c_g1.value = "1. IDENTIFICACIÓN Y MERCADO DESTINO"
    c_g1.font = font_grp_header
    c_g1.fill = PatternFill(start_color=COLOR_GRP_ID, end_color=COLOR_GRP_ID, fill_type="solid")
    c_g1.alignment = Alignment(horizontal="center", vertical="center")

    ws_main.merge_cells("G7:M7")
    c_g2 = ws_main["G7"]
    c_g2.value = "2. SITUACIÓN DE DEMANDA Y PREVISIÓN POR FECHAS"
    c_g2.font = font_grp_header
    c_g2.fill = PatternFill(start_color=COLOR_GRP_DEMAND, end_color=COLOR_GRP_DEMAND, fill_type="solid")
    c_g2.alignment = Alignment(horizontal="center", vertical="center")

    ws_main.merge_cells("N7:P7")
    c_g3 = ws_main["N7"]
    c_g3.value = "3. PLAN DE FABRICACIÓN SUGERIDO"
    c_g3.font = font_grp_header
    c_g3.fill = PatternFill(start_color=COLOR_GRP_PLAN, end_color=COLOR_GRP_PLAN, fill_type="solid")
    c_g3.alignment = Alignment(horizontal="center", vertical="center")

    ws_main.row_dimensions[7].height = 20

    main_headers = [
        ("A", "Código Base", COLOR_GRP_ID, Alignment(horizontal="center")),
        ("B", "Nuevo SKU", COLOR_GRP_ID, Alignment(horizontal="center")),
        ("C", "Producto / Descripción", COLOR_GRP_ID, Alignment(horizontal="left")),
        ("D", "Envase", COLOR_GRP_ID, Alignment(horizontal="center")),
        ("E", "Línea Envasado", COLOR_GRP_ID, Alignment(horizontal="center")),
        ("F", "País / Mercado", COLOR_GRP_ID, Alignment(horizontal="center")),
        
        ("G", "Stock Asignado (u)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("H", "Cartera Inmediata (u)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("I", "Oct-26 (u)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("J", "Nov-26 (u)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("K", "Dic-26 (u)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("L", "Demanda Total (u)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("M", "Demanda Total (L/Kg)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        
        ("N", "A Fabricar (u)", COLOR_GRP_PLAN, Alignment(horizontal="right")),
        ("O", "A Fabricar (L/Kg)", COLOR_GRP_PLAN, Alignment(horizontal="right")),
        ("P", "Prioridad / Acción", COLOR_GRP_PLAN, Alignment(horizontal="center"))
    ]

    for col_l, text, bg_color, align in main_headers:
        cell_h = ws_main[f"{col_l}8"]
        cell_h.value = text
        cell_h.font = font_col_header
        cell_h.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        cell_h.alignment = align
        cell_h.border = border_thin
    ws_main.row_dimensions[8].height = 24

    current_row = 9
    start_row = current_row

    for r in rows_data:
        fill_row = fill_alt if (current_row % 2 == 0) else fill_white
        
        ws_main.cell(current_row, 1, r["base_sku"]).font = font_data_bold
        ws_main.cell(current_row, 1).alignment = Alignment(horizontal="center")
        
        # Nuevo SKU en morado destacado
        ws_main.cell(current_row, 2, r["nuevo_sku"]).font = font_data_purple
        ws_main.cell(current_row, 2).alignment = Alignment(horizontal="center")
        
        ws_main.cell(current_row, 3, r["desc"]).font = font_data
        ws_main.cell(current_row, 3).alignment = Alignment(horizontal="left")
        
        ws_main.cell(current_row, 4, r["envase"]).font = font_data
        ws_main.cell(current_row, 4).number_format = '#,##0.##'
        ws_main.cell(current_row, 4).alignment = Alignment(horizontal="center")
        
        ws_main.cell(current_row, 5, r["formato"]).font = font_data
        ws_main.cell(current_row, 5).alignment = Alignment(horizontal="center")
        
        ws_main.cell(current_row, 6, r["pais"]).font = font_data_bold
        ws_main.cell(current_row, 6).alignment = Alignment(horizontal="center")

        num_cells = [
            (7, r["stock_asignado"]),
            (8, r["pending"]),
            (9, r["oct_u"]),
            (10, r["nov_u"]),
            (11, r["dic_u"]),
            (12, r["demand_u"]),
            (13, r["demand_l"]),
            (14, r["sug_u"]),
            (15, r["sug_l"])
        ]

        for col_i, val_n in num_cells:
            c_num = ws_main.cell(current_row, col_i, val_n)
            c_num.number_format = '#,##0'
            c_num.alignment = Alignment(horizontal="right")
            if col_i in [12, 13, 14, 15]:
                c_num.font = font_data_bold
            else:
                c_num.font = font_data

        # Prioridad / Acción con semáforo
        c_act = ws_main.cell(current_row, 16, r["action"])
        c_act.alignment = Alignment(horizontal="center")
        if r["status"] == "CRITICAL":
            c_act.fill = fill_crit
            c_act.font = font_crit
        elif r["status"] == "WARNING":
            c_act.fill = fill_warn
            c_act.font = font_warn
        else:
            c_act.fill = fill_cov
            c_act.font = font_cov

        for col_idx in range(1, 16):
            cell_x = ws_main.cell(current_row, col_idx)
            cell_x.fill = fill_row
            cell_x.border = border_thin
        ws_main.cell(current_row, 16).border = border_thin

        ws_main.row_dimensions[current_row].height = 20
        current_row += 1

    # Fila de Totales Hoja 1
    tot_row = current_row
    ws_main.merge_cells(f"A{tot_row}:F{tot_row}")
    c_tot_lbl = ws_main[f"A{tot_row}"]
    c_tot_lbl.value = f"TOTAL GENERAL ({len(rows_data)} Artículos / Destinos)"
    c_tot_lbl.font = font_total
    c_tot_lbl.alignment = Alignment(horizontal="right", vertical="center")

    sum_cols = [
        (7, "G"), (8, "H"), (9, "I"), (10, "J"), (11, "K"),
        (12, "L"), (13, "M"), (14, "N"), (15, "O")
    ]
    for c_idx, letter in sum_cols:
        cell_s = ws_main.cell(tot_row, c_idx)
        cell_s.value = f"=SUM({letter}{start_row}:{letter}{tot_row - 1})"
        cell_s.number_format = '#,##0'
        cell_s.font = font_total
        cell_s.alignment = Alignment(horizontal="right", vertical="center")

    ws_main.cell(tot_row, 16, "").alignment = Alignment(horizontal="center")

    fill_total = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    for col_idx in range(1, 17):
        cell_t = ws_main.cell(tot_row, col_idx)
        cell_t.fill = fill_total
        cell_t.border = border_total

    ws_main.row_dimensions[tot_row].height = 26

    # Anchos de Columna Hoja 1
    col_widths_main = {
        'A': 14,   # Código Base
        'B': 18,   # Nuevo SKU
        'C': 36,   # Producto / Descripción
        'D': 10,   # Envase
        'E': 18,   # Línea Envasado
        'F': 18,   # País / Mercado
        'G': 17,   # Stock Asignado (u)
        'H': 17,   # Cartera Inmediata (u)
        'I': 14,   # Oct-26 (u)
        'J': 14,   # Nov-26 (u)
        'K': 14,   # Dic-26 (u)
        'L': 18,   # Demanda Total (u)
        'M': 20,   # Demanda Total (L/Kg)
        'N': 17,   # A Fabricar (u)
        'O': 20,   # A Fabricar (L/Kg)
        'P': 22    # Prioridad / Acción
    }
    for col_l, width in col_widths_main.items():
        ws_main.column_dimensions[col_l].width = width

    ws_main.freeze_panes = "A9"
    ws_main.auto_filter.ref = f"A8:P{tot_row - 1}"

    # =========================================================================
    # HOJA 2: PLAN GRANELES POR MES (Graneles en Filas, Meses en Columnas)
    # =========================================================================
    ws_granel = wb.create_sheet(title="Plan Graneles por Mes")
    ws_granel.views.sheetView[0].showGridLines = True

    # Banner Superior Hoja 2
    ws_granel.merge_cells("A1:M1")
    c_g_title = ws_granel["A1"]
    c_g_title.value = "  CODIAGRO S.L.  |  PLAN MAESTRO DE PRODUCCIÓN DE GRANELES (NECESIDAD EN LITROS / KG POR MES)"
    c_g_title.font = font_main_title
    c_g_title.fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
    c_g_title.alignment = Alignment(vertical="center", horizontal="left")
    ws_granel.row_dimensions[1].height = 28

    ws_granel.merge_cells("A2:M2")
    c_g_sub = ws_granel["A2"]
    c_g_sub.value = f"  Horizonte: Octubre 2026 - Diciembre 2026  |  Total Graneles Químicos a Fabricar en Planta  |  Fecha Emisión: {now_str}"
    c_g_sub.font = font_sub_title
    c_g_sub.fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
    c_g_sub.alignment = Alignment(vertical="center", horizontal="left")
    ws_granel.row_dimensions[2].height = 18

    # Encabezados de grupos en Hoja 2 (Fila 4)
    ws_granel.merge_cells("A4:B4")
    c_g_g1 = ws_granel["A4"]
    c_g_g1.value = "1. IDENTIFICACIÓN GRANEL"
    c_g_g1.font = font_grp_header
    c_g_g1.fill = PatternFill(start_color=COLOR_GRP_ID, end_color=COLOR_GRP_ID, fill_type="solid")
    c_g_g1.alignment = Alignment(horizontal="center", vertical="center")

    ws_granel.merge_cells("C4:E4")
    c_g_g2 = ws_granel["C4"]
    c_g_g2.value = "2. DISPONIBILIDAD INICIAL (L/KG)"
    c_g_g2.font = font_grp_header
    c_g_g2.fill = PatternFill(start_color=COLOR_GRP_DEMAND, end_color=COLOR_GRP_DEMAND, fill_type="solid")
    c_g_g2.alignment = Alignment(horizontal="center", vertical="center")

    ws_granel.merge_cells("F4:J4")
    c_g_g3 = ws_granel["F4"]
    c_g_g3.value = "3. DEMANDA Y PREVISIÓN MENSUAL DE GRANEL (L/KG)"
    c_g_g3.font = font_grp_header
    c_g_g3.fill = PatternFill(start_color=COLOR_GRP_DEMAND, end_color=COLOR_GRP_DEMAND, fill_type="solid")
    c_g_g3.alignment = Alignment(horizontal="center", vertical="center")

    ws_granel.merge_cells("K4:M4")
    c_g_g4 = ws_granel["K4"]
    c_g_g4.value = "4. NECESIDAD NETA FABRICACIÓN REACTORES"
    c_g_g4.font = font_grp_header
    c_g_g4.fill = PatternFill(start_color=COLOR_GRP_PLAN, end_color=COLOR_GRP_PLAN, fill_type="solid")
    c_g_g4.alignment = Alignment(horizontal="center", vertical="center")

    ws_granel.row_dimensions[4].height = 20

    granel_headers = [
        ("A", "Cód. Granel", COLOR_GRP_ID, Alignment(horizontal="center")),
        ("B", "Producto Químico / Granel", COLOR_GRP_ID, Alignment(horizontal="left")),
        
        ("C", "Stock Tanque (L/Kg)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("D", "Stock Envasado (L/Kg)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("E", "Stock Total (L/Kg)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        
        ("F", "Cartera Inmediata (L/Kg)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("G", "Oct-26 (L/Kg)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("H", "Nov-26 (L/Kg)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("I", "Dic-26 (L/Kg)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("J", "Demanda Total (L/Kg)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        
        ("K", "Balance Neto (L/Kg)", COLOR_GRP_PLAN, Alignment(horizontal="right")),
        ("L", "FALTA FABRICAR (L/Kg)", COLOR_GRP_PLAN, Alignment(horizontal="right")),
        ("M", "Prioridad Reactor", COLOR_GRP_PLAN, Alignment(horizontal="center"))
    ]

    for col_l, text, bg_color, align in granel_headers:
        cell_gh = ws_granel[f"{col_l}5"]
        cell_gh.value = text
        cell_gh.font = font_col_header
        cell_gh.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        cell_gh.alignment = align
        cell_gh.border = border_thin
    ws_granel.row_dimensions[5].height = 24

    # Agregación de Graneles
    graneles_map = {}
    for g in data.get("graneles", []):
        g_code = g["granel_code"]
        graneles_map[g_code] = {
            "code": g_code,
            "desc": g.get("granel_desc") or f"GRANEL {g_code}",
            "stock_tanque": float(g.get("granel_stock", 0) or 0),
            "stock_envasado_l": 0.0,
            "cartera_l": 0.0,
            "oct_l": 0.0,
            "nov_l": 0.0,
            "dic_l": 0.0
        }

    for b in data.get("base_skus", []):
        g_code = b.get("granel_code")
        if not g_code:
            continue
        if g_code not in graneles_map:
            graneles_map[g_code] = {
                "code": g_code,
                "desc": b.get("granel_desc") or f"GRANEL {g_code}",
                "stock_tanque": float(b.get("granel_stock", 0) or 0),
                "stock_envasado_l": 0.0,
                "cartera_l": 0.0,
                "oct_l": 0.0,
                "nov_l": 0.0,
                "dic_l": 0.0
            }
        
        env_l = parse_envase_litros(b.get("envase"), b.get("desc"))
        st_u = float(b.get("stock_actual", 0) or 0)
        pe_u = float(b.get("pedidos_pendientes", 0) or 0)
        
        graneles_map[g_code]["stock_envasado_l"] += st_u * env_l
        graneles_map[g_code]["cartera_l"] += pe_u * env_l
        graneles_map[g_code]["oct_l"] += float(b.get("monthly_forecast", {}).get("2026-10", 0) or 0) * env_l
        graneles_map[g_code]["nov_l"] += float(b.get("monthly_forecast", {}).get("2026-11", 0) or 0) * env_l
        graneles_map[g_code]["dic_l"] += float(b.get("monthly_forecast", {}).get("2026-12", 0) or 0) * env_l

    # Ordenar graneles por necesidad de fabricación descendente
    graneles_list = list(graneles_map.values())
    for g in graneles_list:
        g["dem_tot_l"] = g["cartera_l"] + g["oct_l"] + g["nov_l"] + g["dic_l"]
        g["stock_tot_l"] = g["stock_tanque"] + g["stock_envasado_l"]
        g["balance_l"] = g["stock_tot_l"] - g["dem_tot_l"]
        g["falta_fabricar_l"] = max(0.0, -g["balance_l"])

    graneles_list.sort(key=lambda x: (-x["falta_fabricar_l"], -x["dem_tot_l"]))

    cur_g_row = 6
    start_g_row = cur_g_row

    for g in graneles_list:
        # Filtrar graneles sin movimiento (sin stock y sin demanda)
        if g["stock_tot_l"] == 0 and g["dem_tot_l"] == 0:
            continue

        fill_row = fill_alt if (cur_g_row % 2 == 0) else fill_white
        
        ws_granel.cell(cur_g_row, 1, g["code"]).font = font_data_bold
        ws_granel.cell(cur_g_row, 1).alignment = Alignment(horizontal="center")
        
        ws_granel.cell(cur_g_row, 2, g["desc"]).font = font_data_bold
        ws_granel.cell(cur_g_row, 2).alignment = Alignment(horizontal="left")

        g_nums = [
            (3, g["stock_tanque"]),
            (4, g["stock_envasado_l"]),
            (5, g["stock_tot_l"]),
            (6, g["cartera_l"]),
            (7, g["oct_l"]),
            (8, g["nov_l"]),
            (9, g["dic_l"]),
            (10, g["dem_tot_l"]),
            (11, g["balance_l"]),
            (12, g["falta_fabricar_l"])
        ]

        for col_i, val_n in g_nums:
            c_g = ws_granel.cell(cur_g_row, col_i, val_n)
            c_g.number_format = '#,##0'
            c_g.alignment = Alignment(horizontal="right")
            if col_i in [5, 10, 12]:
                c_g.font = font_data_bold
            else:
                c_g.font = font_data

        # Prioridad Reactor
        c_prio = ws_granel.cell(cur_g_row, 13)
        c_prio.alignment = Alignment(horizontal="center")
        if g["falta_fabricar_l"] > 0:
            c_prio.value = "FABRICAR GRANEL"
            c_prio.fill = fill_crit
            c_prio.font = font_crit
        else:
            c_prio.value = "CUBIERTO"
            c_prio.fill = fill_cov
            c_prio.font = font_cov

        for col_idx in range(1, 13):
            cell_x = ws_granel.cell(cur_g_row, col_idx)
            cell_x.fill = fill_row
            cell_x.border = border_thin
        ws_granel.cell(cur_g_row, 13).border = border_thin

        ws_granel.row_dimensions[cur_g_row].height = 20
        cur_g_row += 1

    # Fila de Totales Hoja 2
    tot_g_row = cur_g_row
    ws_granel.merge_cells(f"A{tot_g_row}:B{tot_g_row}")
    c_tot_g_lbl = ws_granel[f"A{tot_g_row}"]
    c_tot_g_lbl.value = "TOTAL PLANTA GRANELES (L/KG)"
    c_tot_g_lbl.font = font_total
    c_tot_g_lbl.alignment = Alignment(horizontal="right", vertical="center")

    sum_g_cols = [
        (3, "C"), (4, "D"), (5, "E"), (6, "F"),
        (7, "G"), (8, "H"), (9, "I"), (10, "J"),
        (11, "K"), (12, "L")
    ]
    for c_idx, letter in sum_g_cols:
        cell_s = ws_granel.cell(tot_g_row, c_idx)
        cell_s.value = f"=SUM({letter}{start_g_row}:{letter}{tot_g_row - 1})"
        cell_s.number_format = '#,##0'
        cell_s.font = font_total
        cell_s.alignment = Alignment(horizontal="right", vertical="center")

    ws_granel.cell(tot_g_row, 13, "").alignment = Alignment(horizontal="center")

    for col_idx in range(1, 14):
        cell_t = ws_granel.cell(tot_g_row, col_idx)
        cell_t.fill = fill_total
        cell_t.border = border_total

    ws_granel.row_dimensions[tot_g_row].height = 26

    # Anchos de Columna Hoja 2
    col_widths_granel = {
        'A': 14,   # Cód. Granel
        'B': 36,   # Producto Químico / Granel
        'C': 20,   # Stock Tanque (L/Kg)
        'D': 22,   # Stock Envasado (L/Kg)
        'E': 20,   # Stock Total (L/Kg)
        'F': 22,   # Cartera Inmediata (L/Kg)
        'G': 18,   # Oct-26 (L/Kg)
        'H': 18,   # Nov-26 (L/Kg)
        'I': 18,   # Dic-26 (L/Kg)
        'J': 22,   # Demanda Total (L/Kg)
        'K': 20,   # Balance Neto (L/Kg)
        'L': 24,   # FALTA FABRICAR (L/Kg)
        'M': 20    # Prioridad Reactor
    }
    for col_l, width in col_widths_granel.items():
        ws_granel.column_dimensions[col_l].width = width

    ws_granel.freeze_panes = "A6"
    ws_granel.auto_filter.ref = f"A5:M{tot_g_row - 1}"

    # =========================================================================
    # HOJA 3: RESUMEN POR LÍNEA DE ENVASADO
    # =========================================================================
    ws_env = wb.create_sheet(title="Resumen por Envase")
    ws_env.views.sheetView[0].showGridLines = True
    
    ws_env.merge_cells("A1:I1")
    c_env_title = ws_env["A1"]
    c_env_title.value = "  CODIAGRO S.L.  |  RESUMEN DE NECESIDADES POR LÍNEA DE ENVASADO Y FORMATO"
    c_env_title.font = font_main_title
    c_env_title.fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
    c_env_title.alignment = Alignment(vertical="center", horizontal="left")
    ws_env.row_dimensions[1].height = 28

    env_headers = [
        ("A", "Línea / Formato de Envasado", 26, Alignment(horizontal="left")),
        ("B", "Nº Registros", 14, Alignment(horizontal="center")),
        ("C", "Stock Envasado (u)", 18, Alignment(horizontal="right")),
        ("D", "Cartera Inmediata (u)", 18, Alignment(horizontal="right")),
        ("E", "Oct-26 (u)", 16, Alignment(horizontal="right")),
        ("F", "Nov-26 (u)", 16, Alignment(horizontal="right")),
        ("G", "Dic-26 (u)", 16, Alignment(horizontal="right")),
        ("H", "Demanda Total (u)", 18, Alignment(horizontal="right")),
        ("I", "A Fabricar Sugerido (u)", 22, Alignment(horizontal="right"))
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

    from collections import defaultdict
    format_agg = defaultdict(lambda: {"count": 0, "stock": 0, "pending": 0, "oct": 0, "nov": 0, "dic": 0, "demand": 0, "sug": 0})
    for r in rows_data:
        fmt = r["formato"]
        format_agg[fmt]["count"] += 1
        format_agg[fmt]["stock"] += r["stock_asignado"]
        format_agg[fmt]["pending"] += r["pending"]
        format_agg[fmt]["oct"] += r["oct_u"]
        format_agg[fmt]["nov"] += r["nov_u"]
        format_agg[fmt]["dic"] += r["dic_u"]
        format_agg[fmt]["demand"] += r["demand_u"]
        format_agg[fmt]["sug"] += r["sug_u"]

    cur_env_row = 4
    start_env_row = cur_env_row

    # Ordenar líneas de envasado lógicamente: Depósito 1000L, Depósito 200L, Garrafas 20L, Garrafas 5L, Botellas 1L, Otros
    order_priority = {
        "Depósito 1000 L": 1,
        "Depósito 200 L": 2,
        "Garrafas 20L": 3,
        "Garrafas 5L": 4,
        "Botellas 1L": 5,
        "Botellas 0.5L": 6
    }
    sorted_formats = sorted(format_agg.keys(), key=lambda x: order_priority.get(x, 99))

    for fmt_name in sorted_formats:
        agg = format_agg[fmt_name]
        ws_env.cell(cur_env_row, 1, fmt_name).font = font_data_bold
        ws_env.cell(cur_env_row, 1).alignment = Alignment(horizontal="left")
        
        ws_env.cell(cur_env_row, 2, agg["count"]).font = font_data
        ws_env.cell(cur_env_row, 2).alignment = Alignment(horizontal="center")
        
        vals_env = [
            (3, agg["stock"]),
            (4, agg["pending"]),
            (5, agg["oct"]),
            (6, agg["nov"]),
            (7, agg["dic"]),
            (8, agg["demand"]),
            (9, agg["sug"])
        ]

        for c_i, val_n in vals_env:
            c_val = ws_env.cell(cur_env_row, c_i, val_n)
            c_val.number_format = '#,##0'
            c_val.font = font_data_bold if c_i in [8, 9] else font_data
            c_val.alignment = Alignment(horizontal="right")

        for col_i in range(1, 10):
            ws_env.cell(cur_env_row, col_i).border = border_thin
            ws_env.cell(cur_env_row, col_i).fill = fill_alt if (cur_env_row % 2 == 0) else fill_white

        ws_env.row_dimensions[cur_env_row].height = 20
        cur_env_row += 1

    ws_env.cell(cur_env_row, 1, "TOTAL PLANTA ENVASADO").font = font_total
    ws_env.cell(cur_env_row, 1).alignment = Alignment(horizontal="left")
    ws_env.cell(cur_env_row, 2, f"=SUM(B{start_env_row}:B{cur_env_row - 1})").font = font_total
    ws_env.cell(cur_env_row, 2).alignment = Alignment(horizontal="center")
    
    for c_i, letter in [(3, "C"), (4, "D"), (5, "E"), (6, "F"), (7, "G"), (8, "H"), (9, "I")]:
        c_val = ws_env.cell(cur_env_row, c_i, f"=SUM({letter}{start_env_row}:{letter}{cur_env_row - 1})")
        c_val.number_format = '#,##0'
        c_val.font = font_total
        c_val.alignment = Alignment(horizontal="right")

    for col_i in range(1, 10):
        ws_env.cell(cur_env_row, col_i).fill = fill_total
        ws_env.cell(cur_env_row, col_i).border = border_total
    ws_env.row_dimensions[cur_env_row].height = 24

    wb.save(output_path)
    print(f"¡Orden de Fabricación MRP generada exitosamente en: {output_path}!")
    return output_path

if __name__ == "__main__":
    generar_excel_orden_fabricacion()
