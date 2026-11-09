import os
import sys

# Inclusión obligatoria y segura de output_staging
_scripts_dir = os.path.dirname(os.path.abspath(__file__))
if _scripts_dir not in sys.path:
    sys.path.insert(0, _scripts_dir)

try:
    from output_staging import get_dest, get_source
except Exception as _e:
    if os.environ.get("CODIAGRO_STAGING_DIR"):
        sys.stderr.write(f"[ERROR FATAL] Fallo al importar output_staging con CODIAGRO_STAGING_DIR activa: {_e}\n")
        sys.exit(1)
    def get_dest(p): return p
    def get_source(p): return p
"""
Generador Profesional de Orden de Fabricación, Planificación MRP y Rotación (Slow/NoMovers) en Excel (XLSX).
Produce un archivo corporativo Codiagro con 4 Hojas Maestras:
1. Hoja "Orden de Fabricación": Desglose individual de cada artículo y país/mercado con su 'Nuevo SKU'
   específico, unidades (u) físicas y litros (L) por país, con columnas mensuales Oct-26, Nov-26, Dic-26.
   La exportación se cataloga como "Sobre Pedido".
2. Hoja "Plan Graneles por Mes": Resumen maestro de semielaborados químicos a fabricar en reactores (L/Kg),
   con los Graneles en filas y los Meses en columnas.
3. Hoja "Resumen por Envase": Clasificación estricta de líneas de envasado (Depósito 1000 L, Depósito 200 L,
   Garrafas 20L, Garrafas 5L, Botellas 1L, etc.) en unidades (u).
4. Hoja "SlowMovers & NoMovers": Análisis de rotación de stock envasado calculado EXCLUSIVAMENTE sobre
   las previsiones del MERCADO NACIONAL (la exportación se fabrica siempre sobre pedido).
"""

import json
import os
import re
import math
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

def is_kg_product(desc="", sku="", fam=None):
    if fam in [38, 39, 42, 43]:
        return True
    text = f"{desc} {sku}".upper()
    if re.search(r'\b(KG|KILOS?|KGS)\b', text):
        return True
    solid_brands = ['ALCAPLANT', 'CODIORGAN A-50', 'CODIORGAN-A50', 'SALWAX-CA', 'SALWAX CA', 'DEFENS-CA', 'BITTER-CAL', 'MAGNIFIC', 'PREZINC', 'BR-59', 'BIOCULTURE BASI']
    if any(b in text for b in solid_brands):
        return True
    return False

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

def get_packaging_format(env_val, desc="", sku="", fam=None):
    l_val = parse_envase_litros(env_val, desc)
    is_solid = is_kg_product(desc, sku, fam)
    
    if is_solid:
        # Todos los kg van en bolsas (o Big Bag si >= 500)
        if l_val >= 500:
            return f"Big Bag {int(l_val) if l_val.is_integer() else l_val} Kg"
        elif l_val == int(l_val):
            return f"Bolsas {int(l_val)} Kg"
        else:
            return f"Bolsas {l_val} Kg"
    else:
        # Formatos líquidos
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
        elif l_val == int(l_val):
            return f"Botellas {int(l_val)}L"
        else:
            return f"Botellas {l_val}L"
    return "Otros Formatos"

def generar_excel_orden_fabricacion(
    json_path="web_dashboard/production_planning_data.json",
    output_path="Orden_Fabricacion_MRP_Codiagro.xlsx",
    scope="ALL",
    horizon_preset="REST_2026",
    batch_size=0,
    format_filter="ALL"
):
    json_path = get_source(json_path)
    if not os.path.exists(json_path):
        json_path = "production_planning_data.json"
    
    with open(json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    meta = data.get("metadata", {})
    presets = meta.get("presets", {})
    selected_months = presets.get(horizon_preset, presets.get("REST_2026", ["2026-11", "2026-12"]))
    
    months_dict = {m["key"]: m.get("label_es", m["key"]) for m in meta.get("months", [])}
    horizon_label = " + ".join([months_dict.get(m, m) for m in selected_months])

    wb = openpyxl.Workbook()
    
    # ---------------- PALETA CORPORATIVA CODIAGRO ----------------
    COLOR_HEADER_BG = "0F172A"       # Slate 900
    COLOR_TITLE_TEXT = "FFFFFF"
    
    COLOR_GRP_ID = "1E3A8A"          # Azul cobalto oscuro (Identificación)
    COLOR_GRP_DEMAND = "334155"      # Pizarra oscuro (Demanda & Fechas)
    COLOR_GRP_PLAN = "065F46"        # Verde bosque oscuro (Planificación MRP)
    COLOR_GRP_ROT = "7C2D12"         # Ámbar/Rojo oscuro (Rotación)
    
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
    COLOR_INFO_BG = "E0F2FE"         # Azul suave
    COLOR_INFO_FG = "0369A1"

    # Fuentes
    font_main_title = Font(name="Calibri", size=13, bold=True, color=COLOR_TITLE_TEXT)
    font_sub_title = Font(name="Calibri", size=9, italic=True, color="94A3B8")
    font_kpi_label = Font(name="Calibri", size=8, bold=True, color="64748B")
    font_kpi_val = Font(name="Calibri", size=12, bold=True, color="0F172A")
    font_kpi_val_red = Font(name="Calibri", size=12, bold=True, color="DC2626")
    font_kpi_val_amber = Font(name="Calibri", size=12, bold=True, color="D97706")
    font_kpi_val_green = Font(name="Calibri", size=12, bold=True, color="059669")
    font_kpi_val_blue = Font(name="Calibri", size=12, bold=True, color="0284C7")
    
    font_grp_header = Font(name="Calibri", size=9.5, bold=True, color="FFFFFF")
    font_col_header = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    font_data = Font(name="Calibri", size=9, color="1E293B")
    font_data_bold = Font(name="Calibri", size=9, bold=True, color="0F172A")
    font_data_purple = Font(name="Calibri", size=9, bold=True, color="6B21A8")
    font_data_mono = Font(name="Consolas", size=9, color="0F172A")
    font_data_mono_bold = Font(name="Consolas", size=9, bold=True, color="0F172A")
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
    fill_info = PatternFill(start_color=COLOR_INFO_BG, end_color=COLOR_INFO_BG, fill_type="solid")
    fill_total = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")

    font_crit = Font(name="Calibri", size=8.5, bold=True, color=COLOR_CRITICAL_FG)
    font_warn = Font(name="Calibri", size=8.5, bold=True, color=COLOR_WARNING_FG)
    font_cov = Font(name="Calibri", size=8.5, bold=True, color=COLOR_COVERED_FG)
    font_info = Font(name="Calibri", size=8.5, bold=True, color=COLOR_INFO_FG)

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
    c_sub.value = f"  Horizonte: {horizon_label} ({len(selected_months)} meses)  |  Ámbito: {scope_str}  |  Criterio Producción: Descuento stock inicial en Nacional  |  Fecha Emisión: {now_str}"
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
        fmt_label = b.get("formato_label") or get_packaging_format(env_str, desc, base_sku)
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
                "is_nacional": b.get("has_nacional", True),
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

        # PRIORIDAD PRODUCCIÓN (Criterio S&OP):
        # El mercado Nacional absorbe primero el stock inicial disponible para evitar sobreproducción.
        active_destinos.sort(key=lambda x: 0 if (x.get("is_nacional") or x.get("pais", "").upper() in ["ESPAÑA", "ESP", ""]) else 1)

        rem_stock = stock_total_b
        tot_stock_envasado += stock_total_b

        for d in active_destinos:
            # Los valores en el JSON ya son UNIDADES FÍSICAS (u)
            p_u = float(d.get("pendientes_uds", 0) or 0)
            nov_u = float(d.get("monthly_forecast", {}).get("2026-11", 0) or 0)
            dic_u = float(d.get("monthly_forecast", {}).get("2026-12", 0) or 0)
            
            dem_period_u = sum(float(d.get("monthly_forecast", {}).get(m, 0) or 0) for m in selected_months)
            
            # Criterio S&OP: Si hay previsión en el horizonte (Nov-Dic), esa es la demanda planificada.
            # Si un artículo no tiene previsión futura pero tiene cartera pendiente inmediata, se toma la cartera.
            total_dem_u = dem_period_u if dem_period_u > 0 else p_u

            # Stock asignado
            allocated_stock_u = min(rem_stock, total_dem_u)
            rem_stock -= allocated_stock_u
            
            falta_envasar_u = max(0.0, total_dem_u - allocated_stock_u)

            # Lote mínimo si procede
            sug_fab_u = falta_envasar_u
            if batch_size > 0 and falta_envasar_u > 0:
                num_batches = math.ceil(falta_envasar_u / batch_size)
                sug_fab_u = num_batches * batch_size

            # Las demandas y órdenes están expresadas directamente en Litros o Kilos (L/Kg)
            total_dem_l = round(total_dem_u, 2)
            sug_fab_l = round(sug_fab_u, 2)
            falta_envasar_l = round(falta_envasar_u, 2)

            is_nac = bool(d.get("is_nacional", False)) or (d.get("pais", "").upper() in ["ESPAÑA", "ESP", ""])
            tipo_fab = "Para Previsión (Nac)" if is_nac else "Sobre Pedido (Export)"

            # Estado y acción
            if falta_envasar_u > 0:
                if not is_nac:
                    action_label = "SOBRE PEDIDO (EXP)"
                    status = "WARNING" if granel_stock >= falta_envasar_l else "CRITICAL"
                    if status == "CRITICAL":
                        count_critical += 1
                    else:
                        count_warning += 1
                else:
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
                cov_pct = min(1.0, allocated_stock_u / total_dem_u)
            elif allocated_stock_u == 0:
                cov_pct = 0.0

            pais_nombre = d.get("pais", "ESPAÑA")
            nuevo_sku = d.get("sku", base_sku)

            tot_unidades_fabricar += sug_fab_u
            tot_litros_fabricar += sug_fab_l
            tot_pedidos_cartera += p_u

            rows_data.append({
                "base_sku": base_sku,
                "nuevo_sku": nuevo_sku,
                "desc": desc,
                "envase": env_l,
                "formato": fmt_label,
                "pais": pais_nombre,
                "tipo_fab": tipo_fab,
                "stock_asignado": allocated_stock_u,
                "pending": p_u,
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
        ("A4", "B4", "A5", "B5", "TOTAL A ENVASAR (L/Kg)", f"{tot_litros_fabricar:,.0f} L".replace(",", "."), font_kpi_val_red),
        ("C4", "D4", "C5", "D5", "TOTAL UNIDADES A ENVASAR", f"{tot_unidades_fabricar:,.0f} u".replace(",", "."), font_kpi_val_red),
        ("E4", "F4", "E5", "F5", "DÉFICIT QUÍMICO URGENTE", f"{count_critical} Artículos", font_kpi_val_red),
        ("G4", "H4", "G5", "H5", "SOLO ENVASAR / SOBRE PEDIDO", f"{count_warning} Artículos", font_kpi_val_amber),
        ("I4", "K4", "I5", "K5", "PEDIDOS CARTERA INMEDIATOS", f"{tot_pedidos_cartera:,.0f} u".replace(",", "."), font_kpi_val),
        ("L4", "P4", "L5", "P5", "STOCK ENVASADO ALMACÉN", f"{tot_stock_envasado:,.0f} u".replace(",", "."), font_kpi_val_green)
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

    # Encabezados de grupos en Hoja 1 (Fila 7)
    ws_main.merge_cells("A7:G7")
    c_grp1 = ws_main["A7"]
    c_grp1.value = "1. IDENTIFICACIÓN Y ESPECIFICACIÓN ARTÍCULO / DESTINO"
    c_grp1.font = font_grp_header
    c_grp1.fill = PatternFill(start_color=COLOR_GRP_ID, end_color=COLOR_GRP_ID, fill_type="solid")
    c_grp1.alignment = Alignment(horizontal="center", vertical="center")

    ws_main.merge_cells("H7:L7")
    c_grp2 = ws_main["H7"]
    c_grp2.value = "2. DISPONIBILIDAD Y DEMANDA HORIZONTE (UNIDADES FÍSICAS)"
    c_grp2.font = font_grp_header
    c_grp2.fill = PatternFill(start_color=COLOR_GRP_DEMAND, end_color=COLOR_GRP_DEMAND, fill_type="solid")
    c_grp2.alignment = Alignment(horizontal="center", vertical="center")

    ws_main.merge_cells("M7:P7")
    c_grp3 = ws_main["M7"]
    c_grp3.value = "3. PLANIFICACIÓN DE PRODUCCIÓN Y ENVASADO"
    c_grp3.font = font_grp_header
    c_grp3.fill = PatternFill(start_color=COLOR_GRP_PLAN, end_color=COLOR_GRP_PLAN, fill_type="solid")
    c_grp3.alignment = Alignment(horizontal="center", vertical="center")

    ws_main.row_dimensions[7].height = 20

    # Encabezados de Columnas (Fila 8)
    col_headers = [
        ("A", "Cód. Base", COLOR_GRP_ID, Alignment(horizontal="center")),
        ("B", "Nuevo SKU (País)", COLOR_GRP_ID, Alignment(horizontal="center")),
        ("C", "Producto / Descripción", COLOR_GRP_ID, Alignment(horizontal="left")),
        ("D", "Env. (L)", COLOR_GRP_ID, Alignment(horizontal="right")),
        ("E", "Línea Envasado", COLOR_GRP_ID, Alignment(horizontal="left")),
        ("F", "País / Mercado", COLOR_GRP_ID, Alignment(horizontal="center")),
        ("G", "Tipo Fabricación", COLOR_GRP_ID, Alignment(horizontal="center")),
        
        ("H", "Stock Asignado (u)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("I", "Cartera Inmediata (u)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("J", "Nov-26 (u)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("K", "Dic-26 (u)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("L", "Demanda Total (u)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        
        ("M", "Demanda Total (L/Kg)", COLOR_GRP_PLAN, Alignment(horizontal="right")),
        ("N", "A Envasar (u)", COLOR_GRP_PLAN, Alignment(horizontal="right")),
        ("O", "A Envasar (L/Kg)", COLOR_GRP_PLAN, Alignment(horizontal="right")),
        ("P", "Prioridad / Acción", COLOR_GRP_PLAN, Alignment(horizontal="center"))
    ]

    for col_l, text, bg_color, align in col_headers:
        cell_h = ws_main[f"{col_l}8"]
        cell_h.value = text
        cell_h.font = font_col_header
        cell_h.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        cell_h.alignment = align
        cell_h.border = border_thin
    ws_main.row_dimensions[8].height = 24

    # Rellenar Datos Hoja 1
    cur_row = 9
    start_row = cur_row

    for r in rows_data:
        fill_row = fill_alt if (cur_row % 2 == 0) else fill_white

        ws_main.cell(cur_row, 1, r["base_sku"]).font = font_data_bold
        ws_main.cell(cur_row, 1).alignment = Alignment(horizontal="center")
        
        c_nsku = ws_main.cell(cur_row, 2, r["nuevo_sku"])
        c_nsku.font = font_data_purple if (r["nuevo_sku"] != r["base_sku"]) else font_data
        c_nsku.alignment = Alignment(horizontal="center")

        ws_main.cell(cur_row, 3, r["desc"]).font = font_data
        ws_main.cell(cur_row, 3).alignment = Alignment(horizontal="left")

        c_env = ws_main.cell(cur_row, 4, r["envase"])
        c_env.number_format = '0.00'
        c_env.font = font_data
        c_env.alignment = Alignment(horizontal="right")

        ws_main.cell(cur_row, 5, r["formato"]).font = font_data
        ws_main.cell(cur_row, 5).alignment = Alignment(horizontal="left")

        ws_main.cell(cur_row, 6, r["pais"]).font = font_data
        ws_main.cell(cur_row, 6).alignment = Alignment(horizontal="center")

        c_tf = ws_main.cell(cur_row, 7, r["tipo_fab"])
        c_tf.font = font_info if "Sobre Pedido" in r["tipo_fab"] else font_data
        c_tf.alignment = Alignment(horizontal="center")

        # Numéricos (Cols 8 a 15)
        nums = [
            (8, r["stock_asignado"]),
            (9, r["pending"]),
            (10, r["nov_u"]),
            (11, r["dic_u"]),
            (12, r["demand_u"]),
            (13, r["demand_l"]),
            (14, r["sug_u"]),
            (15, r["sug_l"])
        ]

        for col_i, val_n in nums:
            c_num = ws_main.cell(cur_row, col_i, val_n)
            c_num.number_format = '#,##0'
            c_num.alignment = Alignment(horizontal="right")
            if col_i in [12, 13, 14, 15]:
                c_num.font = font_data_bold
            else:
                c_num.font = font_data

        # Acción / Estado (Col 16)
        c_act = ws_main.cell(cur_row, 16, r["action"])
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
            cell_x = ws_main.cell(cur_row, col_idx)
            cell_x.fill = fill_row
            cell_x.border = border_thin
        ws_main.cell(cur_row, 16).border = border_thin

        ws_main.row_dimensions[cur_row].height = 19
        cur_row += 1

    # Fila de Totales Hoja 1
    tot_row = cur_row
    ws_main.merge_cells(f"A{tot_row}:G{tot_row}")
    c_tot_lbl = ws_main[f"A{tot_row}"]
    c_tot_lbl.value = "TOTAL GENERAL (UNIDADES Y VOLUMEN)"
    c_tot_lbl.font = font_total
    c_tot_lbl.alignment = Alignment(horizontal="right", vertical="center")

    sum_cols = [
        (8, "H"), (9, "I"), (10, "J"), (11, "K"),
        (12, "L"), (13, "M"), (14, "N"), (15, "O")
    ]
    for c_idx, letter in sum_cols:
        cell_s = ws_main.cell(tot_row, c_idx)
        cell_s.value = f"=SUM({letter}{start_row}:{letter}{tot_row - 1})"
        cell_s.number_format = '#,##0'
        cell_s.font = font_total
        cell_s.alignment = Alignment(horizontal="right", vertical="center")

    ws_main.cell(tot_row, 16, "").alignment = Alignment(horizontal="center")

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
        'G': 22,   # Tipo Fabricación
        'H': 17,   # Stock Asignado (u)
        'I': 17,   # Cartera Inmediata (u)
        'J': 14,   # Nov-26 (u)
        'K': 14,   # Dic-26 (u)
        'L': 18,   # Demanda Total (u)
        'M': 20,   # Demanda Total (L/Kg)
        'N': 17,   # A Envasar (u)
        'O': 20,   # A Envasar (L/Kg)
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
    ws_granel.merge_cells("A1:L1")
    c_g_title = ws_granel["A1"]
    c_g_title.value = "  CODIAGRO S.L.  |  PLAN MAESTRO DE PRODUCCIÓN DE GRANELES (NECESIDAD EN LITROS / KG POR MES)"
    c_g_title.font = font_main_title
    c_g_title.fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
    c_g_title.alignment = Alignment(vertical="center", horizontal="left")
    ws_granel.row_dimensions[1].height = 28

    ws_granel.merge_cells("A2:L2")
    c_g_sub = ws_granel["A2"]
    c_g_sub.value = f"  Horizonte: Noviembre 2026 - Diciembre 2026 (2 meses)  |  Criterio Producción: Demanda neta reactores tras tanques  |  Fecha Emisión: {now_str}"
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

    ws_granel.merge_cells("F4:I4")
    c_g_g3 = ws_granel["F4"]
    c_g_g3.value = "3. DEMANDA HORIZONTE NOV-DIC (L/KG)"
    c_g_g3.font = font_grp_header
    c_g_g3.fill = PatternFill(start_color=COLOR_GRP_DEMAND, end_color=COLOR_GRP_DEMAND, fill_type="solid")
    c_g_g3.alignment = Alignment(horizontal="center", vertical="center")

    ws_granel.merge_cells("J4:L4")
    c_g_g4 = ws_granel["J4"]
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
        ("G", "Nov-26 (L/Kg)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("H", "Dic-26 (L/Kg)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("I", "Demanda Total (L/Kg)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        
        ("J", "Balance Neto (L/Kg)", COLOR_GRP_PLAN, Alignment(horizontal="right")),
        ("K", "FALTA FABRICAR (L/Kg)", COLOR_GRP_PLAN, Alignment(horizontal="right")),
        ("L", "Prioridad Reactor", COLOR_GRP_PLAN, Alignment(horizontal="center"))
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
                "nov_l": 0.0,
                "dic_l": 0.0
            }
        
        env_l = parse_envase_litros(b.get("envase"), b.get("desc"))
        st_u = float(b.get("stock_actual", 0) or 0)
        pe_u = float(b.get("pedidos_pendientes", 0) or 0)
        nov_u = float(b.get("monthly_forecast", {}).get("2026-11", 0) or 0)
        dic_u = float(b.get("monthly_forecast", {}).get("2026-12", 0) or 0)
        
        # En graneles la demanda y el stock ya están expresados en Litros o Kilos (L/Kg)
        graneles_map[g_code]["stock_envasado_l"] += round(st_u, 2)
        graneles_map[g_code]["cartera_l"] += round(pe_u, 2)
        graneles_map[g_code]["nov_l"] += round(nov_u, 2)
        graneles_map[g_code]["dic_l"] += round(dic_u, 2)

    graneles_list = list(graneles_map.values())
    for g in graneles_list:
        g["dem_tot_l"] = g["nov_l"] + g["dic_l"]
        g["stock_tot_l"] = g["stock_tanque"] + g["stock_envasado_l"]
        g["balance_l"] = g["stock_tot_l"] - g["dem_tot_l"]
        g["falta_fabricar_l"] = max(0.0, g["dem_tot_l"] - g["stock_tanque"])

    graneles_list.sort(key=lambda x: (-x["falta_fabricar_l"], -x["dem_tot_l"]))

    cur_g_row = 6
    start_g_row = cur_g_row

    for g in graneles_list:
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
            (7, g["nov_l"]),
            (8, g["dic_l"]),
            (9, g["dem_tot_l"]),
            (10, g["balance_l"]),
            (11, g["falta_fabricar_l"])
        ]

        for col_i, val_n in g_nums:
            c_g = ws_granel.cell(cur_g_row, col_i, val_n)
            c_g.number_format = '#,##0'
            c_g.alignment = Alignment(horizontal="right")
            if col_i in [5, 9, 11]:
                c_g.font = font_data_bold
            else:
                c_g.font = font_data

        # Prioridad Reactor (Col 12)
        c_prio = ws_granel.cell(cur_g_row, 12)
        c_prio.alignment = Alignment(horizontal="center")
        if g["falta_fabricar_l"] > 0:
            c_prio.value = "FABRICAR GRANEL"
            c_prio.fill = fill_crit
            c_prio.font = font_crit
        else:
            c_prio.value = "CUBIERTO"
            c_prio.fill = fill_cov
            c_prio.font = font_cov

        for col_idx in range(1, 12):
            cell_x = ws_granel.cell(cur_g_row, col_idx)
            cell_x.fill = fill_row
            cell_x.border = border_thin
        ws_granel.cell(cur_g_row, 12).border = border_thin

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
        (11, "K")
    ]
    for c_idx, letter in sum_g_cols:
        cell_s = ws_granel.cell(tot_g_row, c_idx)
        cell_s.value = f"=SUM({letter}{start_g_row}:{letter}{tot_g_row - 1})"
        cell_s.number_format = '#,##0'
        cell_s.font = font_total
        cell_s.alignment = Alignment(horizontal="right", vertical="center")

    ws_granel.cell(tot_g_row, 12, "").alignment = Alignment(horizontal="center")

    for col_idx in range(1, 13):
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
        'G': 18,   # Nov-26 (L/Kg)
        'H': 18,   # Dic-26 (L/Kg)
        'I': 22,   # Demanda Total (L/Kg)
        'J': 20,   # Balance Neto (L/Kg)
        'K': 24,   # FALTA FABRICAR (L/Kg)
        'L': 20    # Prioridad Reactor
    }
    for col_l, width in col_widths_granel.items():
        ws_granel.column_dimensions[col_l].width = width

    ws_granel.freeze_panes = "A6"
    ws_granel.auto_filter.ref = f"A5:L{tot_g_row - 1}"

    # =========================================================================
    # HOJA 3: RESUMEN POR LÍNEA DE ENVASADO
    # =========================================================================
    ws_env = wb.create_sheet(title="Resumen por Envase")
    ws_env.views.sheetView[0].showGridLines = True
    
    ws_env.merge_cells("A1:H1")
    c_env_title = ws_env["A1"]
    c_env_title.value = "  CODIAGRO S.L.  |  RESUMEN DE NECESIDADES POR LÍNEA DE ENVASADO Y FORMATO (UNIDADES FÍSICAS)"
    c_env_title.font = font_main_title
    c_env_title.fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
    c_env_title.alignment = Alignment(vertical="center", horizontal="left")
    ws_env.row_dimensions[1].height = 28

    env_headers = [
        ("A", "Línea / Formato de Envasado", 26, Alignment(horizontal="left")),
        ("B", "Nº Registros", 14, Alignment(horizontal="center")),
        ("C", "Stock Envasado (u)", 18, Alignment(horizontal="right")),
        ("D", "Cartera Inmediata (u)", 18, Alignment(horizontal="right")),
        ("E", "Nov-26 (u)", 16, Alignment(horizontal="right")),
        ("F", "Dic-26 (u)", 16, Alignment(horizontal="right")),
        ("G", "Demanda Total (u)", 18, Alignment(horizontal="right")),
        ("H", "A Envasar Sugerido (u)", 22, Alignment(horizontal="right"))
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
    format_agg = defaultdict(lambda: {"count": 0, "stock": 0, "pending": 0, "nov": 0, "dic": 0, "demand": 0, "sug": 0})
    for r in rows_data:
        fmt = r["formato"]
        format_agg[fmt]["count"] += 1
        format_agg[fmt]["stock"] += r["stock_asignado"]
        format_agg[fmt]["pending"] += r["pending"]
        format_agg[fmt]["nov"] += r["nov_u"]
        format_agg[fmt]["dic"] += r["dic_u"]
        format_agg[fmt]["demand"] += r["demand_u"]
        format_agg[fmt]["sug"] += r["sug_u"]

    cur_env_row = 4
    start_env_row = cur_env_row

    order_priority = {
        "Depósito 1000 L": 1,
        "Big Bag 1000 Kg": 2,
        "Depósito 200 L": 3,
        "Garrafas 20L": 4,
        "Bolsas 20 Kg": 5,
        "Garrafas 5L": 6,
        "Bolsas 5 Kg": 7,
        "Botellas 1L": 8,
        "Bolsas 1 Kg": 9,
        "Botellas 0.5L": 10
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
            (5, agg["nov"]),
            (6, agg["dic"]),
            (7, agg["demand"]),
            (8, agg["sug"])
        ]

        for c_i, val_n in vals_env:
            c_val = ws_env.cell(cur_env_row, c_i, val_n)
            c_val.number_format = '#,##0'
            c_val.font = font_data_bold if c_i in [7, 8] else font_data
            c_val.alignment = Alignment(horizontal="right")

        for col_i in range(1, 9):
            ws_env.cell(cur_env_row, col_i).border = border_thin
            ws_env.cell(cur_env_row, col_i).fill = fill_alt if (cur_env_row % 2 == 0) else fill_white

        ws_env.row_dimensions[cur_env_row].height = 20
        cur_env_row += 1

    ws_env.cell(cur_env_row, 1, "TOTAL PLANTA ENVASADO").font = font_total
    ws_env.cell(cur_env_row, 1).alignment = Alignment(horizontal="left")
    ws_env.cell(cur_env_row, 2, f"=SUM(B{start_env_row}:B{cur_env_row - 1})").font = font_total
    ws_env.cell(cur_env_row, 2).alignment = Alignment(horizontal="center")
    
    for c_i, letter in [(3, "C"), (4, "D"), (5, "E"), (6, "F"), (7, "G"), (8, "H")]:
        c_val = ws_env.cell(cur_env_row, c_i, f"=SUM({letter}{start_env_row}:{letter}{cur_env_row - 1})")
        c_val.number_format = '#,##0'
        c_val.font = font_total
        c_val.alignment = Alignment(horizontal="right")

    for col_i in range(1, 9):
        ws_env.cell(cur_env_row, col_i).fill = fill_total
        ws_env.cell(cur_env_row, col_i).border = border_total
    ws_env.row_dimensions[cur_env_row].height = 24

    # =========================================================================
    # HOJA 4: SLOWMOVERS & NOMOVERS (MERCADO NACIONAL)
    # =========================================================================
    ws_rot = wb.create_sheet(title="SlowMovers & NoMovers")
    ws_rot.views.sheetView[0].showGridLines = True

    # Banner Superior Hoja 4 (A1:P2)
    ws_rot.merge_cells("A1:P1")
    c_rot_title = ws_rot["A1"]
    c_rot_title.value = "  CODIAGRO S.L.  |  ANÁLISIS DE ROTACIÓN DE STOCK: SLOW MOVERS & NO MOVERS (MERCADO NACIONAL)"
    c_rot_title.font = font_main_title
    c_rot_title.fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
    c_rot_title.alignment = Alignment(vertical="center", horizontal="left")
    ws_rot.row_dimensions[1].height = 28

    ws_rot.merge_cells("A2:P2")
    c_rot_sub = ws_rot["A2"]
    c_rot_sub.value = f"  Premisa de Gestión: Previsiones exclusivamente del Mercado Nacional (España). Horizonte: Nov-Dic 2026 + Q1-2027. La exportación se fabrica sobre pedido.  |  Fecha: {now_str}"
    c_rot_sub.font = font_sub_title
    c_rot_sub.fill = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
    c_rot_sub.alignment = Alignment(vertical="center", horizontal="left")
    ws_rot.row_dimensions[2].height = 18

    ws_rot.row_dimensions[3].height = 8

    # Procesamiento de datos de rotación desde base_skus
    rot_rows = []
    tot_rot_stock_u = 0
    tot_rot_stock_eur = 0
    tot_rot_nomover_u = 0
    tot_rot_nomover_eur = 0
    tot_rot_slowmover_u = 0
    tot_rot_slowmover_eur = 0
    tot_rot_fastmover_u = 0
    count_nomover_skus = 0
    count_slowmover_skus = 0
    count_fastmover_skus = 0

    for b in data.get("base_skus", []):
        st_u = float(b.get("stock_actual", 0) or 0)
        if st_u <= 0:
            continue
        
        env_l = float(b.get("envase_litros", 1.0) or 1.0)
        st_l = st_u  # Stock físico real ya en Litros o Kilos (L/Kg)
        envases_est = round(st_u / env_l, 1) if env_l > 0 else st_u
        p_eur = float(b.get("precio_unitario", 0) or 0)
        st_eur = round(st_u * p_eur, 2)
        
        d_q4 = float(b.get("demanda_nac_q4", 0) or 0)
        d_q1_27 = float(b.get("demanda_nac_q1_27", 0) or 0)
        d_6m = float(b.get("demanda_nac_6m", 0) or 0)
        
        nm_u = float(b.get("no_mover_qty", 0) or 0)
        sm_u = float(b.get("slow_mover_qty", 0) or 0)
        fm_u = float(b.get("fast_mover_qty", 0) or 0)
        
        nm_l = nm_u
        sm_l = sm_u
        nm_envases = round(nm_u / env_l, 1) if env_l > 0 else nm_u
        sm_envases = round(sm_u / env_l, 1) if env_l > 0 else sm_u
        nm_eur = round(nm_u * p_eur, 2)
        sm_eur = round(sm_u * p_eur, 2)
        
        tot_rot_stock_u += st_u
        tot_rot_stock_eur += st_eur
        tot_rot_nomover_u += nm_u
        tot_rot_nomover_eur += nm_eur
        tot_rot_slowmover_u += sm_u
        tot_rot_slowmover_eur += sm_eur
        tot_rot_fastmover_u += fm_u
        
        # Clasificación principal
        if nm_u > 0:
            clasif = "🔴 NoMover (>6M Nac)"
            status_rot = "NOMOVER"
            count_nomover_skus += 1
            accion_sug = "Plan Choque Comercial Nac / Oferta" if d_6m > 0 else "Sin Demanda Nacional (Inmovilizado Crítico)"
        elif sm_u > 0:
            clasif = "🟡 SlowMover (3-6M Nac)"
            status_rot = "SLOWMOVER"
            count_slowmover_skus += 1
            accion_sug = "Consumo previsto en Q1-27 Nacional"
        else:
            clasif = "🟢 Alta Rotación (0-3M Nac)"
            status_rot = "FASTMOVER"
            count_fastmover_skus += 1
            accion_sug = "Rotación Normal en Campaña Q4"

        rot_rows.append({
            "base_sku": b.get("base_sku", ""),
            "desc": b.get("desc", ""),
            "formato": b.get("formato_label", ""),
            "stock_u": st_u,
            "envases_est": envases_est,
            "d_q4": d_q4,
            "d_q1_27": d_q1_27,
            "d_6m": d_6m,
            "sm_u": sm_u,
            "nm_u": nm_u,
            "nm_envases": nm_envases,
            "clasif": clasif,
            "status_rot": status_rot,
            "precio": p_eur,
            "nm_eur": nm_eur,
            "sm_eur": sm_eur,
            "accion": accion_sug
        })

    # Ordenar: primero NoMovers por mayor valor € inmovilizado, luego SlowMovers, luego Alta Rotación
    rot_prio = {"NOMOVER": 0, "SLOWMOVER": 1, "FASTMOVER": 2}
    rot_rows.sort(key=lambda x: (rot_prio.get(x["status_rot"], 9), -x["nm_eur"], -x["nm_u"], -x["sm_u"]))

    # Tarjetas KPI Hoja 4 (Filas 4 y 5)
    pct_nm = (tot_rot_nomover_u / tot_rot_stock_u * 100) if tot_rot_stock_u > 0 else 0
    pct_sm = (tot_rot_slowmover_u / tot_rot_stock_u * 100) if tot_rot_stock_u > 0 else 0
    pct_fm = (tot_rot_fastmover_u / tot_rot_stock_u * 100) if tot_rot_stock_u > 0 else 0

    kpis_rot = [
        ("A4", "D4", "A5", "D5", "STOCK TOTAL ALMACÉN (L/Kg)", f"{tot_rot_stock_u:,.0f} L/Kg  ({tot_rot_stock_eur:,.0f} €)".replace(",", "."), font_kpi_val_blue),
        ("E4", "H4", "E5", "H5", f"🔴 NOMOVERS (>6M NACIONAL) — {pct_nm:.1f}%", f"{tot_rot_nomover_u:,.0f} L/Kg  |  {tot_rot_nomover_eur:,.2f} €".replace(",", "."), font_kpi_val_red),
        ("I4", "L4", "I5", "L5", f"🟡 SLOWMOVERS (3-6M NACIONAL) — {pct_sm:.1f}%", f"{tot_rot_slowmover_u:,.0f} L/Kg  |  {tot_rot_slowmover_eur:,.2f} €".replace(",", "."), font_kpi_val_amber),
        ("M4", "P4", "M5", "P5", f"🟢 ALTA ROTACIÓN (Q4 NACIONAL) — {pct_fm:.1f}%", f"{tot_rot_fastmover_u:,.0f} L/Kg  ({count_fastmover_skus} SKUs)".replace(",", "."), font_kpi_val_green)
    ]

    for top_l, top_r, bot_l, bot_r, label, val_text, val_font in kpis_rot:
        ws_rot.merge_cells(f"{top_l}:{top_r}")
        ws_rot.merge_cells(f"{bot_l}:{bot_r}")
        
        c_lbl = ws_rot[top_l]
        c_lbl.value = label
        c_lbl.font = font_kpi_label
        c_lbl.alignment = Alignment(horizontal="center", vertical="center")
        
        c_val = ws_rot[bot_l]
        c_val.value = val_text
        c_val.font = val_font
        c_val.alignment = Alignment(horizontal="center", vertical="center")

        start_col = openpyxl.utils.column_index_from_string(top_l[0])
        end_col = openpyxl.utils.column_index_from_string(top_r[0]) if len(top_r) == 2 else openpyxl.utils.column_index_from_string(top_r[:len(top_r)-1])
        row_top = int(top_l[1:])
        row_bot = int(bot_l[1:])
        
        for r_k in range(row_top, row_bot + 1):
            for c_k in range(start_col, end_col + 1):
                cell_k = ws_rot.cell(r_k, c_k)
                cell_k.fill = fill_kpi
                cell_k.border = border_kpi

    ws_rot.row_dimensions[4].height = 16
    ws_rot.row_dimensions[5].height = 24
    ws_rot.row_dimensions[6].height = 10

    # Encabezados de grupos en Hoja 4 (Fila 7)
    ws_rot.merge_cells("A7:C7")
    c_rot_g1 = ws_rot["A7"]
    c_rot_g1.value = "1. IDENTIFICACIÓN ARTÍCULO"
    c_rot_g1.font = font_grp_header
    c_rot_g1.fill = PatternFill(start_color=COLOR_GRP_ID, end_color=COLOR_GRP_ID, fill_type="solid")
    c_rot_g1.alignment = Alignment(horizontal="center", vertical="center")

    ws_rot.merge_cells("D7:E7")
    c_rot_g2 = ws_rot["D7"]
    c_rot_g2.value = "2. STOCK FÍSICO ALMACÉN"
    c_rot_g2.font = font_grp_header
    c_rot_g2.fill = PatternFill(start_color=COLOR_GRP_DEMAND, end_color=COLOR_GRP_DEMAND, fill_type="solid")
    c_rot_g2.alignment = Alignment(horizontal="center", vertical="center")

    ws_rot.merge_cells("F7:G7")
    c_rot_g3 = ws_rot["F7"]
    c_rot_g3.value = "3. PREVISIÓN MERCADO NACIONAL (L/KG)"
    c_rot_g3.font = font_grp_header
    c_rot_g3.fill = PatternFill(start_color=COLOR_GRP_DEMAND, end_color=COLOR_GRP_DEMAND, fill_type="solid")
    c_rot_g3.alignment = Alignment(horizontal="center", vertical="center")

    ws_rot.merge_cells("H7:O7")
    c_rot_g4 = ws_rot["H7"]
    c_rot_g4.value = "4. EVALUACIÓN DE ROTACIÓN Y VALORACIÓN ECONÓMICA"
    c_rot_g4.font = font_grp_header
    c_rot_g4.fill = PatternFill(start_color=COLOR_GRP_ROT, end_color=COLOR_GRP_ROT, fill_type="solid")
    c_rot_g4.alignment = Alignment(horizontal="center", vertical="center")

    ws_rot.row_dimensions[7].height = 20

    rot_headers = [
        ("A", "Cód. Base SKU", COLOR_GRP_ID, Alignment(horizontal="center")),
        ("B", "Descripción Producto", COLOR_GRP_ID, Alignment(horizontal="left")),
        ("C", "Línea / Formato", COLOR_GRP_ID, Alignment(horizontal="left")),
        
        ("D", "Stock Físico (L/Kg)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("E", "Envases Est. (u)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        
        ("F", "(N+2) (L/Kg)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        ("G", "(N+5) (L/Kg)", COLOR_GRP_DEMAND, Alignment(horizontal="right")),
        
        ("H", "SlowMovers (L/Kg)", COLOR_GRP_ROT, Alignment(horizontal="right")),
        ("I", "NoMovers (L/Kg)", COLOR_GRP_ROT, Alignment(horizontal="right")),
        ("J", "Envases NoMover (u)", COLOR_GRP_ROT, Alignment(horizontal="right")),
        ("K", "Clasificación Rotación", COLOR_GRP_ROT, Alignment(horizontal="center")),
        ("L", "Precio Ref. (€/(L·Kg))", COLOR_GRP_ROT, Alignment(horizontal="right")),
        ("M", "Inmovilizado NoMover (€)", COLOR_GRP_ROT, Alignment(horizontal="right")),
        ("N", "Inmovilizado SlowMover (€)", COLOR_GRP_ROT, Alignment(horizontal="right")),
        ("O", "Acción Comercial Recomendada", COLOR_GRP_ROT, Alignment(horizontal="left"))
    ]

    for col_l, text, bg_color, align in rot_headers:
        cell_rh = ws_rot[f"{col_l}8"]
        cell_rh.value = text
        cell_rh.font = font_col_header
        cell_rh.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
        cell_rh.alignment = align
        cell_rh.border = border_thin
    ws_rot.row_dimensions[8].height = 24

    cur_rot_row = 9
    start_rot_row = cur_rot_row

    for r in rot_rows:
        fill_row = fill_alt if (cur_rot_row % 2 == 0) else fill_white
        
        ws_rot.cell(cur_rot_row, 1, r["base_sku"]).font = font_data_mono_bold
        ws_rot.cell(cur_rot_row, 1).alignment = Alignment(horizontal="center")
        
        ws_rot.cell(cur_rot_row, 2, r["desc"]).font = font_data
        ws_rot.cell(cur_rot_row, 2).alignment = Alignment(horizontal="left")
        
        ws_rot.cell(cur_rot_row, 3, r["formato"]).font = font_data
        ws_rot.cell(cur_rot_row, 3).alignment = Alignment(horizontal="left")

        rot_nums = [
            (4, r["stock_u"], '#,##0', font_data_bold),
            (5, r["envases_est"], '#,##0', font_data),
            (6, r["d_q4"], '#,##0', font_data),
            (7, r["d_q1_27"], '#,##0', font_data),
            (8, r["sm_u"], '#,##0', font_data_bold),
            (9, r["nm_u"], '#,##0', font_data_bold),
            (10, r["nm_envases"], '#,##0', font_data),
        ]

        for col_i, val_n, n_fmt, f_style in rot_nums:
            c_r = ws_rot.cell(cur_rot_row, col_i, val_n)
            c_r.number_format = n_fmt
            c_r.font = f_style
            c_r.alignment = Alignment(horizontal="right")

        # Clasificación
        c_cl = ws_rot.cell(cur_rot_row, 11, r["clasif"])
        c_cl.alignment = Alignment(horizontal="center")
        if r["status_rot"] == "NOMOVER":
            c_cl.fill = fill_crit
            c_cl.font = font_crit
        elif r["status_rot"] == "SLOWMOVER":
            c_cl.fill = fill_warn
            c_cl.font = font_warn
        else:
            c_cl.fill = fill_cov
            c_cl.font = font_cov

        # Precios y valores
        c_pr = ws_rot.cell(cur_rot_row, 12, r["precio"])
        c_pr.number_format = '#,##0.00 €'
        c_pr.font = font_data
        c_pr.alignment = Alignment(horizontal="right")

        c_vnm = ws_rot.cell(cur_rot_row, 13, r["nm_eur"])
        c_vnm.number_format = '#,##0.00 €'
        c_vnm.font = font_crit if r["nm_eur"] > 0 else font_data
        c_vnm.alignment = Alignment(horizontal="right")

        c_vsm = ws_rot.cell(cur_rot_row, 14, r["sm_eur"])
        c_vsm.number_format = '#,##0.00 €'
        c_vsm.font = font_warn if r["sm_eur"] > 0 else font_data
        c_vsm.alignment = Alignment(horizontal="right")

        # Acción recomendada
        c_ac = ws_rot.cell(cur_rot_row, 15, r["accion"])
        c_ac.font = font_data
        c_ac.alignment = Alignment(horizontal="left")

        for col_idx in range(1, 11):
            cell_x = ws_rot.cell(cur_rot_row, col_idx)
            cell_x.fill = fill_row
            cell_x.border = border_thin
        ws_rot.cell(cur_rot_row, 11).border = border_thin
        for col_idx in range(12, 16):
            cell_x = ws_rot.cell(cur_rot_row, col_idx)
            cell_x.fill = fill_row
            cell_x.border = border_thin

        ws_rot.row_dimensions[cur_rot_row].height = 20
        cur_rot_row += 1

    # Fila de Totales Hoja 4
    tot_rot_row = cur_rot_row
    ws_rot.merge_cells(f"A{tot_rot_row}:C{tot_rot_row}")
    c_tot_rot_lbl = ws_rot[f"A{tot_rot_row}"]
    c_tot_rot_lbl.value = "TOTAL ROTACIÓN PLANTA (MERCADO NACIONAL)"
    c_tot_rot_lbl.font = font_total
    c_tot_rot_lbl.alignment = Alignment(horizontal="right", vertical="center")

    sum_rot_cols = [
        (4, "D", '#,##0'), (5, "E", '#,##0'), (6, "F", '#,##0'), (7, "G", '#,##0'),
        (8, "H", '#,##0'), (9, "I", '#,##0'), (10, "J", '#,##0'),
        (13, "M", '#,##0.00 €'), (14, "N", '#,##0.00 €')
    ]
    for c_idx, letter, n_fmt in sum_rot_cols:
        cell_s = ws_rot.cell(tot_rot_row, c_idx)
        cell_s.value = f"=SUM({letter}{start_rot_row}:{letter}{tot_rot_row - 1})"
        cell_s.number_format = n_fmt
        cell_s.font = font_total
        cell_s.alignment = Alignment(horizontal="right", vertical="center")

    ws_rot.cell(tot_rot_row, 11, "").alignment = Alignment(horizontal="center")
    ws_rot.cell(tot_rot_row, 12, "").alignment = Alignment(horizontal="center")
    ws_rot.cell(tot_rot_row, 15, "").alignment = Alignment(horizontal="left")

    for col_idx in range(1, 16):
        cell_t = ws_rot.cell(tot_rot_row, col_idx)
        cell_t.fill = fill_total
        cell_t.border = border_total

    ws_rot.row_dimensions[tot_rot_row].height = 26

    # Anchos de Columna Hoja 4
    col_widths_rot = {
        'A': 16,   # Cód. Base SKU
        'B': 36,   # Descripción Producto
        'C': 18,   # Línea / Formato
        'D': 16,   # Stock Físico (L/Kg)
        'E': 16,   # Envases Est. (u)
        'F': 16,   # (N+2) (L/Kg)
        'G': 16,   # (N+5) (L/Kg)
        'H': 16,   # SlowMovers (L/Kg)
        'I': 16,   # NoMovers (L/Kg)
        'J': 18,   # Envases NoMover (u)
        'K': 24,   # Clasificación Rotación
        'L': 16,   # Precio Ref (€/(L·Kg))
        'M': 22,   # Inmovilizado NoMover (€)
        'N': 22,   # Inmovilizado SlowMover (€)
        'O': 34    # Acción Comercial Recomendada
    }
    for col_l, width in col_widths_rot.items():
        ws_rot.column_dimensions[col_l].width = width

    ws_rot.freeze_panes = "A9"
    ws_rot.auto_filter.ref = f"A8:O{tot_rot_row - 1}"

    # =========================================================================
    # HOJA 5: SUGERENCIAS STOCK (PLAN EJECUTIVO CEO: DIO & LIQUIDACIÓN)
    # =========================================================================
    sug_data = data.get("stock_suggestions", {})
    if sug_data:
        ws_sug = wb.create_sheet(title="Sugerencias Stock")
        ws_sug.views.sheetView[0].showGridLines = True

        sug_kpis = sug_data.get("kpis", {})
        sug_plans = sug_data.get("plans", [])
        sug_items = sug_data.get("items", [])

        font_title = font_main_title
        font_subtitle = font_sub_title
        font_kpi_lbl = font_kpi_label
        font_grp = font_grp_header
        font_th = font_col_header
        fill_header_dark = PatternFill(start_color=COLOR_HEADER_BG, end_color=COLOR_HEADER_BG, fill_type="solid")
        fill_grp_id = PatternFill(start_color=COLOR_GRP_ID, end_color=COLOR_GRP_ID, fill_type="solid")
        fill_grp_demand = PatternFill(start_color=COLOR_GRP_DEMAND, end_color=COLOR_GRP_DEMAND, fill_type="solid")
        fill_grp_plan = PatternFill(start_color=COLOR_GRP_PLAN, end_color=COLOR_GRP_PLAN, fill_type="solid")
        fill_kpi_box = PatternFill(start_color=COLOR_KPI_BG, end_color=COLOR_KPI_BG, fill_type="solid")
        fill_row_alt = fill_alt
        fill_row_white = fill_white
        C_CRIT_FG = COLOR_CRITICAL_FG
        C_WARN_FG = COLOR_WARNING_FG
        C_COV_FG = COLOR_COVERED_FG
        C_DARK_NAVY = "0F172A"

        # 1. Banner Superior Corporativo
        ws_sug.merge_cells("A1:V1")
        c_st = ws_sug["A1"]
        c_st.value = "  CODIAGRO S.L.  |  SUGERENCIAS STOCK: PLAN EJECUTIVO DE OPTIMIZACIÓN DE DIO Y SALIDA DE EXCEDENTES (CEO PLAYBOOK)"
        c_st.font = font_title
        c_st.fill = fill_header_dark
        c_st.alignment = Alignment(vertical="center", horizontal="left")
        ws_sug.row_dimensions[1].height = 28

        ws_sug.merge_cells("A2:V2")
        c_ss = ws_sug["A2"]
        c_ss.value = (
            f"  S&OP Noviembre 2026 | Balance Oficial: DIO Base {sug_kpis.get('dio_base_total_days', 97.9):.1f}d (Inventario Total 1.323.814 € / COGS LTM 4.935.072 €) | "
            f"Ahorro DIO Empresa: -{sug_kpis.get('dio_saving_total_days', 5.77):.2f}d (Post: {sug_kpis.get('dio_post_total_days', 92.1):.1f}d) | "
            f"DSO Global: +{sug_kpis.get('dso_impact_total_days', 0.56):.2f}d (Post: {sug_kpis.get('dso_post_total_days', 104.1):.1f}d | Ratio Arbitraje: {sug_kpis.get('ratio_dio_vs_dso', 10.3):.1f}x) | "
            f"Ciclo Neto CCC: {sug_kpis.get('ccc_net_impact_days', -5.21):.2f}d (Post: {sug_kpis.get('ccc_post_total_days', 131.8):.1f}d) | "
            f"Ganancia Fin. Neta: +{sug_kpis.get('net_financial_gain_total_eur', 4607.76):,.2f} € | "
            f"Palets Liberados: {sug_kpis.get('palets_liberados_total', 201.8):.1f} | Emisión: {now_str}"
        )
        c_ss.font = font_subtitle
        c_ss.fill = fill_header_dark
        c_ss.alignment = Alignment(vertical="center", horizontal="left")
        ws_sug.row_dimensions[2].height = 18
        ws_sug.row_dimensions[3].height = 8

        # 2. Tarjetas KPI Ejecutivas (Filas 4 y 5)
        kpis_sug_cards = [
            ("A4:C4", "A5:C5", "COSTE NOMOVERS (>6M SIN ROTACIÓN)", f"{sug_kpis.get('nomover_cost_eur', 0):,.2f} € ({sug_kpis.get('nomover_skus', 42)} SKUs)", C_CRIT_FG),
            ("D4:F4", "D5:F5", "ARBITRAJE WORKING CAPITAL", f"DIO -{sug_kpis.get('dio_saving_total_days', 5.77):.2f}d vs DSO +{sug_kpis.get('dso_impact_total_days', 0.56):.2f}d ({sug_kpis.get('ratio_dio_vs_dso', 10.3):.1f}x)", C_COV_FG),
            ("G4:I4", "G5:I5", "DESGLOSE COSTES COMERCIALES", f"Dto: -{sug_kpis.get('coste_descuentos_total_eur', 0):,.0f}€ | Merch: -{sug_kpis.get('coste_merchandising_total_eur', 0):,.0f}€ | Fin: -{sug_kpis.get('coste_financiero_aplazamiento_total_eur', 0):,.0f}€", C_WARN_FG),
            ("J4:L4", "J5:L5", "CAJA NETA A RECUPERAR (CASH-IN)", f"{sug_kpis.get('ingreso_neto_total_eur', 0):,.2f} € (NM: {sug_kpis.get('ingreso_nm_eur', 0):,.0f} €)", C_DARK_NAVY),
            ("M4:P4", "M5:P5", "MARGEN NETO REAL POST-COSTES", f"+{sug_kpis.get('margen_neto_total_eur', 0):,.2f} € ({sug_kpis.get('margen_neto_pct', 60.5):.1f}% Real)", C_COV_FG),
            ("Q4:V4", "Q5:V5", "GANANCIA FINANCIERA & ESPACIO", f"+{sug_kpis.get('net_financial_gain_total_eur', 0):,.2f} € Neta (Holding vs Fin) | {sug_kpis.get('palets_liberados_total', 0):.1f} Palets", C_COV_FG),
        ]
        for top_r, bot_r, lbl, val, col_hex in kpis_sug_cards:
            ws_sug.merge_cells(top_r)
            ws_sug.merge_cells(bot_r)
            top_c = top_r.split(":")[0]
            bot_c = bot_r.split(":")[0]
            ws_sug[top_c].value = lbl
            ws_sug[top_c].font = font_kpi_lbl
            ws_sug[top_c].alignment = Alignment(horizontal="center", vertical="center")
            ws_sug[bot_c].value = val
            ws_sug[bot_c].font = Font(name="Calibri", size=10, bold=True, color=col_hex)
            ws_sug[bot_c].alignment = Alignment(horizontal="center", vertical="center")
            s_row, s_col = openpyxl.utils.coordinate_to_tuple(top_c)
            e_row, e_col = openpyxl.utils.coordinate_to_tuple(bot_r.split(":")[1])
            for r_i in range(s_row, e_row + 1):
                for c_i in range(s_col, e_col + 1):
                    cell = ws_sug.cell(row=r_i, column=c_i)
                    cell.fill = fill_kpi_box
                    cell.border = border_kpi

        ws_sug.row_dimensions[4].height = 16
        ws_sug.row_dimensions[5].height = 24
        ws_sug.row_dimensions[6].height = 10

        # 3. TABLA I: RESUMEN EJECUTIVO DE LOS 5 PLANES ESTRATÉGICOS (PLAYBOOKS CEO)
        ws_sug.merge_cells("A7:V7")
        ws_sug["A7"].value = "  I. RESUMEN EJECUTIVO DE LOS 5 PLANES DE CHOQUE (PLAYBOOKS COMERCIALES, FINANCIEROS Y TÉCNICOS)"
        ws_sug["A7"].font = font_grp
        ws_sug["A7"].fill = fill_grp_id
        ws_sug["A7"].alignment = Alignment(horizontal="left", vertical="center")
        ws_sug.row_dimensions[7].height = 22

        plan_headers = [
            (1, "Cód. Plan", fill_grp_id),
            (2, "Playbook Estratégico / Línea de Acción", fill_grp_id),
            (3, "Target / Enfoque", fill_grp_id),
            (4, "Nº SKUs", fill_grp_demand),
            (5, "SKUs NM", fill_grp_demand),
            (6, "Vol. Total (L/Kg)", fill_grp_demand),
            (7, "Coste Total (€)", fill_grp_demand),
            (8, "Ahorro DIO (d)", fill_grp_plan),
            (9, "DSO Global (d)", fill_grp_plan),
            (10, "Ratio DIO/DSO", fill_grp_plan),
            (11, "Palets Lib.", fill_grp_plan),
            (12, "Ingreso Neto (€)", fill_grp_plan),
            (13, "Coste Dto (€)", fill_grp_plan),
            (14, "Coste Merch (€)", fill_grp_plan),
            (15, "Coste Fin (€)", fill_grp_plan),
            (16, "Margen Real (€)", fill_grp_plan),
            (17, "Margen Real (%)", fill_grp_plan),
            (18, "Plazo Pago", fill_grp_id),
            (19, "Palanca Precio", fill_grp_id),
        ]
        for c_i, h_t, fl in plan_headers:
            c = ws_sug.cell(row=8, column=c_i, value=h_t)
            c.font = font_th
            c.fill = fl
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = border_thin
        ws_sug.merge_cells("T8:V8")
        ws_sug["T8"].value = "Incentivos No-Precio (Merchandising, Muestras Gratis, Rappel) & Racional"
        ws_sug["T8"].font = font_th
        ws_sug["T8"].fill = fill_grp_id
        ws_sug["T8"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        for c_i in range(20, 23):
            ws_sug.cell(row=8, column=c_i).border = border_thin
            ws_sug.cell(row=8, column=c_i).fill = fill_grp_id
        ws_sug.row_dimensions[8].height = 24

        cur_p_row = 9
        start_p_row = cur_p_row
        for p in sug_plans:
            fill_r = fill_row_alt if (cur_p_row % 2 == 0) else fill_row_white
            ws_sug.cell(cur_p_row, 1, p.get("plan_id", "")).alignment = Alignment(horizontal="center", vertical="center")
            ws_sug.cell(cur_p_row, 1).font = font_data_bold
            ws_sug.cell(cur_p_row, 2, p.get("title", "")).alignment = Alignment(horizontal="left", vertical="center")
            ws_sug.cell(cur_p_row, 2).font = font_data_bold
            ws_sug.cell(cur_p_row, 3, p.get("target_segment", "")).alignment = Alignment(horizontal="left", vertical="center")
            ws_sug.cell(cur_p_row, 3).font = font_data

            c4 = ws_sug.cell(cur_p_row, 4, p.get("skus_count", 0))
            c4.number_format = '#,##0'
            c4.alignment = Alignment(horizontal="center", vertical="center")
            c4.font = font_data_bold

            c5 = ws_sug.cell(cur_p_row, 5, p.get("nomover_skus_count", 0))
            c5.number_format = '#,##0'
            c5.alignment = Alignment(horizontal="center", vertical="center")
            c5.font = font_crit

            c6 = ws_sug.cell(cur_p_row, 6, p.get("total_qty", 0))
            c6.number_format = '#,##0'
            c6.alignment = Alignment(horizontal="right", vertical="center")
            c6.font = font_data_bold

            c7 = ws_sug.cell(cur_p_row, 7, p.get("total_cost_eur", 0))
            c7.number_format = '#,##0.00 €'
            c7.alignment = Alignment(horizontal="right", vertical="center")
            c7.font = font_data_bold

            c8 = ws_sug.cell(cur_p_row, 8, -abs(p.get("dio_saving_days", 0)))
            c8.number_format = '+0.00;-0.00;0.00'
            c8.alignment = Alignment(horizontal="right", vertical="center")
            c8.font = font_cov

            c9 = ws_sug.cell(cur_p_row, 9, p.get("dso_impact_days", 0))
            c9.number_format = '+0.00;-0.00;0.00'
            c9.alignment = Alignment(horizontal="right", vertical="center")
            c9.font = font_warn

            c10 = ws_sug.cell(cur_p_row, 10, p.get("ratio_dio_vs_dso", 0))
            c10.number_format = '0.0"x"'
            c10.alignment = Alignment(horizontal="right", vertical="center")
            c10.font = font_cov

            c11 = ws_sug.cell(cur_p_row, 11, p.get("palets_liberados", 0))
            c11.number_format = '#,##0.0'
            c11.alignment = Alignment(horizontal="right", vertical="center")
            c11.font = font_data_bold

            c12 = ws_sug.cell(cur_p_row, 12, p.get("ingreso_neto_eur", 0))
            c12.number_format = '#,##0.00 €'
            c12.alignment = Alignment(horizontal="right", vertical="center")
            c12.font = font_data_bold

            c13 = ws_sug.cell(cur_p_row, 13, -abs(p.get("coste_dto_eur", 0)))
            c13.number_format = '#,##0.00 €'
            c13.alignment = Alignment(horizontal="right", vertical="center")
            c13.font = font_warn

            c14 = ws_sug.cell(cur_p_row, 14, -abs(p.get("coste_merch_eur", 0)))
            c14.number_format = '#,##0.00 €'
            c14.alignment = Alignment(horizontal="right", vertical="center")
            c14.font = font_warn

            c15 = ws_sug.cell(cur_p_row, 15, -abs(p.get("coste_fin_pago_eur", 0)))
            c15.number_format = '#,##0.00 €'
            c15.alignment = Alignment(horizontal="right", vertical="center")
            c15.font = font_warn

            c16 = ws_sug.cell(cur_p_row, 16, p.get("margen_neto_eur", 0))
            c16.number_format = '#,##0.00 €'
            c16.alignment = Alignment(horizontal="right", vertical="center")
            c16.font = font_cov

            c17 = ws_sug.cell(cur_p_row, 17, p.get("margen_neto_pct", 0) / 100.0)
            c17.number_format = '0.0%'
            c17.alignment = Alignment(horizontal="right", vertical="center")
            c17.font = font_cov

            c18 = ws_sug.cell(cur_p_row, 18, p.get("payment_terms", ""))
            c18.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            c18.font = font_data

            c19 = ws_sug.cell(cur_p_row, 19, p.get("price_lever", ""))
            c19.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            c19.font = font_data

            ws_sug.merge_cells(f"T{cur_p_row}:V{cur_p_row}")
            c20 = ws_sug.cell(cur_p_row, 20, f"{p.get('non_price_incentives', '')} — {p.get('rationale', '')}")
            c20.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            c20.font = font_data

            for c_i in range(1, 23):
                cell_p = ws_sug.cell(cur_p_row, c_i)
                cell_p.fill = fill_r
                cell_p.border = border_thin

            ws_sug.row_dimensions[cur_p_row].height = 28
            cur_p_row += 1

        # Fila Totales Tabla I
        tot_p_row = cur_p_row
        ws_sug.merge_cells(f"A{tot_p_row}:C{tot_p_row}")
        ws_sug[f"A{tot_p_row}"].value = "TOTAL CONSOLIDADO 5 PLAYBOOKS ESTRATÉGICOS"
        ws_sug[f"A{tot_p_row}"].font = font_total
        ws_sug[f"A{tot_p_row}"].alignment = Alignment(horizontal="right", vertical="center")

        sum_p_cols = [
            (4, "D", '#,##0'), (5, "E", '#,##0'), (6, "F", '#,##0'), (7, "G", '#,##0.00 €'),
            (8, "H", '+0.00;-0.00;0.00'), (9, "I", '+0.00;-0.00;0.00'), (11, "K", '#,##0.0'),
            (12, "L", '#,##0.00 €'), (13, "M", '#,##0.00 €'), (14, "N", '#,##0.00 €'), (15, "O", '#,##0.00 €'),
            (16, "P", '#,##0.00 €')
        ]
        for c_idx, letter, n_fmt in sum_p_cols:
            c_s = ws_sug.cell(tot_p_row, c_idx, f"=SUM({letter}{start_p_row}:{letter}{tot_p_row - 1})")
            c_s.number_format = n_fmt
            c_s.font = font_total
            c_s.alignment = Alignment(horizontal="right", vertical="center")

        # Ratio total
        c_rat_tot = ws_sug.cell(tot_p_row, 10, f"=IF(I{tot_p_row}>0, ABS(H{tot_p_row})/I{tot_p_row}, 0)")
        c_rat_tot.number_format = '0.0"x"'
        c_rat_tot.font = font_total
        c_rat_tot.alignment = Alignment(horizontal="right", vertical="center")

        c_mb_tot = ws_sug.cell(tot_p_row, 17, f"=IF(L{tot_p_row}>0, P{tot_p_row}/L{tot_p_row}, 0)")
        c_mb_tot.number_format = '0.0%'
        c_mb_tot.font = font_total
        c_mb_tot.alignment = Alignment(horizontal="right", vertical="center")

        for c_i in range(1, 23):
            c_t = ws_sug.cell(tot_p_row, c_i)
            c_t.fill = fill_total
            c_t.border = border_total
        ws_sug.row_dimensions[tot_p_row].height = 25

        # 4. TABLA II: CATÁLOGO OPERATIVO DETALLADO SKU POR SKU (89 SKUs CON CLIENTES DIANA Y PROPUESTA COMERCIAL)
        sep_row = tot_p_row + 2
        ws_sug.merge_cells(f"A{sep_row}:V{sep_row}")
        ws_sug[f"A{sep_row}"].value = "  II. MATRIZ OPERATIVA SKU POR SKU: CLIENTES DIANA (CON HISTÓRICO), PRECIOS OBJETIVO, PLAZOS DE PAGO Y OFERTAS COMPLEMENTARIAS"
        ws_sug[f"A{sep_row}"].font = font_grp
        ws_sug[f"A{sep_row}"].fill = fill_grp_plan
        ws_sug[f"A{sep_row}"].alignment = Alignment(horizontal="left", vertical="center")
        ws_sug.row_dimensions[sep_row].height = 22

        hdr_sku_row = sep_row + 1
        sku_headers = [
            (1, "Prioridad", fill_grp_id),
            (2, "Plan Asignado", fill_grp_id),
            (3, "Cód. Base SKU", fill_grp_id),
            (4, "Producto / Descripción", fill_grp_id),
            (5, "Línea / Formato", fill_grp_id),
            (6, "Segmento Rotación", fill_grp_id),
            (7, "Stock Actual (L/Kg)", fill_grp_demand),
            (8, "Exceso NoMover (L/Kg)", fill_grp_demand),
            (9, "Exceso SlowMover (L/Kg)", fill_grp_demand),
            (10, "Total Sobrante (L/Kg)", fill_grp_demand),
            (11, "Coste Unit. (€)", fill_grp_demand),
            (12, "Inmov. NoMover (€)", fill_grp_demand),
            (13, "Coste Total Sobrante (€)", fill_grp_demand),
            (14, "Ahorro DIO (días)", fill_grp_plan),
            (15, "Palets Liberados", fill_grp_plan),
            (16, "PVP Histórico (€)", fill_grp_plan),
            (17, "PVP Oferta Sugerido (€)", fill_grp_plan),
            (18, "Palanca Precio / Dto.", fill_grp_plan),
            (19, "Ingreso Neto Est. (€)", fill_grp_plan),
            (20, "Margen Neto Est. (€)", fill_grp_plan),
            (21, "Clientes Diana Recomendados (Histórico Real y Afinidad)", fill_grp_id),
            (22, "Facilidades de Pago + Incentivos (Merchandising / Muestras / Rappel)", fill_grp_id),
        ]
        for c_i, h_t, fl in sku_headers:
            c = ws_sug.cell(row=hdr_sku_row, column=c_i, value=h_t)
            c.font = font_th
            c.fill = fl
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = border_thin
        ws_sug.row_dimensions[hdr_sku_row].height = 26

        cur_s_row = hdr_sku_row + 1
        start_s_row = cur_s_row
        for it in sug_items:
            fill_r = fill_row_alt if (cur_s_row % 2 == 0) else fill_row_white
            is_nm = bool(it.get("is_nomover", False))
            prio = "P1 - Crítica (>1.500€ NM)" if (is_nm and it.get("no_mover_eur", 0) >= 1500) else ("P1 - NoMover (>6M)" if is_nm else "P2 - SlowMover (3-6M)")
            c_p = ws_sug.cell(cur_s_row, 1, prio)
            c_p.alignment = Alignment(horizontal="center", vertical="center")
            if is_nm:
                c_p.fill = fill_crit
                c_p.font = font_crit
            else:
                c_p.fill = fill_warn
                c_p.font = font_warn

            c_pl = ws_sug.cell(cur_s_row, 2, it.get("plan_name", ""))
            c_pl.font = font_data_bold
            c_pl.alignment = Alignment(horizontal="left", vertical="center")

            c_sk = ws_sug.cell(cur_s_row, 3, it.get("base_sku", ""))
            c_sk.font = font_data_bold
            c_sk.alignment = Alignment(horizontal="center", vertical="center")

            c_de = ws_sug.cell(cur_s_row, 4, it.get("desc", ""))
            c_de.font = font_data
            c_de.alignment = Alignment(horizontal="left", vertical="center")

            c_fm = ws_sug.cell(cur_s_row, 5, it.get("formato", ""))
            c_fm.font = font_data
            c_fm.alignment = Alignment(horizontal="center", vertical="center")

            c_sg = ws_sug.cell(cur_s_row, 6, it.get("segmento", ""))
            c_sg.alignment = Alignment(horizontal="center", vertical="center")
            if is_nm:
                c_sg.fill = fill_crit
                c_sg.font = font_crit
            else:
                c_sg.fill = fill_warn
                c_sg.font = font_warn

            c7 = ws_sug.cell(cur_s_row, 7, it.get("stock_actual", 0))
            c7.number_format = '#,##0'
            c7.font = font_data
            c7.alignment = Alignment(horizontal="right", vertical="center")

            c8 = ws_sug.cell(cur_s_row, 8, it.get("no_mover_qty", 0))
            c8.number_format = '#,##0'
            c8.font = font_crit if is_nm else font_data
            c8.alignment = Alignment(horizontal="right", vertical="center")

            c9 = ws_sug.cell(cur_s_row, 9, it.get("slow_mover_qty", 0))
            c9.number_format = '#,##0'
            c9.font = font_warn if it.get("slow_mover_qty", 0) > 0 else font_data
            c9.alignment = Alignment(horizontal="right", vertical="center")

            c10 = ws_sug.cell(cur_s_row, 10, it.get("exceso_total_qty", 0))
            c10.number_format = '#,##0'
            c10.font = font_data_bold
            c10.alignment = Alignment(horizontal="right", vertical="center")

            c11 = ws_sug.cell(cur_s_row, 11, it.get("coste_unit", 0))
            c11.number_format = '#,##0.00 €'
            c11.font = font_data
            c11.alignment = Alignment(horizontal="right", vertical="center")

            c12 = ws_sug.cell(cur_s_row, 12, it.get("no_mover_eur", 0))
            c12.number_format = '#,##0.00 €'
            c12.font = font_crit if is_nm else font_data
            c12.alignment = Alignment(horizontal="right", vertical="center")

            c13 = ws_sug.cell(cur_s_row, 13, it.get("exceso_total_eur", 0))
            c13.number_format = '#,##0.00 €'
            c13.font = font_data_bold
            c13.alignment = Alignment(horizontal="right", vertical="center")

            c14 = ws_sug.cell(cur_s_row, 14, -abs(it.get("dio_saving_days", 0)))
            c14.number_format = '+0.00;-0.00;0.00'
            c14.font = font_cov
            c14.alignment = Alignment(horizontal="right", vertical="center")

            c15 = ws_sug.cell(cur_s_row, 15, it.get("palets_liberados", 0))
            c15.number_format = '#,##0.0'
            c15.font = font_data
            c15.alignment = Alignment(horizontal="right", vertical="center")

            c16 = ws_sug.cell(cur_s_row, 16, it.get("pvp_historico", 0))
            c16.number_format = '#,##0.00 €'
            c16.font = font_data
            c16.alignment = Alignment(horizontal="right", vertical="center")

            c17 = ws_sug.cell(cur_s_row, 17, it.get("pvp_objetivo", 0))
            c17.number_format = '#,##0.00 €'
            c17.font = font_data_bold
            c17.alignment = Alignment(horizontal="right", vertical="center")

            c18 = ws_sug.cell(cur_s_row, 18, f"{it.get('dto_label', '')} (Dto: -{it.get('coste_dto_eur', 0):,.1f} €)")
            c18.font = font_warn
            c18.alignment = Alignment(horizontal="left", vertical="center")

            c19 = ws_sug.cell(cur_s_row, 19, it.get("ingreso_neto_eur", 0))
            c19.number_format = '#,##0.00 €'
            c19.font = font_data_bold
            c19.alignment = Alignment(horizontal="right", vertical="center")

            c20 = ws_sug.cell(cur_s_row, 20, it.get("margen_neto_eur", 0))
            c20.number_format = '#,##0.00 €'
            c20.font = font_cov
            c20.alignment = Alignment(horizontal="right", vertical="center")

            c21 = ws_sug.cell(cur_s_row, 21, f"{it.get('clientes_diana', '')} — [{it.get('motivo_coherente', '')}]")
            c21.font = font_data
            c21.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            c22 = ws_sug.cell(cur_s_row, 22, f"Pago: {it.get('condiciones_pago', '')} (+{it.get('dias_aplazamiento_extra', 0)}d | Fin: -{it.get('coste_fin_pago_eur', 0):,.1f}€) | Merch: -{it.get('coste_merch_eur', 0):,.1f}€ | Incentivos: {it.get('incentivos_adicionales', '')}")
            c22.font = font_data
            c22.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            for c_i in range(2, 6):
                ws_sug.cell(cur_s_row, c_i).fill = fill_r
            for c_i in range(7, 23):
                ws_sug.cell(cur_s_row, c_i).fill = fill_r
            for c_i in range(1, 23):
                ws_sug.cell(cur_s_row, c_i).border = border_thin

            ws_sug.row_dimensions[cur_s_row].height = 28
            cur_s_row += 1

        # Fila Totales Tabla II
        tot_s_row = cur_s_row
        ws_sug.merge_cells(f"A{tot_s_row}:F{tot_s_row}")
        ws_sug[f"A{tot_s_row}"].value = f"TOTAL CATÁLOGO ACCIONABLE ({len(sug_items)} SKUs CON EXCESO DE STOCK)"
        ws_sug[f"A{tot_s_row}"].font = font_total
        ws_sug[f"A{tot_s_row}"].alignment = Alignment(horizontal="right", vertical="center")

        sum_s_cols = [
            (7, "G", '#,##0'), (8, "H", '#,##0'), (9, "I", '#,##0'), (10, "J", '#,##0'),
            (12, "L", '#,##0.00 €'), (13, "M", '#,##0.00 €'),
            (14, "N", '+0.00;-0.00;0.00'), (15, "O", '#,##0.0'),
            (19, "S", '#,##0.00 €'), (20, "T", '#,##0.00 €')
        ]
        for c_idx, letter, n_fmt in sum_s_cols:
            c_s = ws_sug.cell(tot_s_row, c_idx, f"=SUM({letter}{start_s_row}:{letter}{tot_s_row - 1})")
            c_s.number_format = n_fmt
            c_s.font = font_total
            c_s.alignment = Alignment(horizontal="right", vertical="center")

        for c_i in range(1, 23):
            c_t = ws_sug.cell(tot_s_row, c_i)
            c_t.fill = fill_total
            c_t.border = border_total
        ws_sug.row_dimensions[tot_s_row].height = 26

        col_widths_sug = {
            'A': 16, 'B': 30, 'C': 16, 'D': 32, 'E': 16, 'F': 22,
            'G': 16, 'H': 18, 'I': 18, 'J': 18, 'K': 14, 'L': 18,
            'M': 20, 'N': 16, 'O': 15, 'P': 15, 'Q': 16, 'R': 28,
            'S': 18, 'T': 18, 'U': 52, 'V': 58
        }
        for col_l, w in col_widths_sug.items():
            ws_sug.column_dimensions[col_l].width = w

        ws_sug.freeze_panes = f"A{start_s_row}"
        ws_sug.auto_filter.ref = f"A{hdr_sku_row}:V{tot_s_row - 1}"

    # Guardar Excel en web_dashboard, web_dashboard_direccion y en ruta destino con control de bloqueo
    dashboard_excel_path = get_dest(os.path.join("web_dashboard", os.path.basename(output_path)))
    try:
        wb.save(dashboard_excel_path)
        print(f"Archivo guardado exitosamente en: {dashboard_excel_path}")
    except Exception as e:
        print(f"Aviso guardando en dashboard: {e}")

    dir_dashboard_excel_path = get_dest(os.path.join("web_dashboard_direccion", os.path.basename(output_path)))
    try:
        if os.path.exists("web_dashboard_direccion"):
            wb.save(dir_dashboard_excel_path)
            print(f"Archivo guardado exitosamente en: {dir_dashboard_excel_path}")
    except Exception as e:
        print(f"Aviso guardando en dashboard direccion: {e}")

    output_path = get_dest(output_path)
    saved_path = output_path
    try:
        wb.save(output_path)
        print(f"¡Orden de Fabricación MRP, Rotación y Sugerencias Stock generada exitosamente con 5 hojas en: {output_path}!")
    except PermissionError:
        alt_path = output_path.replace(".xlsx", "_Actualizado.xlsx")
        wb.save(alt_path)
        saved_path = alt_path
        print(f"¡ATENCIÓN: {output_path} estaba abierto en Excel! Se ha guardado la versión actualizada con 5 hojas como: {alt_path} y en {dashboard_excel_path}")
    return saved_path

if __name__ == "__main__":
    generar_excel_orden_fabricacion()
