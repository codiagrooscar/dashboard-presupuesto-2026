import sys
import os
# pyrefly: ignore[missing-import]
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def create_presentation():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # Colors
    NAVY = RGBColor(11, 37, 69)         # #0B2545 Primary Dark
    NAVY_LIGHT = RGBColor(31, 78, 121)   # #1F4E79 Secondary
    GREEN = RGBColor(22, 101, 52)        # #166534 Accent Green
    GREEN_ACCENT = RGBColor(16, 185, 129)# #10B981 Vibrant Green
    CARD_BG = RGBColor(248, 250, 252)    # #F8FAFC
    CARD_BORDER = RGBColor(226, 232, 240)# #E2E8F0
    WHITE = RGBColor(255, 255, 255)
    GRAY_TEXT = RGBColor(100, 116, 139)  # #64748B
    DARK_TEXT = RGBColor(15, 23, 42)     # #0F172A
    ALERT_RED = RGBColor(185, 28, 28)
    GOLD = RGBColor(180, 83, 9)

    logo_path = 'web_dashboard/codiagro_logo.png'
    has_logo = os.path.exists(logo_path)

    def add_header(slide, title, category="CODIAGRO · CONTROL PRESUPUESTARIO 2027"):
        # Top banner background
        top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.15))
        top_bar.fill.solid()
        top_bar.fill.fore_color.rgb = NAVY
        top_bar.line.fill.background()

        # Accent line under top bar
        line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(1.12), Inches(13.333), Inches(0.04))
        line.fill.solid()
        line.fill.fore_color.rgb = GREEN_ACCENT
        line.line.fill.background()

        # Category / breadcrumb
        tb_cat = slide.shapes.add_textbox(Inches(0.8), Inches(0.15), Inches(9.0), Inches(0.3))
        p_cat = tb_cat.text_frame.paragraphs[0]
        p_cat.text = category.upper()
        p_cat.font.size = Pt(9.5)
        p_cat.font.bold = True
        p_cat.font.color.rgb = GREEN_ACCENT
        p_cat.font.name = "Calibri"

        # Main Title
        tb_title = slide.shapes.add_textbox(Inches(0.8), Inches(0.42), Inches(9.5), Inches(0.6))
        p_title = tb_title.text_frame.paragraphs[0]
        p_title.text = title
        p_title.font.size = Pt(20)
        p_title.font.bold = True
        p_title.font.color.rgb = WHITE
        p_title.font.name = "Calibri"

        # Logo on right
        if has_logo:
            slide.shapes.add_picture(logo_path, Inches(10.8), Inches(0.22), height=Inches(0.65))

        # Bottom footer bar
        footer = slide.shapes.add_textbox(Inches(0.8), Inches(7.1), Inches(11.733), Inches(0.3))
        p_f = footer.text_frame.paragraphs[0]
        p_f.text = "Codiagro S.L. | Informe Confidencial para Dirección General · Presupuesto 2027"
        p_f.font.size = Pt(9)
        p_f.font.color.rgb = GRAY_TEXT
        p_f.font.name = "Calibri"

    def add_card(slide, left, top, width, height, bg_color=CARD_BG, border_color=CARD_BORDER):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = bg_color
        card.line.color.rgb = border_color
        card.line.width = Pt(1)
        return card

    # ==========================================
    # SLIDE 1: PORTADA
    # ==========================================
    s1 = prs.slides.add_slide(blank_layout)
    bg1 = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = NAVY
    bg1.line.fill.background()

    # Accent decorative bar
    dec = s1.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(1.6), Inches(0.12), Inches(4.3))
    dec.fill.solid()
    dec.fill.fore_color.rgb = GREEN_ACCENT
    dec.line.fill.background()

    # Logo
    if has_logo:
        s1.shapes.add_picture(logo_path, Inches(1.5), Inches(1.6), height=Inches(0.9))

    tb_p = s1.shapes.add_textbox(Inches(1.5), Inches(2.7), Inches(10.5), Inches(3.2))
    tf_p = tb_p.text_frame
    tf_p.word_wrap = True

    p0 = tf_p.paragraphs[0]
    p0.text = "PLAN ESTRATÉGICO & CONTROL DE GESTIÓN"
    p0.font.size = Pt(12)
    p0.font.bold = True
    p0.font.color.rgb = GREEN_ACCENT
    p0.font.name = "Calibri"
    p0.space_after = Pt(10)

    p1 = tf_p.add_paragraph()
    p1.text = "PRESUPUESTO 2027 &\nCIERRE FORECAST 2026"
    p1.font.size = Pt(36)
    p1.font.bold = True
    p1.font.color.rgb = WHITE
    p1.font.name = "Calibri"
    p1.space_after = Pt(14)

    p2 = tf_p.add_paragraph()
    p2.text = "Visión Integral: Cifra de Negocios, Rentabilidad Operativa (EBITDA), Estructura de Gastos y Desempeño Comercial"
    p2.font.size = Pt(14)
    p2.font.color.rgb = RGBColor(203, 213, 225)
    p2.font.name = "Calibri"
    p2.space_after = Pt(28)

    p3 = tf_p.add_paragraph()
    p3.text = "CODIAGRO AGROQUÍMICA S.L.  |  Presentación para Dirección General  |  Octubre 2026"
    p3.font.size = Pt(11)
    p3.font.color.rgb = RGBColor(148, 163, 184)
    p3.font.name = "Calibri"

    # ==========================================
    # SLIDE 2: RESUMEN EJECUTIVO (LOS 5 PILARES)
    # ==========================================
    s2 = prs.slides.add_slide(blank_layout)
    add_header(s2, "Resumen Ejecutivo: Los 5 Pilares del Presupuesto 2027")

    pillars = [
        ("1. Crecimiento Sostenido", "17,18 M€", "+13,2% (+2,00 M€)", "Incremento del volumen comercial (+8,1% en unidades) y consolidación de precios medios y mix por productos de alta tecnología nutricional.", NAVY_LIGHT),
        ("2. EBITDA Récord", "7,30 M€", "42,5% Margen", "El beneficio operativo crece +874 k€ (+13,6% vs 2026), preservando un margen de rentabilidad récord del 42,5% sobre ingresos.", GREEN),
        ("3. Margen Bruto Fuerte", "65,5%", "11,26 M€ Bruto", "Absorción íntegra de la inflación de materias primas (+4,5%) gracias a economías de escala fabriles y optimización en compras.", NAVY_LIGHT),
        ("4. Eficiencia de OPEX", "13,0%", "Coste Personal/Vtas", "Gastos de estructura bajo control. La masa salarial (2,23 M€) reduce su peso relativo en ventas del 13,2% (2026) al 13,0% (2027).", NAVY_LIGHT),
        ("5. Tracción Inmediata", "129,8%", "Cierre Sep-26", "La demanda actual valida la ambición del plan: Septiembre 2026 cierra con 236.105 uds pedidas vs 181.966 presupuestadas (+29,8%).", GREEN),
    ]

    card_w = Inches(2.22)
    card_h = Inches(5.4)
    start_x = Inches(0.8)
    top_y = Inches(1.45)
    gap_x = Inches(0.18)

    for i, (title, big_val, sub_val, desc, accent_col) in enumerate(pillars):
        x = start_x + i * (card_w + gap_x)
        add_card(s2, x, top_y, card_w, card_h)

        # Header bar in card
        hbar = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, top_y, card_w, Inches(0.45))
        hbar.fill.solid()
        hbar.fill.fore_color.rgb = accent_col
        hbar.line.fill.background()

        tb_h = s2.shapes.add_textbox(x, top_y, card_w, Inches(0.45))
        p = tb_h.text_frame.paragraphs[0]
        p.text = title
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = WHITE
        p.alignment = PP_ALIGN.CENTER

        # Big Value
        tb_v = s2.shapes.add_textbox(x + Inches(0.1), top_y + Inches(0.65), card_w - Inches(0.2), Inches(0.8))
        pv = tb_v.text_frame.paragraphs[0]
        pv.text = big_val
        pv.font.size = Pt(24)
        pv.font.bold = True
        pv.font.color.rgb = NAVY
        pv.alignment = PP_ALIGN.CENTER

        # Sub Value
        tb_sub = s2.shapes.add_textbox(x + Inches(0.1), top_y + Inches(1.45), card_w - Inches(0.2), Inches(0.4))
        ps = tb_sub.text_frame.paragraphs[0]
        ps.text = sub_val
        ps.font.size = Pt(11)
        ps.font.bold = True
        ps.font.color.rgb = GREEN_ACCENT if accent_col == GREEN else NAVY_LIGHT
        ps.alignment = PP_ALIGN.CENTER

        # Description
        tb_d = s2.shapes.add_textbox(x + Inches(0.15), top_y + Inches(2.0), card_w - Inches(0.3), Inches(3.2))
        tb_d.text_frame.word_wrap = True
        pd = tb_d.text_frame.paragraphs[0]
        pd.text = desc
        pd.font.size = Pt(10.5)
        pd.font.color.rgb = DARK_TEXT

    # ==========================================
    # SLIDE 3: CUADRO GENERAL P&L (2026 FC vs 2027 BUDGET)
    # ==========================================
    s3 = prs.slides.add_slide(blank_layout)
    add_header(s3, "Cuenta de Resultados (P&L): Cierre 2026 Forecast vs Presupuesto 2027")

    # Table
    table_shape = s3.shapes.add_table(11, 6, Inches(0.8), Inches(1.45), Inches(8.3), Inches(5.3))
    table = table_shape.table
    table.columns[0].width = Inches(2.8)
    table.columns[1].width = Inches(1.1)
    table.columns[2].width = Inches(1.1)
    table.columns[3].width = Inches(1.1)
    table.columns[4].width = Inches(1.1)
    table.columns[5].width = Inches(1.1)

    pnl_data = [
        ("Línea P&L (en miles de €)", "2026 FC", "% Vtas", "2027 B", "% Vtas", "Var (€ / %)"),
        ("1. Cifra Neta de Negocios (Ventas)", "15.180", "100,0%", "17.180", "100,0%", "+2.000 (+13,2%)"),
        ("2. Coste de las Ventas (COGS)", "-5.234", "34,5%", "-5.919", "34,5%", "-686 (+13,1%)"),
        ("3. Margen Bruto (Gross Profit)", "9.946", "65,5%", "11.261", "65,5%", "+1.315 (+13,2%)"),
        ("4. Gastos de Personal", "-2.007", "13,2%", "-2.229", "13,0%", "-222 (+11,1%)"),
        ("5. Otros Ingresos Operativos (Export Transp)", "+169", "1,1%", "+305", "1,8%", "+137 (+81,1%)"),
        ("6. Otros Gastos Operativos (Opex)", "-1.684", "11,1%", "-2.038", "11,9%", "-355 (+21,0%)"),
        ("7. Total Gastos Fijos (OPEX Neto)", "-3.522", "23,2%", "-3.962", "23,1%", "-440 (+12,5%)"),
        ("8. EBITDA Ajustado", "6.424", "42,3%", "7.299", "42,5%", "+874 (+13,6%)"),
        ("9. Amortizaciones y Depreciaciones (D&A)", "-230", "1,5%", "-231", "1,3%", "-1 (+0,4%)"),
        ("10. EBIT Ajustado (Resultado Operativo)", "6.195", "40,8%", "7.068", "41,1%", "+873 (+14,1%)"),
    ]

    for r_idx, row in enumerate(pnl_data):
        for c_idx, val in enumerate(row):
            cell = table.cell(r_idx, c_idx)
            cell.text = val
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.font.name = "Calibri"
            p.font.size = Pt(10) if r_idx > 0 else Pt(10.5)

            if r_idx == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = NAVY
                p.font.bold = True
                p.font.color.rgb = WHITE
                p.alignment = PP_ALIGN.CENTER if c_idx > 0 else PP_ALIGN.LEFT
            elif r_idx in (1, 3, 8, 10): # Key lines
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(220, 238, 225) if r_idx in (3, 8, 10) else RGBColor(241, 245, 249)
                p.font.bold = True
                p.font.color.rgb = NAVY
                if c_idx > 0: p.alignment = PP_ALIGN.RIGHT
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = WHITE if r_idx % 2 == 1 else RGBColor(248, 250, 252)
                p.font.color.rgb = DARK_TEXT
                if c_idx > 0: p.alignment = PP_ALIGN.RIGHT

    # Side KPI cards on right
    side_w = Inches(3.2)
    add_card(s3, Inches(9.3), Inches(1.45), side_w, Inches(2.55))
    tb_s1 = s3.shapes.add_textbox(Inches(9.5), Inches(1.6), side_w - Inches(0.4), Inches(2.2))
    tf1 = tb_s1.text_frame
    tf1.word_wrap = True
    p = tf1.paragraphs[0]
    p.text = "RENTABILIDAD OPERATIVA"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = GREEN_ACCENT
    p1 = tf1.add_paragraph()
    p1.text = "7,30 M€"
    p1.font.size = Pt(28)
    p1.font.bold = True
    p1.font.color.rgb = NAVY
    p2 = tf1.add_paragraph()
    p2.text = "EBITDA récord en la historia de la compañía (+13,6% vs 2026), preservando un margen excelente del 42,5% a pesar de la inflación de suministros y refuerzos operativos."
    p2.font.size = Pt(10)
    p2.font.color.rgb = DARK_TEXT

    add_card(s3, Inches(9.3), Inches(4.2), side_w, Inches(2.55))
    tb_s2 = s3.shapes.add_textbox(Inches(9.5), Inches(4.35), side_w - Inches(0.4), Inches(2.2))
    tf2 = tb_s2.text_frame
    tf2.word_wrap = True
    p = tf2.paragraphs[0]
    p.text = "APALANCAMIENTO OPERATIVO"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = NAVY_LIGHT
    p1 = tf2.add_paragraph()
    p1.text = "+2,00 M€"
    p1.font.size = Pt(28)
    p1.font.bold = True
    p1.font.color.rgb = NAVY
    p2 = tf2.add_paragraph()
    p2.text = "Crecimiento en ingresos (+13,2%) sostenido en España y mercados clave de exportación, convirtiendo cada euro extra de venta en 44 céntimos de EBITDA adicional."
    p2.font.size = Pt(10)
    p2.font.color.rgb = DARK_TEXT

    # ==========================================
    # SLIDE 4: VENTAS POR COMERCIAL (2026 vs 2027)
    # ==========================================
    s4 = prs.slides.add_slide(blank_layout)
    add_header(s4, "Desglose Comercial: Objetivos de Ventas 2027 por Zona y Gestor")

    t4_shape = s4.shapes.add_table(9, 7, Inches(0.8), Inches(1.45), Inches(8.5), Inches(5.3))
    t4 = t4_shape.table
    t4.columns[0].width = Inches(1.7)
    t4.columns[1].width = Inches(1.2)
    t4.columns[2].width = Inches(1.2)
    t4.columns[3].width = Inches(1.0)
    t4.columns[4].width = Inches(1.1)
    t4.columns[5].width = Inches(1.1)
    t4.columns[6].width = Inches(1.2)

    com_data = [
        ("Comercial", "2026 (€)", "2027 (€)", "Var (€)", "Var (Uds)", "Cuota '27", "Mercado Clave"),
        ("García", "5.758.078", "6.425.613", "+11,6%", "+7,6%", "37,4%", "España Centro-Sur / Clientes A"),
        ("Ricardo", "3.824.612", "4.281.919", "+12,0%", "+6,6%", "24,9%", "Levante / Distribución"),
        ("Mehmet", "1.778.148", "2.027.764", "+14,0%", "+7,5%", "11,8%", "Exportación Internacional"),
        ("Pedro", "1.239.220", "1.413.541", "+14,1%", "+9,3%", "8,2%", "Zona Norte / Agroindustria"),
        ("Irene", "1.077.392", "1.251.434", "+16,2%", "+11,2%", "7,3%", "Bioestimulantes / Especialidades"),
        ("Javier", "973.009", "1.142.039", "+17,4%", "+13,6%", "6,6%", "Cultivos Leñosos y Frutales"),
        ("Alfonso", "529.543", "637.691", "+20,4%", "+13,8%", "3,7%", "Nuevas Cuentas / Expansión"),
        ("TOTAL CODIAGRO", "15.180.000", "17.180.000", "+13,2%", "+8,1%", "100,0%", "Consolidado Comercial")
    ]

    for r_idx, row in enumerate(com_data):
        for c_idx, val in enumerate(row):
            cell = t4.cell(r_idx, c_idx)
            cell.text = val
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.font.name = "Calibri"
            p.font.size = Pt(10) if r_idx > 0 else Pt(10.5)

            if r_idx == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = NAVY
                p.font.bold = True
                p.font.color.rgb = WHITE
                p.alignment = PP_ALIGN.CENTER if c_idx > 0 else PP_ALIGN.LEFT
            elif r_idx == 8: # Total row
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(220, 238, 225)
                p.font.bold = True
                p.font.color.rgb = NAVY
                if c_idx > 0: p.alignment = PP_ALIGN.RIGHT
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = WHITE if r_idx % 2 == 1 else RGBColor(248, 250, 252)
                p.font.color.rgb = DARK_TEXT
                if c_idx > 0: p.alignment = PP_ALIGN.RIGHT

    # Side Insight Cards
    add_card(s4, Inches(9.5), Inches(1.45), Inches(3.0), Inches(5.3))
    tb_ci = s4.shapes.add_textbox(Inches(9.7), Inches(1.6), Inches(2.6), Inches(5.0))
    tf_ci = tb_ci.text_frame
    tf_ci.word_wrap = True

    p = tf_ci.paragraphs[0]
    p.text = "CLAVES COMERCIALES 2027"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(10)

    bullets = [
        ("Liderazgo Consolidado:", "García y Ricardo representan el 62,3% de la cifra global (10,71 M€), garantizando solidez y tracción en el negocio recurrente."),
        ("Acelerador Exportación:", "Mehmet supera los 2,02 M€ (+14,0% en € / +7,5% en volumen), consolidando mercados exteriores clave con ratios ajustados."),
        ("Dinamismo en Red:", "Alfonso (+20,4%), Javier (+17,4%) e Irene (+16,2%) potencian la venta técnica de gamas biológicas de alto margen."),
        ("Ajuste Quirúrgico Opción 4:", "El plan 2027 refleja el ajuste a 17,18 M€, eliminando sobreexpectativas en Turquía, Egipto, Chile y Grecia.")
    ]
    for b_title, b_desc in bullets:
        p1 = tf_ci.add_paragraph()
        p1.text = b_title + " "
        p1.font.bold = True
        p1.font.size = Pt(10)
        p1.font.color.rgb = GREEN
        p1.space_before = Pt(8)

        p2 = tf_ci.add_paragraph()
        p2.text = b_desc
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = DARK_TEXT

    # ==========================================
    # SLIDE 5: ESTRUCTURA DE COSTES Y MARGEN BRUTO
    # ==========================================
    s5 = prs.slides.add_slide(blank_layout)
    add_header(s5, "Estructura de Aprovisionamientos y Margen Bruto (COGS)")

    cogs_items = [
        ("Materias Primas (601)", "4.462 k€", "25,97% s/ventas", "Consumo directo en formulación química. Se traslada un +4,5% de inflación contenida en 2027.", NAVY_LIGHT),
        ("Envases y Embalajes (602001)", "921 k€", "5,36% s/ventas", "Bidones, garrafas y contenedores IBC. Optimización de compras con acuerdos marco.", NAVY_LIGHT),
        ("Etiquetas y Palets (602002/3)", "303 k€", "1,76% s/ventas", "Materiales auxiliares y consumibles logísticos vinculados al envasado en planta.", NAVY_LIGHT),
        ("Subcontratación y Fabril (607)", "233 k€", "1,36% s/ventas", "Refuerzo puntual en picos de demanda estacional para maximizar servicio sin sobrecargar estructura fija.", NAVY_LIGHT),
    ]

    for i, (title, val, ratio, desc, color) in enumerate(cogs_items):
        y = Inches(1.45) + i * Inches(1.3)
        add_card(s5, Inches(0.8), y, Inches(7.5), Inches(1.15))

        tb = s5.shapes.add_textbox(Inches(1.0), y + Inches(0.12), Inches(7.1), Inches(0.95))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = f"{title}: {val}  ({ratio})"
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = NAVY

        p2 = tf.add_paragraph()
        p2.text = desc
        p2.font.size = Pt(10)
        p2.font.color.rgb = DARK_TEXT

    # Right panel: Margen Bruto Synthesis
    add_card(s5, Inches(8.6), Inches(1.45), Inches(3.9), Inches(5.3), bg_color=RGBColor(240, 253, 244), border_color=RGBColor(187, 247, 208))
    tb_gp = s5.shapes.add_textbox(Inches(8.8), Inches(1.65), Inches(3.5), Inches(4.9))
    tf_gp = tb_gp.text_frame
    tf_gp.word_wrap = True

    p = tf_gp.paragraphs[0]
    p.text = "SÍNTESIS DE MARGEN BRUTO"
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = GREEN

    p_val = tf_gp.add_paragraph()
    p_val.text = "67,3%"
    p_val.font.size = Pt(36)
    p_val.font.bold = True
    p_val.font.color.rgb = NAVY
    p_val.space_before = Pt(6)

    p_sub = tf_gp.add_paragraph()
    p_sub.text = "11.565.940 € de Gross Profit"
    p_sub.font.size = Pt(12)
    p_sub.font.bold = True
    p_sub.font.color.rgb = GREEN_ACCENT
    p_sub.space_after = Pt(14)

    p_pts = [
        "El Gross Profit en euros se incrementa en +1,45 M€ (+14,3% vs cierre 2026), alcanzando una tasa récord del 67,32% sobre ventas.",
        "La contención del incremento de compras al +4,5% estricto de inflación preserva intacto el valor añadido generado en fábrica.",
        "Mayor aportación del mix de bioestimulantes y formulaciones de alto rendimiento tecnológico en cartera."
    ]
    for pt in p_pts:
        p_item = tf_gp.add_paragraph()
        p_item.text = "• " + pt
        p_item.font.size = Pt(10)
        p_item.font.color.rgb = DARK_TEXT
        p_item.space_before = Pt(6)

    # ==========================================
    # SLIDE 6: GASTOS DE PERSONAL Y OPEX
    # ==========================================
    s6 = prs.slides.add_slide(blank_layout)
    add_header(s6, "Estructura Operativa: Gastos de Personal y Gastos Fijos (OPEX)")

    cards_opex = [
        ("Gastos de Personal", "2.228.610 €", "12,97% de Ventas", "Incremento de +222 k€ (+11,1% vs 2026) motivado por:\n• Actualización salarial convenio químico y antigüedad.\n• Refuerzo del equipo técnico de desarrollo agronómico en campo.\n• Nueva incorporación en planta para absorber mayor ritmo de envasado.\n• Mejora de eficiencia: el peso s/ventas se reduce del 13,22% al 12,97%."),
        ("Otros Gastos Operativos", "2.038.489 €", "11,87% de Ventas", "Partidas de funcionamiento diario y red comercial:\n• Formación no bonificada: 53.215 € (4.435 €/mes).\n• Gastos comerciales: 92.400 € (Pere, Ricardo y García a 2.000 €/m; Javier 1.200 €/m; Irene y Pedro 600 €/m; Ago 50%, Ene/Dic 75%).\n• Marketing y Ferias: 266 k€ | Regulatory: 234 k€ | Transporte neto: 246 k€."),
        ("Otros Ingresos Operativos", "305.223 €", "Compensación Export", "Salvedad contable de transporte:\n• Los fletes marítimos y aéreos de exportación se facturan íntegramente al cliente exterior en factura de venta (700001624).\n• Impacto en EBITDA: Totalmente neutro.\n• Transparencia y rigor normativo PGC.")
    ]

    c_w = Inches(3.75)
    for i, (title, amt, ratio, details) in enumerate(cards_opex):
        x = Inches(0.8) + i * Inches(3.95)
        add_card(s6, x, Inches(1.45), c_w, Inches(5.3))

        tb = s6.shapes.add_textbox(x + Inches(0.15), Inches(1.6), c_w - Inches(0.3), Inches(5.0))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = title.upper()
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = GREEN_ACCENT if i == 2 else NAVY_LIGHT

        p1 = tf.add_paragraph()
        p1.text = amt
        p1.font.size = Pt(24)
        p1.font.bold = True
        p1.font.color.rgb = NAVY

        p2 = tf.add_paragraph()
        p2.text = ratio
        p2.font.size = Pt(11)
        p2.font.bold = True
        p2.font.color.rgb = GREEN if i == 0 else NAVY_LIGHT
        p2.space_after = Pt(12)

        p3 = tf.add_paragraph()
        p3.text = details
        p3.font.size = Pt(10)
        p3.font.color.rgb = DARK_TEXT

    # ==========================================
    # SLIDE 7: PRESUPUESTOS DEPARTAMENTALES VS P&L CONTABLE
    # ==========================================
    s7 = prs.slides.add_slide(blank_layout)
    add_header(s7, "Conciliación Departamental: Presupuestos de Área vs Criterio PGC")

    t7_shape = s7.shapes.add_table(7, 5, Inches(0.8), Inches(1.45), Inches(8.5), Inches(5.3))
    t7 = t7_shape.table
    t7.columns[0].width = Inches(2.5)
    t7.columns[1].width = Inches(1.4)
    t7.columns[2].width = Inches(1.4)
    t7.columns[3].width = Inches(1.4)
    t7.columns[4].width = Inches(1.8)

    dept_rows = [
        ("Departamento / Área", "Ppto Depto (€)", "Ppto P&L (€)", "Diferencia (€)", "Criterio de Conciliación PGC"),
        ("Compras y Aprovisionamientos", "6.120.000", "5.919.283", "+200.717", "Ajuste de consumo real vs compras con stock"),
        ("Transportes (Neto Codiagro)", "245.670", "301.549", "-55.879", "Exportación (305 k€) repercutida al cliente"),
        ("Marketing y Comunicación", "266.072", "49.697", "+216.375", "Reclasificación ferias y campañas directas"),
        ("Regulatory y Registros", "233.918", "105.551", "+128.367", "Registro de marcas UE activable (203/623)"),
        ("Mantenimiento de Fábrica", "91.234", "197.422", "-106.188", "Subcontratación repuestos y revisiones"),
        ("TOTAL GASTOS DEPARTAMENTOS", "6.956.894", "6.573.502", "+383.392", "Conciliado con Estados Financieros")
    ]

    for r_idx, row in enumerate(dept_rows):
        for c_idx, val in enumerate(row):
            cell = t7.cell(r_idx, c_idx)
            cell.text = val
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.font.name = "Calibri"
            p.font.size = Pt(9.5) if r_idx > 0 else Pt(10)

            if r_idx == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = NAVY
                p.font.bold = True
                p.font.color.rgb = WHITE
                p.alignment = PP_ALIGN.CENTER if c_idx > 0 else PP_ALIGN.LEFT
            elif r_idx == 6:
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(220, 238, 225)
                p.font.bold = True
                p.font.color.rgb = NAVY
                if c_idx > 0: p.alignment = PP_ALIGN.RIGHT
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = WHITE if r_idx % 2 == 1 else RGBColor(248, 250, 252)
                p.font.color.rgb = DARK_TEXT
                if c_idx > 0: p.alignment = PP_ALIGN.RIGHT

    # Side Box: Auditing note
    add_card(s7, Inches(9.5), Inches(1.45), Inches(3.0), Inches(5.3))
    tb_aud = s7.shapes.add_textbox(Inches(9.7), Inches(1.6), Inches(2.6), Inches(5.0))
    tf_aud = tb_aud.text_frame
    tf_aud.word_wrap = True

    p = tf_aud.paragraphs[0]
    p.text = "GARANTÍA DE CONTROL"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(10)

    notes = [
        ("Salvedad de Transporte:", "Queda formalmente documentado que el coste de flete internacional no incrementa el OPEX neto de la compañía."),
        ("Coordinación Interdepartamental:", "Cada responsable de área ha validado su desglose mensual, asegurando compromiso operativo con el techo de gasto."),
        ("Criterio Conservador:", "Las partidas activables (como registros internacionales de marcas) se han analizado bajo estricto cumplimiento del PGC.")
    ]
    for n_title, n_desc in notes:
        p1 = tf_aud.add_paragraph()
        p1.text = n_title
        p1.font.bold = True
        p1.font.size = Pt(10)
        p1.font.color.rgb = GREEN
        p1.space_before = Pt(8)

        p2 = tf_aud.add_paragraph()
        p2.text = n_desc
        p2.font.size = Pt(9.5)
        p2.font.color.rgb = DARK_TEXT

    # ==========================================
    # SLIDE 8: CUMPLIMIENTO EN TIEMPO REAL (SEPTIEMBRE 2026)
    # ==========================================
    s8 = prs.slides.add_slide(blank_layout)
    add_header(s8, "Tracción Inmediata: Validación con el Cierre de Septiembre 2026")

    # 4 Cards in 2x2 grid
    grid_data = [
        ("OBJETIVO MENSUAL PREVISTO", "181.966 uds", "Presupuesto Septiembre", "Cifra calculada sin factores de aplanamiento artificial (+5,2% en unidades originales).", NAVY_LIGHT),
        ("PEDIDOS REALES REGISTRADOS", "236.105 uds", "+129,8% de Cumplimiento", "Sobrecumplimiento neto de +54.139 unidades respecto a la previsión inicial.", GREEN),
        ("CAPACIDAD DE SERVICIO", "200.331 uds", "84,8% Servido Inmediato", "200k unidades ya expedidas y 35.774 unidades en preparación en fábrica.", NAVY_LIGHT),
        ("IMPORTE NETO FACTURADO", "735.161 €", "Corte al 01/10/2026", "Facturación neta consolidada correspondiente a las órdenes del periodo.", GREEN)
    ]

    for i, (title, val, badge, desc, col) in enumerate(grid_data):
        gx = Inches(0.8) + (i % 2) * Inches(4.3)
        gy = Inches(1.45) + (i // 2) * Inches(2.65)
        add_card(s8, gx, gy, Inches(4.1), Inches(2.45))

        tb = s8.shapes.add_textbox(gx + Inches(0.2), gy + Inches(0.2), Inches(3.7), Inches(2.0))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = title
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = GRAY_TEXT

        p1 = tf.add_paragraph()
        p1.text = val
        p1.font.size = Pt(26)
        p1.font.bold = True
        p1.font.color.rgb = NAVY
        p1.space_before = Pt(4)

        p2 = tf.add_paragraph()
        p2.text = badge
        p2.font.size = Pt(11)
        p2.font.bold = True
        p2.font.color.rgb = col
        p2.space_before = Pt(2)
        p2.space_after = Pt(6)

        p3 = tf.add_paragraph()
        p3.text = desc
        p3.font.size = Pt(9.5)
        p3.font.color.rgb = DARK_TEXT

    # Right side: What this means for 2027
    add_card(s8, Inches(9.4), Inches(1.45), Inches(3.1), Inches(5.3))
    tb_imp = s8.shapes.add_textbox(Inches(9.6), Inches(1.65), Inches(2.7), Inches(4.9))
    tf_imp = tb_imp.text_frame
    tf_imp.word_wrap = True

    p = tf_imp.paragraphs[0]
    p.text = "¿QUÉ DEMUESTRA ESTE RESULTADO?"
    p.font.size = Pt(11)
    p.font.bold = True
    p.font.color.rgb = NAVY
    p.space_after = Pt(10)

    insights = [
        "Confianza del Mercado: El salto de demanda anticipa una excelente campaña agrícola 2026/27.",
        "Aceptación de Tarifas: La actualización de precios no ha erosionado la cartera de pedidos de los distribuidores.",
        "Respaldo al Plan 2027: Los 17,18 M€ presupuestados para 2027 están sólidamente fundamentados en tracción real de compra.",
        "Control en Tiempo Real: El nuevo dashboard automatizado permite detectar desviaciones a nivel cliente y SKU al día siguiente."
    ]
    for ins in insights:
        p_i = tf_imp.add_paragraph()
        p_i.text = "✓ " + ins
        p_i.font.size = Pt(10)
        p_i.font.color.rgb = DARK_TEXT
        p_i.space_before = Pt(8)

    # ==========================================
    # SLIDE 9: ANÁLISIS DE SENSIBILIDAD Y RIESGOS
    # ==========================================
    s9 = prs.slides.add_slide(blank_layout)
    add_header(s9, "Gestión de Riesgos y Análisis de Sensibilidad del Escenario Presupuestario")

    # 3 Scenarios
    scenarios = [
        ("Escenario Optimista (+5% Ventas)", "18,04 M€", "EBITDA: 7,85 M€ (43,5%)", "Mayor penetración internacional y campaña agronómica adelantada.", GREEN),
        ("Escenario Base Presupuestado", "17,18 M€", "EBITDA: 7,30 M€ (42,5%)", "Hipótesis central aprobada, con crecimiento del +13,2% y control de costes.", NAVY_LIGHT),
        ("Escenario Conservador (-5% Ventas)", "16,32 M€", "EBITDA: 6,75 M€ (41,4%)", "Retraso en pedidos climáticos; margen EBITDA preservado por flexibilidad fabril.", GOLD),
    ]

    for i, (name, rev, ebitda, text, color) in enumerate(scenarios):
        x = Inches(0.8) + i * Inches(3.95)
        add_card(s9, x, Inches(1.45), Inches(3.75), Inches(2.2))

        tb = s9.shapes.add_textbox(x + Inches(0.15), Inches(1.55), Inches(3.45), Inches(2.0))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = name
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = color

        p1 = tf.add_paragraph()
        p1.text = rev
        p1.font.size = Pt(22)
        p1.font.bold = True
        p1.font.color.rgb = NAVY

        p2 = tf.add_paragraph()
        p2.text = ebitda
        p2.font.size = Pt(11)
        p2.font.bold = True
        p2.font.color.rgb = color
        p2.space_after = Pt(4)

        p3 = tf.add_paragraph()
        p3.text = text
        p3.font.size = Pt(9.5)
        p3.font.color.rgb = DARK_TEXT

    # Bottom Table of Critical Risks & Mitigation
    t9_shape = s9.shapes.add_table(4, 3, Inches(0.8), Inches(3.9), Inches(11.733), Inches(2.9))
    t9 = t9_shape.table
    t9.columns[0].width = Inches(2.8)
    t9.columns[1].width = Inches(2.4)
    t9.columns[2].width = Inches(6.533)

    risk_rows = [
        ("Riesgo Identificado", "Nivel de Impacto", "Medida Preventiva / Plan de Mitigación Codiagro"),
        ("1. Volatilidad en Precios de Materias Primas", "Medio-Alto", "Contratos de cobertura y compras anticipadas para el 60% de las necesidades del primer semestre; fórmulas de revisión de precios con distribuidores."),
        ("2. Fricciones Regulatorias y Registro de Marcas", "Medio", "Presupuesto específico de 234 k€ en Regulatory; priorización de expedientes con mayor retorno de margen y apoyo de consultoría externa especializada."),
        ("3. Costes Logísticos y Fletes de Exportación", "Bajo (Neutro)", "Salvedad contractual: fletes marítimos/aéreos 100% repercutidos en factura comercial al cliente importador, blindando la rentabilidad neta."),
    ]

    for r_idx, row in enumerate(risk_rows):
        for c_idx, val in enumerate(row):
            cell = t9.cell(r_idx, c_idx)
            cell.text = val
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.font.name = "Calibri"
            p.font.size = Pt(9.5) if r_idx > 0 else Pt(10)

            if r_idx == 0:
                cell.fill.solid()
                cell.fill.fore_color.rgb = NAVY
                p.font.bold = True
                p.font.color.rgb = WHITE
                p.alignment = PP_ALIGN.CENTER if c_idx == 1 else PP_ALIGN.LEFT
            else:
                cell.fill.solid()
                cell.fill.fore_color.rgb = WHITE if r_idx % 2 == 1 else RGBColor(248, 250, 252)
                p.font.color.rgb = ALERT_RED if c_idx == 1 and 'Alto' in val else DARK_TEXT
                if c_idx == 1: p.alignment = PP_ALIGN.CENTER

    # ==========================================
    # SLIDE 10: CALENDARIO DE SEGUIMIENTO Y ROLLING FORECAST
    # ==========================================
    s10 = prs.slides.add_slide(blank_layout)
    add_header(s10, "Gobernanza y Control de Gestión: Calendario de Seguimiento 2027")

    phases = [
        ("Q1 2027 · Arranque y Campaña", "Ene - Mar", [
            "Lanzamiento de nuevas tarifas y contratos anuales.",
            "Revisión de consumo de materias primas frente al ratio 25,97%.",
            "Cierre y balance del primer trimestre con análisis de desviaciones."
        ]),
        ("Q2 2027 · Pico de Ventas y Envasado", "Abr - Jun", [
            "Meses clave de volumen fabril (Mayo y Junio >750 k€ en compras).",
            "Control riguroso de subcontratación y horas de planta.",
            "Revisión semestral del plan de marketing y asistencia a ferias."
        ]),
        ("Q3 2027 · Campaña de Otoño y Export", "Jul - Sep", [
            "Cierre del verano y arranque de pedidos de otoño.",
            "Actualización del Rolling Forecast 2027/2028.",
            "Auditoría intermedia de stock de envases y palets."
        ]),
        ("Q4 2027 · Cierre y Regularización", "Oct - Dic", [
            "Regularización de rappels anuales y convenios comerciales.",
            "Inventario físico anual y ajuste de variación de existencias.",
            "Aprobación definitiva de cuentas del ejercicio."
        ])
    ]

    for i, (p_name, p_time, items) in enumerate(phases):
        x = Inches(0.8) + i * Inches(2.95)
        add_card(s10, x, Inches(1.45), Inches(2.8), Inches(5.3))

        tb = s10.shapes.add_textbox(x + Inches(0.15), Inches(1.6), Inches(2.5), Inches(5.0))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = p_time.upper()
        p.font.size = Pt(10)
        p.font.bold = True
        p.font.color.rgb = GREEN_ACCENT

        p1 = tf.add_paragraph()
        p1.text = p_name
        p1.font.size = Pt(12)
        p1.font.bold = True
        p1.font.color.rgb = NAVY
        p1.space_after = Pt(14)

        for it in items:
            p_it = tf.add_paragraph()
            p_it.text = "• " + it
            p_it.font.size = Pt(9.5)
            p_it.font.color.rgb = DARK_TEXT
            p_it.space_before = Pt(8)

    # ==========================================
    # SLIDE 11: CONCLUSIONES Y DECISIONES A APROBAR
    # ==========================================
    s11 = prs.slides.add_slide(blank_layout)
    add_header(s11, "Conclusiones y Decisiones Solicitadas a Dirección General")

    # 3 Big Recommendation Cards
    decisions = [
        ("1. Aprobación del Marco Financiero 2027", "Presupuesto Oficial", [
            "Aprobar la meta de Cifra Neta de Negocios en 17,18 M€ (+13,2% vs 2026).",
            "Ratificar el objetivo de EBITDA Ajustado en 7,30 M€ (42,5% de margen).",
            "Validar el techo de gasto de aprovisionamientos (COGS) en 5,92 M€."
        ], GREEN),
        ("2. Autorización de Inversiones y Estructura", "Plan de Recursos", [
            "Aprobar el presupuesto de personal de 2,23 M€ con refuerzo técnico y fabril.",
            "Validar la dotación de 266 k€ para Marketing estratégico y ferias sectoriales.",
            "Autorizar la partida de 234 k€ en Regulatory para protección y registros internacionales."
        ], NAVY_LIGHT),
        ("3. Criterios de Seguimiento y Gobernanza", "Control y Reporting", [
            "Adopción del nuevo Dashboard Web de Control Presupuestario como herramienta oficial de monitorización continua.",
            "Revisión formal trimestral con Dirección de cada zona comercial.",
            "Mantenimiento de la política de neutralidad en transporte exterior repercutido."
        ], NAVY_LIGHT),
    ]

    for i, (title, tag, points, col) in enumerate(decisions):
        y = Inches(1.45) + i * Inches(1.75)
        add_card(s11, Inches(0.8), y, Inches(11.733), Inches(1.6))

        tb = s11.shapes.add_textbox(Inches(1.0), y + Inches(0.12), Inches(11.3), Inches(1.4))
        tf = tb.text_frame
        tf.word_wrap = True

        p = tf.paragraphs[0]
        p.text = f"{title.upper()}  [{tag}]"
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = col

        for pt in points:
            p_pt = tf.add_paragraph()
            p_pt.text = "✓  " + pt
            p_pt.font.size = Pt(10)
            p_pt.font.color.rgb = DARK_TEXT
            p_pt.space_before = Pt(3)

    # Save
    out_file = 'Presentacion_Presupuesto_2027_Codiagro.pptx'
    prs.save(out_file)
    print(f"Presentación PowerPoint generada exitosamente: {out_file}")

if __name__ == '__main__':
    create_presentation()
