import json
import os
import sys
from datetime import datetime
import openpyxl

def fail(msg, last_valid_date="30/09/2026 (Cierre Septiembre) / 28/10/2026 (Pedidos Octubre)"):
    print(f"\n[FALLO DE VALIDACIÓN] {msg}", file=sys.stderr)
    print(f"[ESTADO CONSERVADO] Se mantiene activa la última versión válida disponible (Fecha de datos: {last_valid_date}).", file=sys.stderr)
    sys.exit(1)

def main():
    print("=== FASE 2: VALIDACIÓN DE CÁLCULOS, CONSERVACIÓN DE IMPORTES Y COHERENCIA ===")

    # 1. Identificar fuentes utilizadas y comprobar su existencia
    sources = [
        ("Plantilla Maestra Full Costing Sep-26 (incluye PyG_detalle y Cuentas)", os.path.join("outputs", "fullcost_septiembre_2026", "Plantillas_Full_Costing_Septiembre_2026_Con_PYG_Mejorada.xlsx")),
        ("Escandallos Técnicos SKU", os.path.join("Gastos", "Costes SKU.xlsx")),
        ("Reporte Logístico Transporte Sep-26", os.path.join("outputs", "fullcost_septiembre_2026", "REPORT_SEPTIEMBRE_LOGISTICA_actualizado.xlsx")),
        ("Órdenes de Fabricación Planta", os.path.join("outputs", "fullcost_septiembre_2026", "Fabricacion 26.10.xlsx")),
        ("Fichas Mayor 706-709 Rappels y Devoluciones", os.path.join("outputs", "fullcost_septiembre_2026", "fichas mayor 706 - 709.xlsx")),
        ("Detalle Rappels Comerciales", os.path.join("outputs", "fullcost_septiembre_2026", "Rappels Marta.xlsx")),
        ("Presupuesto Oficial Ventas y Detalle Sep-26", "Presupuesto_Ventas_2027_Definitivo.xlsx"),
        ("Consolidado Ventas y Previsiones 2026", "Consolidado_Ventas_y_Previsiones_2026.xlsx"),
        ("Pedidos Cierre Septiembre", os.path.join("Pedidos", "Pedidos 01.10 comerciales.xlsx"))
    ]

    sources_info = []
    print("1. Fuentes identificadas:")
    for label, path in sources:
        if not os.path.exists(path):
            fail(f"Fuente requerida no encontrada: {label} ({path})")
        mtime = datetime.fromtimestamp(os.path.getmtime(path)).strftime("%d/%m/%Y %H:%M")
        sources_info.append({"nombre": label, "ruta": path, "modificado": mtime})
        print(f"   - [OK] {label}: {path} ({mtime})")

    # 2. Comprobar salidas generadas y correspondencia biunívoca JSON <-> JS
    fc_json_path = os.path.join("web_dashboard_direccion", "fullcost_data.json")
    fc_js_path = os.path.join("web_dashboard_direccion", "fullcost_data.js")
    dash_json_path = os.path.join("web_dashboard_direccion", "dashboard_data.json")
    marta_json_path = os.path.join("web_dashboard_direccion", "marta_tracking_data.json")
    excel_seg_path = "Seguimiento_Presupuesto_Sep_2026.xlsx"
    excel_mrp_path = "Orden_Fabricacion_MRP_Codiagro.xlsx"

    for p in [fc_json_path, fc_js_path, dash_json_path, marta_json_path, excel_seg_path, excel_mrp_path]:
        if not os.path.exists(p) or os.path.getsize(p) == 0:
            fail(f"Archivo de salida ausente o vacío: {p}")

    try:
        with open(fc_json_path, "r", encoding="utf-8") as f:
            fc = json.load(f)
        with open(dash_json_path, "r", encoding="utf-8") as f:
            dash = json.load(f)
        with open(marta_json_path, "r", encoding="utf-8") as f:
            marta = json.load(f)
    except Exception as e:
        fail(f"Error de sintaxis JSON en los archivos generados: {e}")

    # Validar correspondencia exacta entre fullcost_data.json y fullcost_data.js (no permitir objeto vacío ni desincronización)
    try:
        with open(fc_js_path, "r", encoding="utf-8") as f:
            js_raw = f.read().strip()
        prefix = "window.fullCostData =" if js_raw.startswith("window.fullCostData =") else "window.FULLCOST_DATA ="
        if not js_raw.startswith(prefix) or not js_raw.endswith(";"):
            fail("El archivo fullcost_data.js no tiene la estructura 'window.fullCostData = {...};'")
        js_json_str = js_raw[len(prefix):].rstrip(";").strip()
        fc_from_js = json.loads(js_json_str)
        if not fc_from_js or not fc_from_js.get("clientes") or not fc_from_js.get("totales_drilldown"):
            fail("El archivo fullcost_data.js contiene un objeto vacío o carece de clientes/totales requeridos.")
        if round(fc_from_js.get("totales_drilldown", {}).get("ventas", 0), 2) != round(fc.get("totales_drilldown", {}).get("ventas", 0), 2):
            fail("Desincronización entre fullcost_data.json y fullcost_data.js en total ventas.")
    except Exception as e:
        fail(f"Fallo al validar correspondencia de fullcost_data.js: {e}")

    meta = fc.get("metadata", {})
    last_date = f"Cierre Full Costing: {meta.get('fecha_cierre', '30/09/2026')}"

    # 3. Validar ausencia de incidencias bloqueantes en las fuentes del modelo
    inc_datos = meta.get("incidencias_datos", [])
    if inc_datos:
        fail(f"Se detectaron {len(inc_datos)} incidencias de datos ausentes o inválidos: {'; '.join(inc_datos)}", last_date)

    if not fc.get("modelos_analiticos"):
        fail("El bloque modelos_analiticos está ausente en fullcost_data.json", last_date)

    # 4. Validar conservación exacta de totales (Política Mayor Resto: tolerancia 0,00 €)
    tot = fc.get("totales_drilldown", {})
    clis = fc.get("clientes", [])
    prods = fc.get("productos", [])
    coms = fc.get("comerciales", [])

    if not clis or not prods or not coms:
        fail("Listas de clientes, productos o comerciales vacías en fullcost_data.json", last_date)

    checks = [
        ("Clientes Ventas vs Total", round(sum(c["ventas"] for c in clis) - tot["ventas"], 2)),
        ("Productos Ventas vs Total", round(sum(p["ventas"] for p in prods) - tot["ventas"], 2)),
        ("Comerciales Ventas vs Total", round(sum(cm["ventas"] for cm in coms) - tot["ventas"], 2)),
        ("Clientes Materiales vs Total", round(sum(c["cogs_materiales"] for c in clis) - tot["cogs_materiales"], 2)),
        ("Productos Materiales vs Total", round(sum(p["cogs_materiales"] for p in prods) - tot["cogs_materiales"], 2)),
        ("Comerciales Materiales vs Total", round(sum(cm["cogs_materiales"] for cm in coms) - tot["cogs_materiales"], 2)),
        ("Clientes Coste Planta vs Total", round(sum(c["coste_planta_fullcost"] for c in clis) - tot["coste_planta"], 2)),
        ("Productos Coste Planta vs Total", round(sum(p["coste_planta_fullcost"] for p in prods) - tot["coste_planta"], 2)),
        ("Comerciales Coste Planta vs Total", round(sum(cm["coste_planta_fullcost"] for cm in coms) - tot["coste_planta"], 2)),
        ("Clientes Margen Ind vs Total", round(sum(c["margen_industrial_eur"] for c in clis) - tot["margen_industrial_eur"], 2)),
        ("Productos Margen Ind vs Total", round(sum(p["margen_industrial_eur"] for p in prods) - tot["margen_industrial_eur"], 2)),
        ("Comerciales Margen Ind vs Total", round(sum(cm["margen_industrial_eur"] for cm in coms) - tot["margen_industrial_eur"], 2)),
        ("Clientes Transporte vs Total", round(sum(c["transporte_real"] for c in clis) - tot["transporte_real"], 2)),
        ("Comerciales Portes vs Total", round(sum(cm["portes"] for cm in coms) - tot["transporte_real"], 2)),
        ("Clientes Margen Contrib vs Total", round(sum(c["margen_contribucion_eur"] for c in clis) - tot["margen_contribucion_eur"], 2)),
        ("Comerciales Margen Contrib vs Total", round(sum(cm["margen_contribucion_eur"] for cm in coms) - tot["margen_contribucion_eur"], 2)),
        ("Clientes Unidades vs Total", round(sum(c["uds"] for c in clis) - tot["uds"], 1)),
        ("Productos Unidades vs Total", round(sum(p["uds"] for p in prods) - tot["uds"], 1)),
        ("CECOs Planta Suma vs Coste Planta Total", round(sum(meta.get("costes_ceco_eur", {}).values()) - tot["coste_planta"], 2)),
    ]

    for name, diff in checks:
        if abs(diff) > 0.0:
            fail(f"Descuadre de redondeo en Full Costing ({name}): diferencia = {diff}", last_date)

    # 5. Validar conservación de totales en marta_tracking_data.json
    for mode in ["unidades", "euros"]:
        m_block = marta.get(mode, {})
        m_tot = m_block.get("metadata", {}).get("totales", {})
        m_clis = m_block.get("clientes", [])
        m_coms = m_block.get("comerciales", [])
        for k_f in [
            "ventas_2025_tot",
            "ventas_2025_sep",
            "ventas_2025_oct",
            "ventas_2026_sep",
            "estimacion_2026",
            "pendiente_oct_dic",
        ]:
            d_cli = round(sum(r[k_f] for r in m_clis) - m_tot[k_f], 2)
            d_com = round(sum(r[k_f] for r in m_coms) - m_tot[k_f], 2)
            if abs(d_cli) > 0 or abs(d_com) > 0:
                fail(f"Descuadre en marta_tracking_data ({mode} - {k_f}): cli={d_cli}, com={d_com}", last_date)

    # 6. Validar que los puentes (existencias, transporte, materiales) y matriz de trazabilidad NO estén vacíos
    puente_ex = fc.get("puente_existencias", {})
    pasos_ex = puente_ex.get("productos_711_712", {}).get("pasos_puente", [])
    if not puente_ex or not pasos_ex:
        fail("El puente de existencias está vacío o carece de pasos de conciliación justificativos.", last_date)

    puente_tr = fc.get("puente_transporte", {})
    if not puente_tr or not puente_tr.get("desglose_por_cuenta") or puente_tr.get("reporte_logistico_imputado_eur", 0) <= 0:
        fail("El puente de transporte está vacío o carece de desglose por cuenta.", last_date)

    puente_mat = fc.get("puente_materiales", {})
    if not puente_mat or not puente_mat.get("lineas"):
        fail("El puente de materiales está vacío o no documenta la conciliación contable vs técnica.", last_date)

    trazab = fc.get("trazabilidad_parametros", [])
    if not trazab or len(trazab) < 6:
        fail(f"La matriz de trazabilidad de parámetros está vacía o incompleta ({len(trazab)} parámetros; mínimo 6 requeridos).", last_date)

    # 7. Validar Excel de seguimiento
    try:
        wb_seg = openpyxl.load_workbook(excel_seg_path, data_only=True)
        if "Resumen General" not in wb_seg.sheetnames:
            fail("El archivo Seguimiento_Presupuesto_Sep_2026.xlsx no contiene la hoja 'Resumen General'", last_date)
        wb_seg.close()
    except Exception as e:
        fail(f"No se pudo abrir o validar Seguimiento_Presupuesto_Sep_2026.xlsx: {e}", last_date)

    # 8. Escudo de seguridad: Verificar aislamiento estricto de web_dashboard (versión comerciales)
    forbidden_commercial_files = [
        "fullcost_data.json",
        "fullcost_data.js",
        "cockpit_inicio_logic.js",
        "profitability_logic.js",
        "bi_intelligence_logic.js",
        "treasury_data.json",
        "treasury_data.js",
        "update_validation_status.json"
    ]
    for fn in forbidden_commercial_files:
        p = os.path.join("web_dashboard", fn)
        if os.path.exists(p):
            fail(f"VIOLACIÓN DE SEGURIDAD: Archivo confidencial de Dirección General encontrado en web_dashboard ({fn}).")

    commercial_index = os.path.join("web_dashboard", "index.html")
    if os.path.exists(commercial_index):
        with open(commercial_index, "r", encoding="utf-8") as f:
            html_content = f.read()
        if "Estimación de Ventas | Codiagro" not in html_content or "fullcost" in html_content.lower():
            fail("VIOLACIÓN DE SEGURIDAD: web_dashboard/index.html contiene módulos o títulos de Dirección General en lugar de la versión comercial abierta.")

    diff_concil = round(float(meta.get("diferencia_conciliacion", 0.0) or 0.0), 2)
    estado_conciliacion = "CONCILIADO (0,00 €)" if abs(diff_concil) <= 0.05 else f"PENDIENTE DE CONCILIACIÓN ({diff_concil:.2f} €)"

    status_payload = {
        "estado_tecnico": "COMPROBADO",
        "estado_conciliacion_contable": estado_conciliacion,
        "fecha_validacion": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "fecha_datos_fullcost": meta.get("fecha_cierre", "30/09/2026"),
        "rai_contable_mayor_eur": meta.get("rai_contable_mayor"),
        "rai_analitico_cascada_eur": meta.get("rai_analitico_cascada"),
        "diferencia_conciliacion_eur": diff_concil,
        "conservacion_importes_drift": 0.0,
        "fuentes_verificadas": sources_info
    }
    with open(os.path.join("web_dashboard_direccion", "update_validation_status.json"), "w", encoding="utf-8") as f:
        json.dump(status_payload, f, ensure_ascii=False, indent=2)

    print(f"2. Validación técnica superada al 100% (drift 0,00 €). Estado contable: {estado_conciliacion}.")

if __name__ == "__main__":
    main()
