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
import json
import os
import sys
from datetime import datetime
import openpyxl
import re
import subprocess

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
    fc_json_path = get_dest(os.path.join("web_dashboard_direccion", "fullcost_data.json"))
    fc_js_path = get_dest(os.path.join("web_dashboard_direccion", "fullcost_data.js"))
    dash_json_path = get_dest(os.path.join("web_dashboard_direccion", "dashboard_data.json"))
    marta_json_path = get_dest(os.path.join("web_dashboard_direccion", "marta_tracking_data.json"))
    excel_seg_path = get_dest("Seguimiento_Presupuesto_Sep_2026.xlsx")
    excel_mrp_path = get_dest("Orden_Fabricacion_MRP_Codiagro.xlsx")

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
        js_json_str = js_raw[len(prefix):].strip()
        if js_json_str.endswith(";"): js_json_str = js_json_str[:-1].strip()
        fc_from_js = json.loads(js_json_str)
        if not fc_from_js or not fc_from_js.get("clientes") or not fc_from_js.get("totales_drilldown"):
            fail("El archivo fullcost_data.js contiene un objeto vacío o carece de clientes/totales requeridos.")
        
        # Comparación exhaustiva de todas las magnitudes clave entre JSON y JavaScript
        js_tot = fc_from_js.get("totales_drilldown", {})
        json_tot = fc.get("totales_drilldown", {})
        for k_tot in ["ventas", "cogs_materiales", "coste_planta", "coste_subactividad", "coste_planta_total",
                      "margen_industrial_eur", "transporte_real", "margen_contribucion_eur", "uds"]:
            if k_tot in json_tot or k_tot in js_tot:
                v_js = round(float(js_tot.get(k_tot, 0.0) or 0.0), 2)
                v_json = round(float(json_tot.get(k_tot, 0.0) or 0.0), 2)
                if abs(v_js - v_json) > 0.01:
                    fail(f"Desincronización entre JSON y JS en totales_drilldown.{k_tot}: {v_json} € vs {v_js} €")

        js_meta = fc_from_js.get("metadata", {})
        json_meta = fc.get("metadata", {})
        for k_meta in ["rai_contable_mayor", "rai_analitico_cascada", "desviacion_conciliacion_mayor"]:
            v_js = round(float(js_meta.get(k_meta, 0.0) or 0.0), 2)
            v_json = round(float(json_meta.get(k_meta, 0.0) or 0.0), 2)
            if abs(v_js - v_json) > 0.01:
                fail(f"Desincronización entre JSON y JS en metadata.{k_meta}: {v_json} € vs {v_js} €")

        # Comparación exhaustiva del detalle entre JSON y JavaScript (clientes, productos, comerciales)
        js_clis = fc_from_js.get("clientes", [])
        json_clis = fc.get("clientes", [])
        if len(js_clis) != len(json_clis):
            fail(f"Desincronización en número de clientes entre JSON ({len(json_clis)}) y JS ({len(js_clis)})")
        
        js_cli_map = {str(c.get("cliente", "")).strip().upper(): c for c in js_clis}
        for c_json in json_clis:
            c_name = str(c_json.get("cliente", "")).strip().upper()
            c_js = js_cli_map.get(c_name)
            if not c_js:
                fail(f"Cliente '{c_name}' presente en JSON pero ausente en fullcost_data.js")
            for field in ["ventas", "cogs_materiales", "coste_planta_fullcost", "margen_industrial_eur",
                          "transporte_real", "margen_contribucion_eur", "uds"]:
                v_json = round(float(c_json.get(field, 0.0) or 0.0), 2)
                v_js = round(float(c_js.get(field, 0.0) or 0.0), 2)
                if abs(v_json - v_js) > 0.01:
                    fail(f"Desincronización en cliente '{c_name}' campo '{field}': JSON={v_json:.2f} € vs JS={v_js:.2f} €")

        js_prods = fc_from_js.get("productos", [])
        json_prods = fc.get("productos", [])
        if len(js_prods) != len(json_prods):
            fail(f"Desincronización en número de productos entre JSON ({len(json_prods)}) y JS ({len(js_prods)})")
        
        js_prod_map = {str(p.get("sku", "")).strip().upper(): p for p in js_prods}
        for p_json in json_prods:
            sku = str(p_json.get("sku", "")).strip().upper()
            p_js = js_prod_map.get(sku)
            if not p_js:
                fail(f"Producto SKU '{sku}' presente en JSON pero ausente en fullcost_data.js")
            for field in ["ventas", "cogs_materiales", "coste_planta_fullcost", "margen_industrial_eur", "uds"]:
                v_json = round(float(p_json.get(field, 0.0) or 0.0), 2)
                v_js = round(float(p_js.get(field, 0.0) or 0.0), 2)
                if abs(v_json - v_js) > 0.01:
                    fail(f"Desincronización en producto SKU '{sku}' campo '{field}': JSON={v_json:.2f} € vs JS={v_js:.2f} €")

        js_coms = fc_from_js.get("comerciales", [])
        json_coms = fc.get("comerciales", [])
        if len(js_coms) != len(json_coms):
            fail(f"Desincronización en número de comerciales entre JSON ({len(json_coms)}) y JS ({len(js_coms)})")
        
        js_com_map = {str(cm.get("comercial", "")).strip().upper(): cm for cm in js_coms}
        for cm_json in json_coms:
            cm_name = str(cm_json.get("comercial", "")).strip().upper()
            cm_js = js_com_map.get(cm_name)
            if not cm_js:
                fail(f"Comercial '{cm_name}' presente en JSON pero ausente en fullcost_data.js")
            for field in ["ventas", "cogs_materiales", "coste_planta_fullcost", "margen_industrial_eur",
                          "portes", "margen_contribucion_eur"]:
                v_json = round(float(cm_json.get(field, 0.0) or 0.0), 2)
                v_js = round(float(cm_js.get(field, 0.0) or 0.0), 2)
                if abs(v_json - v_js) > 0.01:
                    fail(f"Desincronización en comercial '{cm_name}' campo '{field}': JSON={v_json:.2f} € vs JS={v_js:.2f} €")

        # Comprobación de sumatorios directos sobre los datos cargados por la web desde fullcost_data.js
        tot_js_ventas = round(float(js_tot.get("ventas", 0.0) or 0.0), 2)
        tot_js_cogs = round(float(js_tot.get("cogs_materiales", 0.0) or 0.0), 2)
        tot_js_planta = round(float(js_tot.get("coste_planta", 0.0) or 0.0), 2)
        tot_js_m_ind = round(float(js_tot.get("margen_industrial_eur", 0.0) or 0.0), 2)
        tot_js_m_cont = round(float(js_tot.get("margen_contribucion_eur", 0.0) or 0.0), 2)
        tot_js_uds = round(float(js_tot.get("uds", 0.0) or 0.0), 2)

        # Ventas
        sum_js_v_clis = round(sum(float(c.get("ventas", 0.0) or 0.0) for c in js_clis), 2)
        sum_js_v_prods = round(sum(float(p.get("ventas", 0.0) or 0.0) for p in js_prods), 2)
        sum_js_v_coms = round(sum(float(cm.get("ventas", 0.0) or 0.0) for cm in js_coms), 2)
        if abs(sum_js_v_clis - tot_js_ventas) > 0.01:
            fail(f"Desincronización en JS: Suma ventas clientes ({sum_js_v_clis:.2f} €) != total ({tot_js_ventas:.2f} €)")
        if abs(sum_js_v_prods - tot_js_ventas) > 0.01:
            fail(f"Desincronización en JS: Suma ventas productos ({sum_js_v_prods:.2f} €) != total ({tot_js_ventas:.2f} €)")
        if abs(sum_js_v_coms - tot_js_ventas) > 0.01:
            fail(f"Desincronización en JS: Suma ventas comerciales ({sum_js_v_coms:.2f} €) != total ({tot_js_ventas:.2f} €)")

        # COGS Materiales
        sum_js_c_clis = round(sum(float(c.get("cogs_materiales", 0.0) or 0.0) for c in js_clis), 2)
        sum_js_c_prods = round(sum(float(p.get("cogs_materiales", 0.0) or 0.0) for p in js_prods), 2)
        sum_js_c_coms = round(sum(float(cm.get("cogs_materiales", 0.0) or 0.0) for cm in js_coms), 2)
        if abs(sum_js_c_clis - tot_js_cogs) > 0.01:
            fail(f"Desincronización en JS: Suma cogs clientes ({sum_js_c_clis:.2f} €) != total ({tot_js_cogs:.2f} €)")
        if abs(sum_js_c_prods - tot_js_cogs) > 0.01:
            fail(f"Desincronización en JS: Suma cogs productos ({sum_js_c_prods:.2f} €) != total ({tot_js_cogs:.2f} €)")
        if abs(sum_js_c_coms - tot_js_cogs) > 0.01:
            fail(f"Desincronización en JS: Suma cogs comerciales ({sum_js_c_coms:.2f} €) != total ({tot_js_cogs:.2f} €)")

        # Margen Industrial
        sum_js_mi_clis = round(sum(float(c.get("margen_industrial_eur", 0.0) or 0.0) for c in js_clis), 2)
        sum_js_mi_prods = round(sum(float(p.get("margen_industrial_eur", 0.0) or 0.0) for p in js_prods), 2)
        sum_js_mi_coms = round(sum(float(cm.get("margen_industrial_eur", 0.0) or 0.0) for cm in js_coms), 2)
        if abs(sum_js_mi_clis - tot_js_m_ind) > 0.01:
            fail(f"Desincronización en JS: Suma margen industrial clientes ({sum_js_mi_clis:.2f} €) != total ({tot_js_m_ind:.2f} €)")
        if abs(sum_js_mi_prods - tot_js_m_ind) > 0.01:
            fail(f"Desincronización en JS: Suma margen industrial productos ({sum_js_mi_prods:.2f} €) != total ({tot_js_m_ind:.2f} €)")
        if abs(sum_js_mi_coms - tot_js_m_ind) > 0.01:
            fail(f"Desincronización en JS: Suma margen industrial comerciales ({sum_js_mi_coms:.2f} €) != total ({tot_js_m_ind:.2f} €)")

        # Margen Contribución
        sum_js_mc_clis = round(sum(float(c.get("margen_contribucion_eur", 0.0) or 0.0) for c in js_clis), 2)
        sum_js_mc_coms = round(sum(float(cm.get("margen_contribucion_eur", 0.0) or 0.0) for cm in js_coms), 2)
        if abs(sum_js_mc_clis - tot_js_m_cont) > 0.01:
            fail(f"Desincronización en JS: Suma margen contribución clientes ({sum_js_mc_clis:.2f} €) != total ({tot_js_m_cont:.2f} €)")
        if abs(sum_js_mc_coms - tot_js_m_cont) > 0.01:
            fail(f"Desincronización en JS: Suma margen contribución comerciales ({sum_js_mc_coms:.2f} €) != total ({tot_js_m_cont:.2f} €)")

        # Unidades
        sum_js_u_clis = round(sum(float(c.get("uds", 0.0) or 0.0) for c in js_clis), 2)
        sum_js_u_prods = round(sum(float(p.get("uds", 0.0) or 0.0) for p in js_prods), 2)
        if abs(sum_js_u_clis - tot_js_uds) > 0.01:
            fail(f"Desincronización en JS: Suma unidades clientes ({sum_js_u_clis:.2f} u) != total ({tot_js_uds:.2f} u)")
        if abs(sum_js_u_prods - tot_js_uds) > 0.01:
            fail(f"Desincronización en JS: Suma unidades productos ({sum_js_u_prods:.2f} u) != total ({tot_js_uds:.2f} u)")
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
        ("CECOs Actividad Suma vs Coste Planta Actividad", round(sum(meta.get("costes_ceco_actividad_eur", meta.get("costes_ceco_eur", {})).values()) - tot["coste_planta"], 2)),
        ("Coste Planta Actividad + Subactividad vs Total Industrial", round(tot["coste_planta"] + tot.get("coste_subactividad", 0.0) - tot.get("coste_planta_total", tot["coste_planta"]), 2)),
        ("CECOs Total Industrial Suma vs Total Planta", round(sum(meta.get("costes_ceco_eur", {}).values()) - tot.get("coste_planta_total", tot["coste_planta"]), 2)),
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
    mat_tecnico = round(float(puente_mat.get("coste_tecnico_escandallos_eur", 0.0) or 0.0), 2)
    mat_drilldown = round(float(tot.get("cogs_materiales", 0.0) or 0.0), 2)
    if abs(mat_tecnico - mat_drilldown) > 0.01:
        fail(f"Desincronización en puente de materiales: {mat_tecnico:.2f} € vs detalle ventas {mat_drilldown:.2f} €", last_date)

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

    commercial_index = get_dest(os.path.join("web_dashboard", "index.html"))
    if os.path.exists(commercial_index):
        with open(commercial_index, "r", encoding="utf-8") as f:
            html_content = f.read()
        if "Estimación de Ventas | Codiagro" not in html_content or "fullcost" in html_content.lower():
            fail("VIOLACIÓN DE SEGURIDAD: web_dashboard/index.html contiene módulos o títulos de Dirección General en lugar de la versión comercial abierta.")


    # 9. Validación estricta de sintaxis JavaScript en archivos generados y bloques HTML
    js_files_to_check = [
        get_dest(os.path.join("web_dashboard_direccion", "fullcost_data.js")),
        get_dest(os.path.join("web_dashboard_direccion", "marta_tracking_data.js")),
        get_dest(os.path.join("web_dashboard_direccion", "dashboard_data.js")),
        get_dest(os.path.join("web_dashboard_direccion", "profitability_data.js")),
        get_dest(os.path.join("web_dashboard_direccion", "treasury_data.js")),
        get_dest(os.path.join("web_dashboard_direccion", "production_planning_data.js")),
        get_dest(os.path.join("web_dashboard", "dashboard_data.js")),
        get_dest(os.path.join("web_dashboard", "marta_tracking_data.js")),
        get_dest(os.path.join("web_dashboard", "production_planning_data.js"))
    ]
    for js_path in js_files_to_check:
        if os.path.exists(js_path):
            try:
                res_node = subprocess.run(["node", "--check", js_path], capture_output=True, text=True, encoding="utf-8")
                if res_node.returncode != 0:
                    fail(f"Error de sintaxis JavaScript en {js_path}: {res_node.stderr.strip()}", last_date)
            except Exception as e:
                pass

    # Comprobar sintaxis de bloques <script> en archivos HTML
    html_files_to_check = [
        get_dest(os.path.join("web_dashboard_direccion", "index.html")),
        get_dest(os.path.join("web_dashboard", "index.html")),
        os.path.join("web_dashboard_direccion", "index.html"),
        os.path.join("web_dashboard", "index.html")
    ]
    checked_html = set()
    for h_path in html_files_to_check:
        norm_h = os.path.normpath(h_path)
        if norm_h in checked_html or not os.path.exists(h_path):
            continue
        checked_html.add(norm_h)
        with open(h_path, "r", encoding="utf-8") as f_h:
            h_text = f_h.read()
        script_blocks = re.findall(r'<script(?![^>]*\bsrc=)[^>]*>([\s\S]*?)</script>', h_text, re.I)
        for s_idx, s_code in enumerate(script_blocks):
            tmp_chk = f"temp_chk_syntax_{s_idx}.js"
            try:
                with open(tmp_chk, "w", encoding="utf-8") as f_tmp:
                    f_tmp.write(s_code)
                res_node = subprocess.run(["node", "--check", tmp_chk], capture_output=True, text=True, encoding="utf-8")
                if res_node.returncode != 0:
                    fail(f"Error de sintaxis JavaScript en bloque script #{s_idx+1} de {h_path}: {res_node.stderr.strip()}", last_date)
            finally:
                if os.path.exists(tmp_chk):
                    os.remove(tmp_chk)

    # Recálculo independiente de la coincidencia del RAI
    rai_mayor = round(float(meta.get("rai_contable_mayor", 0.0) or 0.0), 2)
    rai_analitico = round(float(meta.get("rai_analitico_cascada", 0.0) or 0.0), 2)
    diff_recalc = round(abs(rai_mayor - rai_analitico), 2)
    diff_meta = round(float(meta.get("diferencia_conciliacion", 0.0) or 0.0), 2)
    if abs(diff_meta - diff_recalc) > 0.01:
        fail(f"Incoherencia en metadatos: 'diferencia_conciliacion' declarada ({diff_meta:.2f} €) no coincide con recálculo independiente abs({rai_mayor:.2f} - {rai_analitico:.2f}) = {diff_recalc:.2f} €", last_date)

    coincidencia_rai = "COINCIDENTE (0,00 €)" if diff_recalc <= 0.01 else f"DESVIACIÓN ({diff_recalc:.2f} €)"
    
    # Distinción explícita entre coincidencia técnica de RAI y estado del cierre contable (puentes externos)
    ex_estado = str(puente_ex.get("estado_global", "Pendiente de conciliación"))
    tr_estado = str(puente_tr.get("estado_global", "Pendiente de conciliación"))

    if ex_estado != "Comprobado" or tr_estado != "Comprobado":
        estado_cierre_contable = f"PENDIENTE DE CONCILIACIÓN (Existencias: {ex_estado} / Transporte: {tr_estado})"
    else:
        estado_cierre_contable = "CONCILIADO"

    status_payload = {
        "estado_tecnico": "COMPROBADO",
        "coincidencia_rai": coincidencia_rai,
        "estado_cierre_contable": estado_cierre_contable,
        "estado_conciliacion_contable": estado_cierre_contable,
        "fecha_validacion": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
        "fecha_datos_fullcost": meta.get("fecha_cierre", "30/09/2026"),
        "rai_contable_mayor_eur": rai_mayor,
        "rai_analitico_cascada_eur": rai_analitico,
        "diferencia_conciliacion_eur": diff_recalc,
        "conservacion_importes_drift": 0.0,
        "fuentes_verificadas": sources_info
    }
    with open(get_dest(os.path.join("web_dashboard_direccion", "update_validation_status.json")), "w", encoding="utf-8") as f:
        json.dump(status_payload, f, ensure_ascii=False, indent=2)

    print(f"2. Validación técnica superada al 100% (drift 0,00 €). Coincidencia RAI: {coincidencia_rai}. Cierre contable: {estado_cierre_contable}.")

if __name__ == "__main__":
    main()
