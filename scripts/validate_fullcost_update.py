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



# Conjunto explícito de entregables comerciales autorizados (allowlist estricta)



ALLOWED_COMMERCIAL_FILES = {



    "chart.umd.min.js",



    "codiagro_logo.png",



    "dashboard_data.js",



    "dashboard_data.json",



    "dashboard_data_BACKUP_PRE_REVISION_Q4.js",



    "dashboard_data_BACKUP_PRE_REVISION_Q4.json",



    "dashboard_i18n.js",



    "exceljs.min.js",



    "index.html",



    "marta_tracking.css",



    "marta_tracking_data.js",



    "marta_tracking_data.json",



    "marta_tracking_logic.js",



    "Orden_Fabricacion_MRP_Codiagro.xlsx",



    "production_planning_data.js",



    "production_planning_data.json",



    "server.py"



}







def extract_js_data(js_path):



    with open(js_path, "r", encoding="utf-8") as f:



        content = f.read().strip()



    idx_brace = -1



    for i, c in enumerate(content):



        if c in ['{', '[']:



            idx_brace = i



            break



    if idx_brace == -1:



        return None



    decoder = json.JSONDecoder()



    obj, _ = decoder.raw_decode(content[idx_brace:])



    return obj







def deep_compare_json_js(v1, v2, p=""):



    if isinstance(v1, (int, float)) and isinstance(v2, (int, float)):



        if abs(float(v1) - float(v2)) > 0.02:



            return f"{p}: JSON ({v1}) != JS ({v2})"



        return None



    if type(v1) != type(v2):



        return f"{p}: tipo JSON ({type(v1).__name__}) != tipo JS ({type(v2).__name__})"



    if isinstance(v1, dict):



        k1, k2 = set(v1.keys()), set(v2.keys())



        if k1 != k2:



            return f"{p}: claves discrepantes {k1 ^ k2}"



        for k in v1:



            err = deep_compare_json_js(v1[k], v2[k], f"{p}.{k}")



            if err: return err



    elif isinstance(v1, list):



        if len(v1) != len(v2):



            return f"{p}: longitud JSON ({len(v1)}) != longitud JS ({len(v2)})"



        for i, (i1, i2) in enumerate(zip(v1, v2)):



            err = deep_compare_json_js(i1, i2, f"{p}[{i}]")



            if err: return err



    else:



        if v1 != v2:



            return f"{p}: JSON ({v1}) != JS ({v2})"



    return None







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



    required_all_outputs = [



        get_dest(os.path.join("web_dashboard_direccion", "fullcost_data.json")),



        get_dest(os.path.join("web_dashboard_direccion", "fullcost_data.js")),



        get_dest(os.path.join("web_dashboard_direccion", "dashboard_data.json")),



        get_dest(os.path.join("web_dashboard_direccion", "dashboard_data.js")),



        get_dest(os.path.join("web_dashboard_direccion", "profitability_data.json")),



        get_dest(os.path.join("web_dashboard_direccion", "profitability_data.js")),



        get_dest(os.path.join("web_dashboard_direccion", "bi_intelligence_data.json")),



        get_dest(os.path.join("web_dashboard_direccion", "bi_intelligence_data.js")),



        get_dest(os.path.join("web_dashboard_direccion", "treasury_data.json")),



        get_dest(os.path.join("web_dashboard_direccion", "treasury_data.js")),



        get_dest(os.path.join("web_dashboard_direccion", "production_planning_data.json")),



        get_dest(os.path.join("web_dashboard_direccion", "production_planning_data.js")),



        get_dest(os.path.join("web_dashboard_direccion", "marta_tracking_data.json")),



        get_dest(os.path.join("web_dashboard_direccion", "marta_tracking_data.js")),



        get_dest(os.path.join("web_dashboard", "dashboard_data.json")),



        get_dest(os.path.join("web_dashboard", "dashboard_data.js")),



        get_dest(os.path.join("web_dashboard", "production_planning_data.json")),



        get_dest(os.path.join("web_dashboard", "production_planning_data.js")),



        get_dest(os.path.join("web_dashboard", "marta_tracking_data.json")),



        get_dest(os.path.join("web_dashboard", "marta_tracking_data.js")),



        get_dest("Seguimiento_Presupuesto_Sep_2026.xlsx"),



        get_dest("Orden_Fabricacion_MRP_Codiagro.xlsx")



    ]



    fc_json_path = get_dest(os.path.join("web_dashboard_direccion", "fullcost_data.json"))



    fc_js_path = get_dest(os.path.join("web_dashboard_direccion", "fullcost_data.js"))



    dash_json_path = get_dest(os.path.join("web_dashboard_direccion", "dashboard_data.json"))



    marta_json_path = get_dest(os.path.join("web_dashboard_direccion", "marta_tracking_data.json"))



    excel_seg_path = get_dest("Seguimiento_Presupuesto_Sep_2026.xlsx")



    excel_mrp_path = get_dest("Orden_Fabricacion_MRP_Codiagro.xlsx")







    for p in required_all_outputs:



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



        



        # Comparación exhaustiva de todas las magnitudes clave entre JSON y JavaScript (Punto 3 Auditoría 30/10)



        js_tot = fc_from_js.get("totales_drilldown", {})



        json_tot = fc.get("totales_drilldown", {})



        all_tot_keys = [



            "ventas", "ventas_reales_cta_700", "ventas_netas_grupo_7", "variacion_existencias_pt",



            "cogs_materiales", "coste_planta", "coste_subactividad", "coste_planta_total",



            "coste_stock_pt_no_vendido", "margen_industrial_eur", "margen_industrial_pct",



            "margen_industrial_tras_subactividad_eur", "margen_industrial_tras_subactividad_pct",



            "margen_industrial_neto_planta_eur", "margen_industrial_neto_planta_pct",



            "transporte_real", "margen_contribucion_eur", "margen_contribucion_pct", "uds"



        ]



        for k_tot in all_tot_keys:



            if k_tot in json_tot or k_tot in js_tot:



                v_js = round(float(js_tot.get(k_tot, 0.0) or 0.0), 2)



                v_json = round(float(json_tot.get(k_tot, 0.0) or 0.0), 2)



                if abs(v_js - v_json) > 0.01:



                    fail(f"Desincronización entre JSON y JS en totales_drilldown.{k_tot}: {v_json} € vs {v_js} €")







        # Comprobar identidades del margen directamente sobre los datos de JS que cargará la web



        js_m_ind = round(float(js_tot.get("margen_industrial_eur", 0.0)), 2)



        js_c_sub = round(float(js_tot.get("coste_subactividad", 0.0)), 2)



        js_c_stk = round(float(js_tot.get("coste_stock_pt_no_vendido", 0.0)), 2)



        js_m_tras = round(float(js_tot.get("margen_industrial_tras_subactividad_eur", js_m_ind - js_c_sub)), 2)



        js_m_net = round(float(js_tot.get("margen_industrial_neto_planta_eur", 0.0)), 2)



        if abs(js_m_tras - round(js_m_ind - js_c_sub, 2)) > 0.02:



            fail(f"Identidad violada en fullcost_data.js: margen tras subactividad ({js_m_tras:.2f} €) != {js_m_ind:.2f} € - {js_c_sub:.2f} €")



        if abs(js_m_net - round(js_m_tras - js_c_stk, 2)) > 0.02:



            fail(f"Identidad violada en fullcost_data.js: margen neto planta ({js_m_net:.2f} €) != {js_m_tras:.2f} € - {js_c_stk:.2f} €")







        # Paridad JS vs JSON en puente_facturacion con comprobación estricta de cuentas duplicadas y líneas (Punto 2 Auditoría)



        js_pf = fc_from_js.get("puente_facturacion", {})



        json_pf = fc.get("puente_facturacion", {})



        if not js_pf or not json_pf:



            fail("El bloque puente_facturacion falta en JSON o JS")



        for k_pf in ["valor_produccion_n1", "ventas_detalle_analitico_sep", "diferencia_produccion_vs_ventas_sep"]:



            v_js = round(float(js_pf.get(k_pf, 0.0) or 0.0), 2)



            v_json = round(float(json_pf.get(k_pf, 0.0) or 0.0), 2)



            if abs(v_js - v_json) > 0.01:



                fail(f"Desincronización entre JSON y JS en puente_facturacion.{k_pf}: {v_json} € vs {v_js} €")







        js_raw_lines = js_pf.get("lineas_conciliacion", [])



        json_raw_lines = json_pf.get("lineas_conciliacion", [])



        if len(js_raw_lines) != len(json_raw_lines):



            fail(f"Desincronización en número de líneas de puente_facturacion: JSON ({len(json_raw_lines)}) vs JS ({len(js_raw_lines)})")







        # Comprobar duplicados directamente en la lista cruda de JS antes de mapear



        seen_js_accounts = []



        for l in js_raw_lines:



            cta = l.get("cuenta")



            if cta in seen_js_accounts:



                fail(f"Cuenta contable duplicada en puente_facturacion de fullcost_data.js: {cta}")



            seen_js_accounts.append(cta)







        for idx, (l_json, l_js) in enumerate(zip(json_raw_lines, js_raw_lines)):



            if l_json.get("cuenta") != l_js.get("cuenta"):



                fail(f"Desincronización en línea #{idx+1} de puente_facturacion: cuenta JSON={l_json.get('cuenta')} vs JS={l_js.get('cuenta')}")



            val_json = round(float(l_json.get("importe_eur", 0.0) or 0.0), 2)



            val_js = round(float(l_js.get("importe_eur", 0.0) or 0.0), 2)



            if abs(val_json - val_js) > 0.01:



                fail(f"Desincronización en línea #{idx+1} ({l_json.get('cuenta')}) de puente_facturacion: JSON={val_json:.2f} € vs JS={val_js:.2f} €")







        # Comprobar sumas e identidades del puente directamente sobre fullcost_data.js



        js_acc_sum = round(sum(float(l.get("importe_eur", 0.0) or 0.0) for l in js_raw_lines if str(l.get("cuenta", "")).isdigit()), 2)



        js_n1 = round(float(js_pf.get("valor_produccion_n1", 0.0) or 0.0), 2)



        if abs(js_acc_sum - js_n1) > 0.01:



            fail(f"Incoherencia en fullcost_data.js: suma de cuentas del puente ({js_acc_sum:.2f} €) != valor_produccion_n1 ({js_n1:.2f} €)")



        js_perim_lines = [l for l in js_raw_lines if l.get("cuenta") == "Ajuste Perímetro"]



        if not js_perim_lines:



            fail("Falta la línea obligatoria 'Ajuste Perímetro' en puente_facturacion de fullcost_data.js")



        js_perim = round(float(js_perim_lines[0].get("importe_eur", 0.0) or 0.0), 2)



        js_detail = round(float(js_pf.get("ventas_detalle_analitico_sep", 0.0) or js_tot.get("ventas", 0.0) or 0.0), 2)



        if abs(round(js_n1 + js_perim, 2) - js_detail) > 0.01:



            fail(f"Incoherencia en fullcost_data.js: cuadre de producción + deducción perímetro ({js_n1 + js_perim:.2f} €) != detalle SKU ({js_detail:.2f} €)")







        # 1. Comprobación profunda exhaustiva (deep comparison) de fullcost_data JSON vs JS



        err_fc_deep = deep_compare_json_js(fc, fc_from_js, "fullcost_data")



        if err_fc_deep:



            fail(f"Desincronización profunda entre fullcost_data.json y fullcost_data.js: {err_fc_deep}")







        # 2. Recálculo y verificación de subtotales de P&L Waterfall en fullcost_data.js y JSON



        for ds_name, pnl_to_check in [("JS", fc_from_js.get("pnl_waterfall", [])), ("JSON", fc.get("pnl_waterfall", []))]:



            if not pnl_to_check or not isinstance(pnl_to_check, list) or len(pnl_to_check) == 0:



                fail(f"El bloque pnl_waterfall en {ds_name} está vacío o ausente")



            



            run_tot = 0.0



            final_rai_val = None



            for item in pnl_to_check:



                nv = str(item.get("nivel", "")).strip().upper()



                tp = str(item.get("tipo", "")).strip().lower()



                imp = round(float(item.get("importe", 0.0) or 0.0), 2)



                if tp in ["ingreso", "coste"]:



                    run_tot = round(run_tot + imp, 2)



                elif tp in ["subtotal", "ebit"]:



                    if abs(imp - run_tot) > 0.05:



                        fail(f"Subtotal incorrecto en cascada P&L ({ds_name}) nivel {nv}: valor reportado {imp:.2f} € != calculado acumulado {run_tot:.2f} €")



                    if nv == "FINAL" or tp == "ebit":



                        final_rai_val = imp



            



            if final_rai_val is None:



                fail(f"Falta el escalón final de RAI/EBIT en pnl_waterfall de {ds_name}")



            expected_meta_rai = round(float(fc_from_js.get("metadata", {}).get("rai_analitico_cascada", 0.0)), 2)



            if abs(final_rai_val - expected_meta_rai) > 0.01:



                fail(f"El RAI final de pnl_waterfall en {ds_name} ({final_rai_val:.2f} €) no coincide con metadata.rai_analitico_cascada ({expected_meta_rai:.2f} €)")







        # 3. Verificación de identidades en cuentas_resumen de fullcost_data.js y JSON



        for ds_name, cres_to_check in [("JS", fc_from_js.get("cuentas_resumen", {})), ("JSON", fc.get("cuentas_resumen", {}))]:



            if not cres_to_check or not isinstance(cres_to_check, dict) or not cres_to_check.get("ingresos_totales"):



                fail(f"El bloque cuentas_resumen en {ds_name} está vacío o ausente")



            cr_rai = round(float(cres_to_check.get("rai_neto", 0.0) or 0.0), 2)



            cr_ing = round(float(cres_to_check.get("ingresos_totales", 0.0) or 0.0), 2)



            cr_gas = round(float(cres_to_check.get("gastos_totales", 0.0) or 0.0), 2)



            if abs(round(cr_ing - cr_gas, 2) - cr_rai) > 0.02:



                fail(f"Incoherencia en cuentas_resumen ({ds_name}): ingresos ({cr_ing:.2f} €) - gastos ({cr_gas:.2f} €) != rai_neto ({cr_rai:.2f} €)")



            expected_mayor_rai = round(float(fc_from_js.get("metadata", {}).get("rai_contable_mayor", 0.0)), 2)



            if abs(cr_rai - expected_mayor_rai) > 0.01:



                fail(f"El rai_neto de cuentas_resumen en {ds_name} ({cr_rai:.2f} €) no coincide con metadata.rai_contable_mayor ({expected_mayor_rai:.2f} €)")







        # 4. Verificación de identidades entre costes, horas, tarifas y capacidad en capacidades



        for ds_name, caps_to_check in [("JS", fc_from_js.get("capacidades", [])), ("JSON", fc.get("capacidades", []))]:



            if not caps_to_check or not isinstance(caps_to_check, list) or len(caps_to_check) == 0:



                fail(f"El bloque capacidades en {ds_name} está vacío o ausente")



            sum_c_cecos = 0.0



            sum_c_sub = 0.0



            sum_c_act = 0.0



            for cap in caps_to_check:



                ceco = str(cap.get("ceco", ""))



                c_tot_ceco = round(float(cap.get("coste_total_ceco", 0.0) or 0.0), 2)



                c_sub_ceco = round(float(cap.get("coste_subactividad_eur", 0.0) or 0.0), 2)



                c_act_ceco = round(float(cap.get("coste_actividad_eur", 0.0) or 0.0), 2)



                h_nom = float(cap.get("horas_nominales", 0.0) or 0.0)



                coste_hora = round(float(cap.get("coste_hora", 0.0) or 0.0), 2)



                tarifa_eur_h = round(float(cap.get("tarifa_eur_h", 0.0) or 0.0), 2)



                tarifa_efectiva_hora = round(float(cap.get("tarifa_efectiva_hora", 0.0) or 0.0), 2)







                # Exigir paridad estricta entre todos los campos equivalentes de tarifa



                if abs(coste_hora - tarifa_eur_h) > 0.01:



                    fail(f"Discrepancia interna en tarifa de CECO {ceco} ({ds_name}): coste_hora ({coste_hora:.2f} €/h) != tarifa_eur_h ({tarifa_eur_h:.2f} €/h)")



                if abs(coste_hora - tarifa_efectiva_hora) > 0.01:



                    fail(f"Discrepancia interna en tarifa de CECO {ceco} ({ds_name}): coste_hora ({coste_hora:.2f} €/h) != tarifa_efectiva_hora ({tarifa_efectiva_hora:.2f} €/h)")







                if abs(round(c_sub_ceco + c_act_ceco, 2) - c_tot_ceco) > 0.05:



                    fail(f"Incoherencia en CECO {ceco} ({ds_name}): coste_total ({c_tot_ceco:.2f} €) != subactividad ({c_sub_ceco:.2f} €) + actividad ({c_act_ceco:.2f} €)")



                if h_nom > 0:



                    tarifa_calc = round(c_tot_ceco / h_nom, 2)



                    if abs(tarifa_eur_h - tarifa_calc) > 0.05:



                        fail(f"Tarifa horaria violada en CECO {ceco} ({ds_name}): tarifa ({tarifa_eur_h:.2f} €/h) != calculada ({c_tot_ceco:.2f} € / {h_nom:.2f} h = {tarifa_calc:.2f} €/h)")



                sum_c_cecos = round(sum_c_cecos + c_tot_ceco, 2)



                sum_c_sub = round(sum_c_sub + c_sub_ceco, 2)



                sum_c_act = round(sum_c_act + c_act_ceco, 2)







            exp_planta_tot = round(float(js_tot.get("coste_planta_total", 0.0) or 0.0), 2)



            if abs(sum_c_cecos - exp_planta_tot) > 0.05:



                fail(f"Suma de coste planta en capacidades {ds_name} ({sum_c_cecos:.2f} €) != totales_drilldown.coste_planta_total ({exp_planta_tot:.2f} €)")







        # 5. Verificación de identidades económicas completas de Maquila en JS y JSON



        for ds_name, maq_to_check in [("JS", fc_from_js.get("maquila", {})), ("JSON", fc.get("maquila", {}))]:



            if not maq_to_check or not isinstance(maq_to_check, dict):



                fail(f"El bloque maquila en {ds_name} está vacío o ausente")







            # 5.1 Desglose de costes absorbidos



            m_ett = round(float(maq_to_check.get("costes_absorbidos", {}).get("mod_envasado_ett", 0.0) or 0.0), 2)



            m_elec = round(float(maq_to_check.get("costes_absorbidos", {}).get("electricidad_proceso", 0.0) or 0.0), 2)



            m_res = round(float(maq_to_check.get("costes_absorbidos", {}).get("gestion_residuos", 0.0) or 0.0), 2)



            m_c_tot = round(float(maq_to_check.get("coste_total", 0.0) or 0.0), 2)



            if abs(round(m_ett + m_elec + m_res, 2) - m_c_tot) > 0.02:



                fail(f"Coste total de maquila ({m_c_tot:.2f} €) != suma de componentes ({m_ett + m_elec + m_res:.2f} €) en {ds_name}")







            # 5.2 Comprobación dinámica de ETT maquila (sin constante fija)



            h_maq = float(maq_to_check.get("horas_envasado_ett", 0.0) or 0.0)



            h_tot_rec = float(maq_to_check.get("horas_totales_recurso", 0.0) or 0.0)



            coste_fact_rec = float(maq_to_check.get("coste_factura_ett_recurso", 0.0) or 0.0)



            if h_tot_rec > 0 and coste_fact_rec > 0:



                ett_calc = round(coste_fact_rec * (h_maq / h_tot_rec), 2)



                if abs(m_ett - ett_calc) > 0.02:



                    fail(f"Coste ETT maquila ({m_ett:.2f} €) != cálculo dinámico {coste_fact_rec:.2f} € × ({h_maq:.1f} / {h_tot_rec:.1f} h) = {ett_calc:.2f} € en {ds_name}")







            # 5.3 Ingresos ordinarios



            c705 = round(float(maq_to_check.get("cuenta_705_servicios", 0.0) or maq_to_check.get("ingreso_servicios_705", 0.0) or 0.0), 2)



            c703 = round(float(maq_to_check.get("cuenta_703_accesorios", 0.0) or maq_to_check.get("ingreso_accesorios_703", 0.0) or 0.0), 2)



            ing_ord = round(float(maq_to_check.get("ingreso_ordinario_total", 0.0) or 0.0), 2)



            if abs(round(c705 + c703, 2) - ing_ord) > 0.02:



                fail(f"Ingreso ordinario de maquila ({ing_ord:.2f} €) != Cta 705 ({c705:.2f} €) + Cta 703 ({c703:.2f} €) en {ds_name}")







            # 5.4 Margen ordinario



            m_ord = round(float(maq_to_check.get("margen_ordinario_eur", 0.0) or 0.0), 2)



            if abs(round(ing_ord - m_c_tot, 2) - m_ord) > 0.02:



                fail(f"Margen ordinario de maquila ({m_ord:.2f} €) != ingreso ({ing_ord:.2f} €) - coste ({m_c_tot:.2f} €) en {ds_name}")







            # 5.5 Ingreso neto y resultado neto



            c702 = round(float(maq_to_check.get("cuenta_702_ajuste_agosto", 0.0) or 0.0), 2)



            ing_net = round(float(maq_to_check.get("ingreso_neto_contable_sep", 0.0) or 0.0), 2)



            if abs(round(ing_ord + c702, 2) - ing_net) > 0.02:



                fail(f"Ingreso neto de maquila ({ing_net:.2f} €) != ordinario ({ing_ord:.2f} €) + ajuste 702 ({c702:.2f} €) en {ds_name}")







            res_net = round(float(maq_to_check.get("resultado_neto_sep", 0.0) or maq_to_check.get("resultado_neto", 0.0) or 0.0), 2)



            if abs(round(ing_net - m_c_tot, 2) - res_net) > 0.02:



                fail(f"Resultado neto de maquila ({res_net:.2f} €) != ingreso neto ({ing_net:.2f} €) - coste ({m_c_tot:.2f} €) en {ds_name}")











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







        import math



        # Validación registro a registro de finitud (sin NaN ni Inf) e identidades de cálculo



        for c in js_clis:



            for f_num in ["ventas", "cogs_materiales", "coste_planta_fullcost", "margen_industrial_eur", "transporte_real", "margen_contribucion_eur", "uds"]:



                v_f = float(c.get(f_num, 0.0) or 0.0)



                if not math.isfinite(v_f):



                    fail(f"Número no finito ({v_f}) en cliente {c.get('cliente')} campo {f_num}")



            c_calc_m = round(float(c.get("ventas", 0.0)) - float(c.get("cogs_materiales", 0.0)) - float(c.get("coste_planta_fullcost", 0.0)), 2)



            if abs(c_calc_m - float(c.get("margen_industrial_eur", 0.0))) > 0.02:



                fail(f"Identidad de margen industrial violada en cliente {c.get('cliente')}: {c.get('margen_industrial_eur')} != {c_calc_m}")







        for p in js_prods:



            for f_num in ["ventas", "cogs_materiales", "coste_planta_fullcost", "margen_industrial_eur", "uds"]:



                v_f = float(p.get(f_num, 0.0) or 0.0)



                if not math.isfinite(v_f):



                    fail(f"Número no finito ({v_f}) en producto SKU {p.get('sku')} campo {f_num}")



            p_calc_m = round(float(p.get("ventas", 0.0)) - float(p.get("cogs_materiales", 0.0)) - float(p.get("coste_planta_fullcost", 0.0)), 2)



            if abs(p_calc_m - float(p.get("margen_industrial_eur", 0.0))) > 0.02:



                fail(f"Identidad de margen industrial violada en producto SKU {p.get('sku')}: {p.get('margen_industrial_eur')} != {p_calc_m}")







        for cm in js_coms:



            for f_num in ["ventas", "cogs_materiales", "coste_planta_fullcost", "margen_industrial_eur", "portes", "margen_contribucion_eur"]:



                v_f = float(cm.get(f_num, 0.0) or 0.0)



                if not math.isfinite(v_f):



                    fail(f"Número no finito ({v_f}) en comercial {cm.get('comercial')} campo {f_num}")



            cm_calc_m = round(float(cm.get("ventas", 0.0)) - float(cm.get("cogs_materiales", 0.0)) - float(cm.get("coste_planta_fullcost", 0.0)), 2)



            if abs(cm_calc_m - float(cm.get("margen_industrial_eur", 0.0))) > 0.02:



                fail(f"Identidad de margen industrial violada en comercial {cm.get('comercial')}: {cm.get('margen_industrial_eur')} != {cm_calc_m}")







        # Comprobación de identidad registro a registro: Margen Contribución = Margen Industrial - Transporte



        for c in js_clis:



            c_calc_mc = round(float(c.get("margen_industrial_eur", 0.0)) - float(c.get("transporte_real", 0.0)), 2)



            if abs(c_calc_mc - float(c.get("margen_contribucion_eur", 0.0))) > 0.02:



                fail(f"Identidad de margen de contribución violada en cliente {c.get('cliente')}: {c.get('margen_contribucion_eur')} != {c_calc_mc}")







        for cm in js_coms:



            cm_calc_mc = round(float(cm.get("margen_industrial_eur", 0.0)) - float(cm.get("portes", 0.0)), 2)



            if abs(cm_calc_mc - float(cm.get("margen_contribucion_eur", 0.0))) > 0.02:



                fail(f"Identidad de margen de contribución violada en comercial {cm.get('comercial')}: {cm.get('margen_contribucion_eur')} != {cm_calc_mc}")







        # Paridad exhaustiva cuenta a cuenta JS vs JSON en puente_transporte (Punto 6 Auditoría)



        js_ptr = fc_from_js.get("puente_transporte", {})



        json_ptr = fc.get("puente_transporte", {})



        if abs(float(js_ptr.get("reporte_logistico_imputado_eur", 0.0)) - float(json_ptr.get("reporte_logistico_imputado_eur", 0.0))) > 0.01:



            fail("Desincronización JS vs JSON en puente_transporte.reporte_logistico_imputado_eur")







        js_desglose = {d.get("cuenta"): d for d in js_ptr.get("desglose_por_cuenta", [])}



        json_desglose = {d.get("cuenta"): d for d in json_ptr.get("desglose_por_cuenta", [])}



        if set(js_desglose.keys()) != set(json_desglose.keys()):



            fail("Desincronización en cuentas de transporte entre JSON y JS")



        for cta_k, json_item in json_desglose.items():



            js_item = js_desglose[cta_k]



            for f_tte in ["reporte_logistico_eur", "mayor_contable_eur", "diferencia_eur", "expediciones_valoradas"]:



                v_js = js_item.get(f_tte, 0)



                v_json = json_item.get(f_tte, 0)



                if abs(float(v_js) - float(v_json)) > 0.01:



                    fail(f"Desincronización JS vs JSON en transporte cuenta {cta_k} campo {f_tte}: {v_json} vs {v_js}")







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



        ("CECOs Actividad Suma vs Coste Actividad Fabril Total", round(sum(meta.get("costes_ceco_actividad_eur", meta.get("costes_ceco_eur", {})).values()) - (tot.get("coste_actividad_planta_total") or tot["coste_planta"]), 2)),



        ("Coste Planta Actividad + Subactividad vs Total Industrial", round((tot.get("coste_actividad_planta_total") or tot["coste_planta"]) + tot.get("coste_subactividad", 0.0) - tot.get("coste_planta_total", tot["coste_planta"]), 2)),



        ("Coste Planta Ventas + Stock No Vendido vs Actividad Total", round(tot["coste_planta"] + tot.get("coste_stock_pt_no_vendido", 0.0) - (tot.get("coste_actividad_planta_total") or tot["coste_planta"]), 2)),



        ("CECOs Total Industrial Suma vs Total Planta", round(sum(meta.get("costes_ceco_eur", {}).values()) - tot.get("coste_planta_total", tot["coste_planta"]), 2)),



    ]







    for name, diff in checks:



        if abs(diff) > 0.0:



            fail(f"Descuadre de redondeo en Full Costing ({name}): diferencia = {diff}", last_date)







    # Validar coherencia del Radar de Prioridades frente a Marta Tracking recién generado
    bi_json_path = get_dest(os.path.join("web_dashboard_direccion", "bi_intelligence_data.json"))
    if os.path.exists(bi_json_path) and os.path.exists(marta_json_path):
        try:
            with open(bi_json_path, "r", encoding="utf-8") as f:
                bi_val = json.load(f)
            with open(marta_json_path, "r", encoding="utf-8") as f:
                marta_val = json.load(f)

            m_clientes = marta_val.get("euros", {}).get("clientes", [])
            val_churn_list = []
            for c in m_clientes:
                name = c.get("cliente", "")
                if "DOGA" in name.upper():
                    continue
                v25 = float(c.get("ventas_2025_sep", 0) or 0)
                v26 = float(c.get("ventas_2026_sep", 0) or 0)
                if v25 > 5000 and v26 < v25 * 0.75:
                    val_churn_list.append(round(v25 - v26, 2))
            val_churn_list.sort(reverse=True)
            val_top15 = val_churn_list[:15] if len(val_churn_list) >= 15 else val_churn_list
            val_impacto_recalculado = -round(sum(val_top15), 2)
            val_cuentas_recalculadas = len(val_top15)

            radar_p1 = bi_val.get("radar_prioridades", {}).get("prioridad_1_churn", {})
            radar_impacto = float(radar_p1.get("impacto_eur", 0.0))
            radar_cuentas = int(radar_p1.get("cuentas_count", 0))

            if abs(radar_impacto - val_impacto_recalculado) > 0.05 or radar_cuentas != val_cuentas_recalculadas:
                fail(f"Desincronización crítica en Radar de Prioridades vs Marta Tracking: Radar={radar_impacto} € ({radar_cuentas} ctas), Recalculado={val_impacto_recalculado} € ({val_cuentas_recalculadas} ctas). Ambos módulos deben consumir la misma ejecución.")
        except Exception as e:
            fail(f"Error al validar coherencia Radar vs Marta: {e}")

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







    # 6.1 Validar Puente de Facturación y cuentas contables dinámicas (Puntos 4, 6 y 7 Auditoría 30/10)



    puente_fac = fc.get("puente_facturacion", {})



    if not puente_fac or not puente_fac.get("lineas_conciliacion"):



        fail("El puente de conciliación de facturación está vacío o carece de líneas de cuentas justificativas.", last_date)







    pf_lines = puente_fac.get("lineas_conciliacion", [])



    cd = fc.get("cuentas_detalle", {})







    # Comprobar cobertura exhaustiva: toda cuenta del Mayor Grupo 7, 71 y 778 debe estar representada exactamente una vez



    g7_expected_accounts = [c for c in sorted(cd.keys()) if (c.startswith('70') or c.startswith('71') or c.startswith('778'))]



    seen_accounts = set()



    account_lines_sum = 0.0







    for l in pf_lines:



        cta = l.get("cuenta")



        if cta and cta.isdigit():



            if cta in seen_accounts:



                fail(f"Cuenta contable duplicada en puente_facturacion: {cta}", last_date)



            seen_accounts.add(cta)



            if cta not in cd:



                fail(f"Cuenta {cta} presente en puente_facturacion pero ausente en cuentas_detalle", last_date)



            exp_val = round(float(cd[cta].get("saldo_mes", 0.0) or 0.0), 2)



            act_val = round(float(l.get("importe_eur", 0.0) or 0.0), 2)



            if abs(exp_val - act_val) > 0.01:



                fail(f"Línea de puente_facturacion desincronizada para cuenta {cta}: en Mayor={exp_val:.2f} € vs puente={act_val:.2f} €", last_date)



            account_lines_sum += act_val







    missing_ctas = set(g7_expected_accounts) - seen_accounts



    if missing_ctas:



        fail(f"Cuentas del Mayor no cubiertas en puente_facturacion ({len(missing_ctas)} cuentas): {sorted(missing_ctas)}", last_date)







    account_lines_sum = round(account_lines_sum, 2)



    valor_n1 = round(float(puente_fac.get("valor_produccion_n1", 0.0)), 2)



    if abs(account_lines_sum - valor_n1) > 0.01:



        fail(f"La suma de cuentas del puente_facturacion ({account_lines_sum:.2f} €) no coincide con valor_produccion_n1 ({valor_n1:.2f} €)", last_date)







    # Validar línea de deducción de perímetro y cuadre exacto con detalle analítico SKU



    perim_lines = [l for l in pf_lines if l.get("cuenta") == "Ajuste Perímetro"]



    if not perim_lines:



        fail("Falta la línea obligatoria de 'Ajuste Perímetro' en puente_facturacion", last_date)



    perim_val = round(float(perim_lines[0].get("importe_eur", 0.0)), 2)



    diff_perim_declared = round(float(puente_fac.get("diferencia_produccion_vs_ventas_sep", 0.0)), 2)



    if abs(perim_val - (-diff_perim_declared)) > 0.01:



        fail(f"La línea 'Ajuste Perímetro' debe aplicarse como deducción negativa (-{diff_perim_declared:.2f} €), valor encontrado: {perim_val:.2f} €", last_date)







    tot_detail = round(float(puente_fac.get("ventas_detalle_analitico_sep", 0.0)), 2)



    if abs(round(valor_n1 + perim_val, 2) - tot_detail) > 0.01:



        fail(f"Cuadre aritmético fallido en puente_facturacion: {valor_n1:.2f} € + ({perim_val:.2f} €) != {tot_detail:.2f} €", last_date)







    # 6.2 Validar identidades contables dinámicas (no estáticas)



    if cd:



        cta_700_exp = round(cd.get("700000000", {}).get("saldo_mes", 0.0), 2)



        cta_700_pub = round(float(tot.get("ventas_reales_cta_700", 0.0)), 2)



        if abs(cta_700_exp - cta_700_pub) > 0.01:



            fail(f"Cifra contable Cta 700 desincronizada: en cuentas_detalle={cta_700_exp:.2f} € vs totales_drilldown={cta_700_pub:.2f} €", last_date)







        g7_exp = round(sum(v.get("saldo_mes", 0.0) for c, v in cd.items() if c.startswith("70")), 2)



        g7_pub = round(float(tot.get("ventas_netas_grupo_7", 0.0)), 2)



        if abs(g7_exp - g7_pub) > 0.01:



            fail(f"Cifra contable ventas netas Grupo 7 desincronizada: cuentas={g7_exp:.2f} € vs totales={g7_pub:.2f} €", last_date)







        var_pt_exp = round(sum(v.get("saldo_mes", 0.0) for c, v in cd.items() if c.startswith("71")), 2)



        var_pt_pub = round(float(tot.get("variacion_existencias_pt", 0.0)), 2)



        if abs(var_pt_exp - var_pt_pub) > 0.01:



            fail(f"Cifra contable variación existencias PT desincronizada: cuentas={var_pt_exp:.2f} € vs totales={var_pt_pub:.2f} €", last_date)







    # 6.3 Validar identidades de los dos márgenes industriales y ajuste de existencias (Puntos 4 y 5 Auditoría)



    m_ind_base = round(float(tot.get("margen_industrial_eur", 0.0)), 2)



    c_subact = round(float(tot.get("coste_subactividad", 0.0)), 2)



    c_stock_nv = round(float(tot.get("coste_stock_pt_no_vendido", 0.0)), 2)



    m_tras_subact = round(float(tot.get("margen_industrial_tras_subactividad_eur", m_ind_base - c_subact)), 2)



    m_neto_planta = round(float(tot.get("margen_industrial_neto_planta_eur", 0.0)), 2)







    if abs(m_tras_subact - round(m_ind_base - c_subact, 2)) > 0.02:



        fail(f"Incoherencia en margen industrial tras subactividad: {m_tras_subact:.2f} € != {m_ind_base:.2f} € - {c_subact:.2f} €", last_date)







    if abs(m_neto_planta - round(m_tras_subact - c_stock_nv, 2)) > 0.02:



        fail(f"Incoherencia en margen industrial neto planta: {m_neto_planta:.2f} € != {m_tras_subact:.2f} € - {c_stock_nv:.2f} €", last_date)







    c_planta_tot = round(float(tot.get("coste_planta_total", 0.0)), 2)



    v_netas = round(float(tot.get("ventas", 0.0)), 2)



    cogs_mat = round(float(tot.get("cogs_materiales", 0.0)), 2)



    if abs(m_neto_planta - round(v_netas - cogs_mat - c_planta_tot, 2)) > 0.02:



        fail(f"Identidad global de margen neto planta violada: {m_neto_planta:.2f} € != {v_netas:.2f} € - {cogs_mat:.2f} € - {c_planta_tot:.2f} €", last_date)







    # 6.4 Validar obligatoriedad, coherencia y cuadre estricto de antigüedad de deuda (Auditoría 30/10)



    bi_path = get_dest(os.path.join("web_dashboard_direccion", "bi_intelligence_data.json"))



    bi_js_path = get_dest(os.path.join("web_dashboard_direccion", "bi_intelligence_data.js"))



    if not os.path.exists(bi_path):



        fail("El archivo obligatorio bi_intelligence_data.json no existe", last_date)







    try:



        with open(bi_path, "r", encoding="utf-8") as f_bi:
            bi_data = json.load(f_bi)

        # Validación estricta de conservación de acuerdos comerciales frente a la fuente
        fuente_acuerdos_path = get_source(os.path.join("Gastos", "acuerdos_comerciales_registrados.json"))
        fuente_acuerdos = []
        if os.path.exists(fuente_acuerdos_path):
            try:
                with open(fuente_acuerdos_path, "r", encoding="utf-8") as f_ac:
                    raw_ac = f_ac.read()
                if not raw_ac.strip():
                    fail(f"El archivo fuente de acuerdos {fuente_acuerdos_path} existe pero está vacío (0 bytes o solo espacios).", last_date)
                fuente_acuerdos = json.loads(raw_ac)
                if not isinstance(fuente_acuerdos, list):
                    fail(f"El archivo fuente de acuerdos {fuente_acuerdos_path} no contiene una lista JSON válida.", last_date)
            except Exception as e:
                fail(f"El archivo fuente de acuerdos {fuente_acuerdos_path} está dañado o no es legible: {e}", last_date)

        def canonical_acuerdo_val(item):
            if not isinstance(item, dict):
                return {}
            return {
                k: (str(v).strip() if isinstance(v, str) else v)
                for k, v in sorted(item.items())
                if not k.startswith('_')
            }

        mr_json = bi_data.get("mejora_10_modo_reunion", {})
        acuerdos_bi_json = mr_json.get("acuerdos_registrados", [])
        if not isinstance(acuerdos_bi_json, list):
            fail("El campo acuerdos_registrados en bi_intelligence_data.json no es una lista válida.", last_date)

        # Cargar y verificar también bi_intelligence_data.js
        acuerdos_bi_js = []
        if not os.path.exists(bi_js_path):
            fail(f"El archivo obligatorio {bi_js_path} no existe.", last_date)
        try:
            with open(bi_js_path, "r", encoding="utf-8") as f_js:
                js_raw = f_js.read()
            prefix = "window.BI_INTELLIGENCE_DATA ="
            p_idx = js_raw.find(prefix)
            if p_idx == -1:
                fail("Estructura inválida en bi_intelligence_data.js: falta window.BI_INTELLIGENCE_DATA.", last_date)
            js_json_str = js_raw[p_idx + len(prefix):].rstrip()
            if js_json_str.endswith(";"):
                js_json_str = js_json_str[:-1].rstrip()
            bi_js_data = json.loads(js_json_str)
            acuerdos_bi_js = bi_js_data.get("mejora_10_modo_reunion", {}).get("acuerdos_registrados", [])
            if not isinstance(acuerdos_bi_js, list):
                fail("El campo acuerdos_registrados en bi_intelligence_data.js no es una lista válida.", last_date)
        except Exception as e:
            fail(f"Error al analizar bi_intelligence_data.js: {e}", last_date)

        # 1. Caso fuente vacía ([]): debe exigir estrictamente cero acuerdos generados
        if len(fuente_acuerdos) == 0:
            if len(acuerdos_bi_json) > 0:
                fail(f"La fuente central de acuerdos está vacía ([]), pero bi_intelligence_data.json contiene {len(acuerdos_bi_json)} acuerdos generados no autorizados.", last_date)
            if len(acuerdos_bi_js) > 0:
                fail(f"La fuente central de acuerdos está vacía ([]), pero bi_intelligence_data.js contiene {len(acuerdos_bi_js)} acuerdos generados no autorizados.", last_date)

        # 2. Comprobación bidireccional estricta de identificadores y recuentos
        fuente_ids = {str(a.get("id")) for a in fuente_acuerdos if isinstance(a, dict) and a.get("id")}
        json_ids = {str(a.get("id")) for a in acuerdos_bi_json if isinstance(a, dict) and a.get("id")}
        js_ids = {str(a.get("id")) for a in acuerdos_bi_js if isinstance(a, dict) and a.get("id")}

        if len(acuerdos_bi_json) != len(fuente_acuerdos):
            fail(f"Discrepancia en recuento de acuerdos en bi_intelligence_data.json: fuente={len(fuente_acuerdos)}, generado={len(acuerdos_bi_json)}.", last_date)
        if len(acuerdos_bi_js) != len(fuente_acuerdos):
            fail(f"Discrepancia en recuento de acuerdos en bi_intelligence_data.js: fuente={len(fuente_acuerdos)}, generado={len(acuerdos_bi_js)}.", last_date)

        if fuente_ids != json_ids:
            extra_json = json_ids - fuente_ids
            miss_json = fuente_ids - json_ids
            if extra_json:
                fail(f"Registros adicionales no autorizados en bi_intelligence_data.json: {extra_json}", last_date)
            if miss_json:
                fail(f"Pérdida de acuerdos en bi_intelligence_data.json: faltan {miss_json}", last_date)

        if fuente_ids != js_ids:
            extra_js = js_ids - fuente_ids
            miss_js = fuente_ids - js_ids
            if extra_js:
                fail(f"Registros adicionales no autorizados en bi_intelligence_data.js: {extra_js}", last_date)
            if miss_js:
                fail(f"Pérdida de acuerdos en bi_intelligence_data.js: faltan {miss_js}", last_date)

        # 3. Comprobación canónica de contenido por acuerdo
        for a_src in fuente_acuerdos:
            aid = str(a_src.get("id"))
            c_src = canonical_acuerdo_val(a_src)
            # Buscar en JSON
            matching_json = next((a for a in acuerdos_bi_json if str(a.get("id")) == aid), None)
            if matching_json is None:
                fail(f"Acuerdo {aid} ausente en bi_intelligence_data.json.", last_date)
            c_json = canonical_acuerdo_val(matching_json)
            if c_src != c_json:
                fail(f"Alteración de contenido detectada en acuerdo {aid} en bi_intelligence_data.json frente a la fuente validada.", last_date)

            # Buscar en JS
            matching_js = next((a for a in acuerdos_bi_js if str(a.get("id")) == aid), None)
            if matching_js is None:
                fail(f"Acuerdo {aid} ausente en bi_intelligence_data.js.", last_date)
            c_js = canonical_acuerdo_val(matching_js)
            if c_src != c_js:
                fail(f"Alteración de contenido detectada en acuerdo {aid} en bi_intelligence_data.js frente a la fuente validada.", last_date)




    except Exception as e:



        fail(f"Error al parsear bi_intelligence_data.json: {e}", last_date)







    deuda = bi_data.get("antiguedad_deuda")



    if not deuda or not isinstance(deuda, dict):



        fail("El bloque obligatorio 'antiguedad_deuda' no existe o está vacío en bi_intelligence_data.json", last_date)







    # Exigir la presencia obligatoria de ambas estructuras 'clientes' y 'detalle_clientes'



    if "clientes" not in deuda or not isinstance(deuda["clientes"], list) or len(deuda["clientes"]) == 0:



        fail("Falta la lista obligatoria 'clientes' en antiguedad_deuda requerida por la interfaz web", last_date)



    if "detalle_clientes" not in deuda or not isinstance(deuda["detalle_clientes"], list) or len(deuda["detalle_clientes"]) == 0:



        fail("Falta la lista obligatoria 'detalle_clientes' en antiguedad_deuda", last_date)



    if len(deuda["clientes"]) != len(deuda["detalle_clientes"]):



        fail(f"Desincronización en número de elementos entre 'clientes' ({len(deuda['clientes'])}) y 'detalle_clientes' ({len(deuda['detalle_clientes'])})", last_date)







    tot_deuda = round(float(deuda.get("total_deuda_eur", 0.0) or deuda.get("total_deuda_gestionada_eur", 0.0)), 2)



    if tot_deuda <= 0.0:



        fail(f"Total de deuda no válido o nulo: {tot_deuda} €", last_date)







    # Comprobar igualdad estricta entre total_deuda_eur y total_deuda_gestionada_eur si ambos existen



    if "total_deuda_eur" in deuda and "total_deuda_gestionada_eur" in deuda:



        if abs(float(deuda["total_deuda_eur"]) - float(deuda["total_deuda_gestionada_eur"])) > 0.01:



            fail(f"Divergencia entre total_deuda_eur ({deuda['total_deuda_eur']}) y total_deuda_gestionada_eur ({deuda['total_deuda_gestionada_eur']})", last_date)







    tramos = deuda.get("total_por_tramo_eur") or deuda.get("totales_tramos")



    if not tramos or not isinstance(tramos, dict):



        fail("El desglose de tramos de deuda 'total_por_tramo_eur' no existe en bi_intelligence_data.json", last_date)







    def norm_t(k):



        k = k.lower().replace("t", "").strip()



        if "plus" in k or ">" in k or "121" in k: return "121_plus"



        if ("90" in k and "120" in k) or "91" in k: return "91_120"



        if ("60" in k and "90" in k) or "61" in k: return "61_90"



        if ("30" in k and "60" in k) or "31" in k: return "31_60"



        if "<" in k or "0_30" in k or "0-30" in k or k == "0": return "0_30"



        return k







    tramos_clean = {norm_t(k): round(float(v), 2) for k, v in tramos.items() if k != "total"}



    sum_tramos = round(sum(tramos_clean.values()), 2)



    if abs(tot_deuda - sum_tramos) > 0.05:



        fail(f"Descuadre en panel de deuda: total deuda ({tot_deuda:.2f} €) != suma tramos ({sum_tramos:.2f} €)", last_date)







    # Auditar cada cliente y comprobar correspondencia exacta entre campos de pantalla y campos validados



    clientes_deuda = deuda["detalle_clientes"]



    clientes_pantalla = deuda["clientes"]







    for idx, (cli_val, cli_ui) in enumerate(zip(clientes_deuda, clientes_pantalla)):



        c_name = cli_val.get("cliente", f"Cliente #{idx+1}")



        if cli_val.get("cliente") != cli_ui.get("cliente"):



            fail(f"Desincronización en orden/nombre de cliente #{idx+1}: detalle_clientes={cli_val.get('cliente')} vs clientes={cli_ui.get('cliente')}", last_date)







        # Comprobar igualdad estricta entre campos de la lista clientes y de detalle_clientes



        for k_t, k_n in [("t0_30", "0_30"), ("t31_60", "31_60"), ("t61_90", "61_90"), ("t91_120", "91_120"), ("t121_plus", "121_plus")]:



            v_ui = round(float(cli_ui.get(k_t, 0.0) or 0.0), 2)



            v_val_flat = round(float(cli_val.get(k_t, 0.0) or 0.0), 2)



            v_val_nested = round(float(cli_val.get("tramos_eur", {}).get(k_n, 0.0) or 0.0), 2)



            if abs(v_ui - v_val_nested) > 0.01:



                fail(f"Divergencia en cliente {c_name}: campo de pantalla clientes.{k_t} ({v_ui:.2f} €) != detalle_clientes.tramos_eur.{k_n} ({v_val_nested:.2f} €)", last_date)



            if abs(v_val_flat - v_val_nested) > 0.01:



                fail(f"Divergencia interna en cliente {c_name}: {k_t} ({v_val_flat:.2f} €) != {k_n} ({v_val_nested:.2f} €)", last_date)







        tot_ui = round(float(cli_ui.get("total_vencido", 0.0) or cli_ui.get("saldo_eur", 0.0) or 0.0), 2)



        tot_val = round(float(cli_val.get("saldo_eur", 0.0) or cli_val.get("total_vencido", 0.0) or 0.0), 2)



        if abs(tot_ui - tot_val) > 0.01:



            fail(f"Divergencia en total de cliente {c_name}: clientes ({tot_ui:.2f} €) != detalle_clientes ({tot_val:.2f} €)", last_date)







        sum_c_tramos = round(sum(float(cli_val.get("tramos_eur", {}).get(tk, 0.0) or 0.0) for tk in ["0_30", "31_60", "61_90", "91_120", "121_plus"]), 2)



        if abs(tot_val - sum_c_tramos) > 0.05:



            fail(f"Descuadre en cliente {c_name}: saldo ({tot_val:.2f} €) != suma tramos ({sum_c_tramos:.2f} €)", last_date)







    sum_sal_clis = round(sum(float(c.get("saldo_eur", 0.0) or 0.0) for c in clientes_deuda), 2)



    if abs(sum_sal_clis - tot_deuda) > 0.05:



        fail(f"Descuadre total clientes de deuda ({sum_sal_clis:.2f} €) != total gestionado ({tot_deuda:.2f} €)", last_date)







    for tk in ["0_30", "31_60", "61_90", "91_120", "121_plus"]:



        val_tramo_g = tramos_clean.get(tk, 0.0)



        tot_exp_r = round(sum(float(cli.get("tramos_eur", {}).get(tk, 0.0) or 0.0) for cli in clientes_deuda), 2)



        if abs(val_tramo_g - tot_exp_r) > 0.05:



            fail(f"Descuadre en tramo {tk}: suma clientes ({tot_exp_r:.2f} €) != total tramo ({val_tramo_g:.2f} €)", last_date)







    # Paridad exhaustiva cliente por cliente y tramo por tramo entre bi_intelligence_data.json y bi_intelligence_data.js



    if not os.path.exists(bi_js_path):



        fail("El archivo obligatorio bi_intelligence_data.js no existe", last_date)







    try:



        with open(bi_js_path, "r", encoding="utf-8") as f_bjs:



            raw_bjs = f_bjs.read().strip()



        prefix_bjs = "window.BI_INTELLIGENCE_DATA ="



        if not raw_bjs.startswith(prefix_bjs):



            fail("El archivo bi_intelligence_data.js no contiene la cabecera 'window.BI_INTELLIGENCE_DATA ='", last_date)



        raw_bjs_clean = raw_bjs[len(prefix_bjs):].strip()



        if raw_bjs_clean.endswith(";"): raw_bjs_clean = raw_bjs_clean[:-1].strip()



        bi_from_js = json.loads(raw_bjs_clean)



        deuda_js = bi_from_js.get("antiguedad_deuda")



        if not deuda_js or not isinstance(deuda_js, dict):



            fail("El bloque 'antiguedad_deuda' falta o está vacío en bi_intelligence_data.js", last_date)







        tot_deuda_js = round(float(deuda_js.get("total_deuda_eur", 0.0) or deuda_js.get("total_deuda_gestionada_eur", 0.0)), 2)



        if abs(tot_deuda - tot_deuda_js) > 0.01:



            fail(f"Desincronización en deuda total entre JSON ({tot_deuda:.2f} €) y JS ({tot_deuda_js:.2f} €)", last_date)







        js_tramos_raw = deuda_js.get("total_por_tramo_eur") or deuda_js.get("totales_tramos") or {}



        js_tramos_clean = {norm_t(k): round(float(v), 2) for k, v in js_tramos_raw.items() if k != "total"}



        for tk in ["0_30", "31_60", "61_90", "91_120", "121_plus"]:



            v_json_tr = tramos_clean.get(tk, 0.0)



            v_js_tr = js_tramos_clean.get(tk, 0.0)



            if abs(v_json_tr - v_js_tr) > 0.01:



                fail(f"Desincronización en tramo {tk} de deuda entre JSON ({v_json_tr:.2f} €) y JS ({v_js_tr:.2f} €)", last_date)







        # Comprobar exhaustivamente ambas listas (clientes y detalle_clientes) y todos los tramos en JS



        for list_name in ["clientes", "detalle_clientes"]:



            js_list = deuda_js.get(list_name)



            if not js_list or not isinstance(js_list, list):



                fail(f"Falta la lista obligatoria '{list_name}' en bi_intelligence_data.js", last_date)



            if len(js_list) != len(clientes_pantalla):



                fail(f"Desincronización en número de clientes en lista '{list_name}' de JS: JSON ({len(clientes_pantalla)}) vs JS ({len(js_list)})", last_date)







            js_cli_map = {str(c.get("cliente", "")).strip().upper(): c for c in js_list}



            for cli_json in clientes_pantalla:



                c_name_norm = str(cli_json.get("cliente", "")).strip().upper()



                c_js = js_cli_map.get(c_name_norm)



                if not c_js:



                    fail(f"Cliente de deuda '{cli_json.get('cliente')}' presente en JSON pero ausente en lista '{list_name}' de JS", last_date)



                for fld in ["saldo_eur", "total_vencido", "t0_30", "t31_60", "t61_90", "t91_120", "t121_plus"]:



                    v_j = round(float(cli_json.get(fld, 0.0) or 0.0), 2)



                    v_s = round(float(c_js.get(fld, 0.0) or 0.0), 2)



                    if abs(v_j - v_s) > 0.01:



                        fail(f"Desincronización en cliente de deuda '{cli_json.get('cliente')}' campo '{fld}' en '{list_name}' de JS: JSON={v_j:.2f} € vs JS={v_s:.2f} €", last_date)



                for tk in ["0_30", "31_60", "61_90", "91_120", "121_plus"]:



                    v_j_tr = round(float(cli_json.get("tramos_eur", {}).get(tk, 0.0) or 0.0), 2)



                    v_s_tr = round(float(c_js.get("tramos_eur", {}).get(tk, 0.0) or 0.0), 2)



                    if abs(v_j_tr - v_s_tr) > 0.01:



                        fail(f"Desincronización en cliente de deuda '{cli_json.get('cliente')}' tramos_eur.{tk} en '{list_name}' de JS: JSON={v_j_tr:.2f} € vs JS={v_s_tr:.2f} €", last_date)







        # Comprobar igualdad estricta entre clientes y detalle_clientes dentro de JS



        js_ui_clis = deuda_js.get("clientes", [])



        js_det_clis = deuda_js.get("detalle_clientes", [])



        for idx, (c_ui, c_det) in enumerate(zip(js_ui_clis, js_det_clis)):



            c_name = c_ui.get("cliente", f"Cliente #{idx+1}")



            for k_t, k_n in [("t0_30", "0_30"), ("t31_60", "31_60"), ("t61_90", "61_90"), ("t91_120", "91_120"), ("t121_plus", "121_plus")]:



                v_ui = round(float(c_ui.get(k_t, 0.0) or 0.0), 2)



                v_det = round(float(c_det.get("tramos_eur", {}).get(k_n, 0.0) or 0.0), 2)



                if abs(v_ui - v_det) > 0.01:



                    fail(f"Divergencia dentro de bi_intelligence_data.js en cliente {c_name}: clientes.{k_t} ({v_ui:.2f} €) != detalle_clientes.{k_n} ({v_det:.2f} €)", last_date)



                v_det_flat = round(float(c_det.get(k_t, 0.0) or 0.0), 2)



                if abs(v_ui - v_det_flat) > 0.01:



                    fail(f"Divergencia dentro de bi_intelligence_data.js en cliente {c_name}: clientes.{k_t} ({v_ui:.2f} €) != detalle_clientes.{k_t} ({v_det_flat:.2f} €)", last_date)



            tot_ui = round(float(c_ui.get("total_vencido", 0.0) or 0.0), 2)



            tot_det = round(float(c_det.get("total_vencido", 0.0) or c_det.get("saldo_eur", 0.0) or 0.0), 2)



            if abs(tot_ui - tot_det) > 0.01:



                fail(f"Divergencia dentro de bi_intelligence_data.js en cliente {c_name}: total_vencido={tot_ui:.2f} € vs {tot_det:.2f} €", last_date)



    except Exception as e:



        fail(f"Error al validar paridad exhaustiva de deuda en bi_intelligence_data.js: {e}", last_date)











    # 6.5 Validación integral de Rentabilidad (profitability_data.json y .js)



    prof_json_p = get_dest(os.path.join("web_dashboard_direccion", "profitability_data.json"))



    prof_js_p = get_dest(os.path.join("web_dashboard_direccion", "profitability_data.js"))



    if not os.path.exists(prof_json_p) or not os.path.exists(prof_js_p):



        fail("Archivos obligatorios profitability_data.json o .js ausentes", last_date)



    try:



        with open(prof_json_p, "r", encoding="utf-8") as f:



            prof_json = json.load(f)



        prof_js = extract_js_data(prof_js_p)



        if not prof_js or not isinstance(prof_js, dict) or len(prof_js.keys()) == 0:



            fail("El archivo profitability_data.js contiene un objeto vacío o carece de contenido", last_date)



        for req_k in ["totales_empresa", "clientes", "productos"]:



            if req_k not in prof_js or not prof_js[req_k]:



                fail(f"El archivo profitability_data.js carece del bloque obligatorio '{req_k}' o está vacío", last_date)



            if req_k not in prof_json or not prof_json[req_k]:



                fail(f"El archivo profitability_data.json carece del bloque obligatorio '{req_k}' o está vacío", last_date)



        



        prof_tot = prof_js.get("totales_empresa", {})



        v27 = round(float(prof_tot.get("ventas_2027", 0.0) or 0.0), 2)



        c27 = round(float(prof_tot.get("cogs_2027", 0.0) or 0.0), 2)



        m27 = round(float(prof_tot.get("margen_2027_eur", 0.0) or 0.0), 2)



        if v27 <= 0 or c27 <= 0:



            fail("profitability_data.js contiene ventas o COGS no válidos o cero en totales_empresa", last_date)



        if abs(round(v27 - c27, 2) - m27) > 0.02:



            fail(f"Identidad de margen bruto violada en profitability_data.js: ventas ({v27:.2f} €) - cogs ({c27:.2f} €) != margen ({m27:.2f} €)", last_date)



        



        err_prof = deep_compare_json_js(prof_json, prof_js, "profitability_data")



        if err_prof:



            fail(f"Desincronización profunda entre profitability_data.json y .js: {err_prof}", last_date)



    except Exception as e:



        fail(f"Error al validar profitability_data: {e}", last_date)







    # 6.6 Comprobación profunda exhaustiva (deep comparison) de todos los demás pares JSON / JS



    dataset_pairs_to_audit = [



        ("treasury_data", get_dest(os.path.join("web_dashboard_direccion", "treasury_data.json")), get_dest(os.path.join("web_dashboard_direccion", "treasury_data.js"))),



        ("marta_tracking_data_dir", get_dest(os.path.join("web_dashboard_direccion", "marta_tracking_data.json")), get_dest(os.path.join("web_dashboard_direccion", "marta_tracking_data.js"))),



        ("dashboard_data_dir", get_dest(os.path.join("web_dashboard_direccion", "dashboard_data.json")), get_dest(os.path.join("web_dashboard_direccion", "dashboard_data.js"))),



        ("production_planning_data_dir", get_dest(os.path.join("web_dashboard_direccion", "production_planning_data.json")), get_dest(os.path.join("web_dashboard_direccion", "production_planning_data.js"))),



        ("dashboard_data_com", get_dest(os.path.join("web_dashboard", "dashboard_data.json")), get_dest(os.path.join("web_dashboard", "dashboard_data.js"))),



        ("marta_tracking_data_com", get_dest(os.path.join("web_dashboard", "marta_tracking_data.json")), get_dest(os.path.join("web_dashboard", "marta_tracking_data.js"))),



        ("production_planning_data_com", get_dest(os.path.join("web_dashboard", "production_planning_data.json")), get_dest(os.path.join("web_dashboard", "production_planning_data.js")))



    ]



    for ds_lbl, j_path, js_path in dataset_pairs_to_audit:



        if os.path.exists(j_path) and os.path.exists(js_path):



            try:



                with open(j_path, "r", encoding="utf-8") as f:



                    j_obj = json.load(f)



                js_obj = extract_js_data(js_path)



                if not js_obj or (isinstance(js_obj, dict) and len(js_obj) == 0):



                    fail(f"El archivo {ds_lbl} en JavaScript está vacío o no contiene datos válidos", last_date)



                err_pair = deep_compare_json_js(j_obj, js_obj, ds_lbl)



                if err_pair:



                    fail(f"Desincronización profunda en dataset {ds_lbl}: {err_pair}", last_date)



            except Exception as e:



                fail(f"Error al auditar paridad en {ds_lbl}: {e}", last_date)







    # 7. Validar Excel de seguimiento



    try:



        wb_seg = openpyxl.load_workbook(excel_seg_path, data_only=True)



        if "Resumen General" not in wb_seg.sheetnames:



            fail("El archivo Seguimiento_Presupuesto_Sep_2026.xlsx no contiene la hoja 'Resumen General'", last_date)



        wb_seg.close()



    except Exception as e:



        fail(f"No se pudo abrir o validar Seguimiento_Presupuesto_Sep_2026.xlsx: {e}", last_date)







    # 8. Escudo de seguridad: Verificar aislamiento estricto de web_dashboard mediante ALLOWLIST explícita



    # Se audita tanto el staging temporal (si está activo) como el directorio de producción



    dirs_comm_to_check = []



    if os.environ.get("CODIAGRO_STAGING_DIR"):



        staging_comm = os.path.join(os.environ["CODIAGRO_STAGING_DIR"], "web_dashboard")



        if os.path.exists(staging_comm):



            dirs_comm_to_check.append(staging_comm)



    if os.path.exists("web_dashboard"):



        dirs_comm_to_check.append("web_dashboard")







    for c_dir in dirs_comm_to_check:



        for fn in os.listdir(c_dir):



            if fn not in ALLOWED_COMMERCIAL_FILES:



                fail(f"VIOLACIÓN DE SEGURIDAD: Archivo no autorizado para la versión comercial detectado en '{c_dir}': {fn}. Sólo se permiten los 17 entregables comerciales autorizados.")







    commercial_index = get_dest(os.path.join("web_dashboard", "index.html"))



    if os.path.exists(commercial_index):



        with open(commercial_index, "r", encoding="utf-8") as f:



            html_content = f.read()



        if "Estimación de Ventas | Codiagro" not in html_content or "fullcost" in html_content.lower():



            fail("VIOLACIÓN DE SEGURIDAD: web_dashboard/index.html contiene módulos o títulos de Dirección General en lugar de la versión comercial abierta.")

        for forbidden_html_tok in ["stock_suggestions", "precio_unitario", "no_mover_eur", "slow_mover_eur"]:
            if forbidden_html_tok in html_content:
                fail(f"VIOLACIÓN DE SEGURIDAD: web_dashboard/index.html contiene referencia prohibida '{forbidden_html_tok}'.")

    # 8B. Validación de aislamiento estricto de datos en el paquete comercial (web_dashboard/production_planning_data.*)
    comm_prod_json = get_dest(os.path.join("web_dashboard", "production_planning_data.json"))
    comm_prod_js = get_dest(os.path.join("web_dashboard", "production_planning_data.js"))
    FORBIDDEN_COMMERCIAL_KEYS = {
        "stock_suggestions", "alternativas_economicas", "ofertas_clientes",
        "campanas_ofertas", "simulador_resumen", "balance_empresa", "hipotesis_calculo",
        "inventario_total_base_eur", "clientes_balance_eur", "proveedores_balance_eur",
        "cogs_ltm_total_eur", "margen_neto_eur", "margen_bruto_eur", "margen_contable_eur",
        "beneficio_incremental_eur", "ganancia_neta_eur", "coste_fin_eur", "coste_unit_eur",
        "cogs_eur", "precio_unitario", "no_mover_eur", "slow_mover_eur", "stock_eur",
        "regalo_coste_eur", "holding_cost_saved_eur", "ingreso_neto_eur", "ingreso_bruto_eur",
        "neto_caja_eur"
    }

    def scan_forbidden_commercial(obj, path="root"):
        if isinstance(obj, dict):
            for k, v in obj.items():
                if k in FORBIDDEN_COMMERCIAL_KEYS:
                    fail(f"VIOLACIÓN DE SEGURIDAD COMERCIAL: clave financiera/interna '{k}' expuesta en {path}.{k}")
                scan_forbidden_commercial(v, f"{path}.{k}")
        elif isinstance(obj, list):
            for idx, item in enumerate(obj):
                scan_forbidden_commercial(item, f"{path}[{idx}]")

    if os.path.exists(comm_prod_json):
        with open(comm_prod_json, "r", encoding="utf-8") as f_cj:
            comm_payload = json.load(f_cj)
        scan_forbidden_commercial(comm_payload, "web_dashboard/production_planning_data.json")
        if os.path.exists(comm_prod_js):
            with open(comm_prod_js, "r", encoding="utf-8") as f_cjs:
                raw_cjs = f_cjs.read().strip()
            if not raw_cjs.startswith("window.PRODUCTION_PLANNING_DATA = "):
                fail("web_dashboard/production_planning_data.js no comienza con window.PRODUCTION_PLANNING_DATA = ")
            js_payload_str = raw_cjs[len("window.PRODUCTION_PLANNING_DATA = "):].rstrip(";")
            scan_forbidden_commercial(json.loads(js_payload_str), "web_dashboard/production_planning_data.js")












    # 9. Validación estricta de sintaxis JavaScript en archivos generados y bloques HTML (Punto 9 Auditoría)



    # Incluye exhaustivamente todos los scripts JS generados y existentes (como dashboard_i18n.js)



    js_files_to_check = [



        get_dest(os.path.join("web_dashboard_direccion", "fullcost_data.js")),



        get_dest(os.path.join("web_dashboard_direccion", "marta_tracking_data.js")),



        get_dest(os.path.join("web_dashboard_direccion", "dashboard_data.js")),



        get_dest(os.path.join("web_dashboard_direccion", "profitability_data.js")),



        get_dest(os.path.join("web_dashboard_direccion", "bi_intelligence_data.js")),



        get_dest(os.path.join("web_dashboard_direccion", "treasury_data.js")),



        get_dest(os.path.join("web_dashboard_direccion", "production_planning_data.js")),



        get_dest(os.path.join("web_dashboard_direccion", "dashboard_i18n.js")),



        get_dest(os.path.join("web_dashboard", "dashboard_data.js")),



        get_dest(os.path.join("web_dashboard", "marta_tracking_data.js")),



        get_dest(os.path.join("web_dashboard", "production_planning_data.js")),



        get_dest(os.path.join("web_dashboard", "dashboard_i18n.js")),



        os.path.join("web_dashboard_direccion", "profitability_logic.js"),



        os.path.join("web_dashboard_direccion", "cockpit_inicio_logic.js"),



        os.path.join("web_dashboard_direccion", "bi_intelligence_logic.js"),



        os.path.join("web_dashboard_direccion", "dashboard_i18n.js"),



        os.path.join("web_dashboard", "dashboard_i18n.js")



    ]



    # Búsqueda dinámica de cualquier otro .js en las carpetas web



    for d_web in ["web_dashboard_direccion", "web_dashboard"]:



        w_dir = get_dest(d_web)



        if os.path.exists(w_dir):



            for f_name in os.listdir(w_dir):



                if f_name.endswith(".js") and not f_name.endswith(".min.js"):



                    full_p = os.path.join(w_dir, f_name)



                    if full_p not in js_files_to_check:



                        js_files_to_check.append(full_p)



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

    # 9B. Validación independiente económica, física, ACCA y de frescura de la Suite Comercial de Stock
    dir_prod_json = get_dest(os.path.join("web_dashboard_direccion", "production_planning_data.json"))
    if not os.path.exists(dir_prod_json):
        fail("Falta web_dashboard_direccion/production_planning_data.json para auditoría de Suite Comercial de Stock.", last_date)

    with open(dir_prod_json, "r", encoding="utf-8") as f_dp:
        dir_prod = json.load(f_dp)

    import math as _math
    import glob as _glob
    from datetime import datetime as _dt

    def _parse_campaign_src_date(fname):
        base_f = os.path.basename(fname)
        if base_f.startswith("~$"):
            return None
        m_full = re.search(r"(\d{2})[._-](\d{2})[._-](\d{4})", base_f)
        if m_full:
            d, m, y = int(m_full.group(1)), int(m_full.group(2)), int(m_full.group(3))
            try:
                return _dt(y, m, d)
            except ValueError:
                return None
        m_dm = re.search(r"(\d{1,2})\.(\d{1,2})(?:\.(\d{2})(?!\d))?", base_f)
        if m_dm:
            d, m = int(m_dm.group(1)), int(m_dm.group(2))
            y = (2000 + int(m_dm.group(3))) if m_dm.group(3) else (2026 if m >= 9 else 2027)
            try:
                return _dt(y, m, d)
            except ValueError:
                return None
        return None

    def _req_finite(val, ctx):
        if isinstance(val, bool) or not isinstance(val, (int, float)) or not _math.isfinite(float(val)):
            fail(f"INTEGRIDAD NUMÉRICA STOCK: valor no finito o inválido ({val!r}) en {ctx}", last_date)
        return float(val)

    def _valid_dmy(s):
        if not isinstance(s, str):
            return False
        m = re.match(r"^(\d{2})/(\d{2})/(\d{4})$", s.strip())
        if not m:
            return False
        d, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            _dt(y, mo, d)
            return True
        except ValueError:
            return False

    latest_stock_disk = max([d for d in (_parse_campaign_src_date(f) for f in _glob.glob(get_source("Stock*.xlsx"))) if d is not None], default=None)
    latest_orders_disk = max([d for d in (_parse_campaign_src_date(f) for f in _glob.glob(get_source("Pedidos *.xlsx"))) if d is not None], default=None)
    meta = dir_prod.get("metadata") or {}
    if latest_stock_disk and meta.get("stock_source_date"):
        if meta["stock_source_date"] != latest_stock_disk.strftime("%Y-%m-%d"):
            fail(f"FRESCURA FUENTES: stock_source_date={meta['stock_source_date']} difiere del Stock más reciente ({latest_stock_disk.strftime('%Y-%m-%d')}).", last_date)
    if latest_orders_disk and meta.get("orders_source_date"):
        if meta["orders_source_date"] != latest_orders_disk.strftime("%Y-%m-%d"):
            fail(f"FRESCURA FUENTES: orders_source_date={meta['orders_source_date']} difiere del Pedidos más reciente ({latest_orders_disk.strftime('%Y-%m-%d')}).", last_date)

    scenarios_dict = dir_prod.get("scenarios")
    if not isinstance(scenarios_dict, dict) or "revisado" not in scenarios_dict or "oficial" not in scenarios_dict:
        fail("INTEGRIDAD STOCK: faltan escenarios 'revisado' y/u 'oficial' en production_planning_data.json.", last_date)

    for scen_name in ("revisado", "oficial"):
        sc_data = scenarios_dict[scen_name]
        ss = sc_data.get("stock_suggestions")
        if not isinstance(ss, dict):
            fail(f"INTEGRIDAD STOCK [{scen_name}]: falta bloque stock_suggestions.", last_date)

        hip = ss.get("hipotesis_calculo")
        if not isinstance(hip, dict) or not (hip.get("estado_aprobacion") or hip.get("estado_validacion")):
            fail(f"TRANSPARENCIA [{scen_name}]: falta metadata hipotesis_calculo/estado_aprobacion.", last_date)

        base_skus = sc_data.get("base_skus") or []
        sku_map = {str(s.get("base_sku", "")).strip(): s for s in base_skus if s.get("base_sku")}
        gift_stock_map = ss.get("gift_stock_map") or {}
        items = ss.get("items") or []
        kpis = ss.get("kpis") or {}
        plans = ss.get("plans") or []

        if not items:
            fail(f"INTEGRIDAD STOCK [{scen_name}]: items vacío.", last_date)

        plan_acc = {}
        sum_it_vol = 0.0
        sum_it_exceso_eur = 0.0
        sum_it_ing_neto = 0.0
        sum_it_margen = 0.0
        sum_it_ahorro = 0.0
        sum_it_fin = 0.0

        for it in items:
            sku = str(it.get("base_sku", "")).strip()
            ctx_it = f"{scen_name}.items[{sku}]"
            st_fis = _req_finite(it.get("stock_actual", 0), f"{ctx_it}.stock_actual")
            ped_p = _req_finite(it.get("pedidos_pendientes", 0), f"{ctx_it}.pedidos_pendientes")
            st_lib = _req_finite(it.get("stock_libre_disponible", 0), f"{ctx_it}.stock_libre_disponible")
            if st_fis < -0.01 or st_lib < -0.01 or st_lib > st_fis + 0.05:
                fail(f"INCOHERENCIA FÍSICA [{ctx_it}]: stock_libre_disponible={st_lib} vs stock_actual={st_fis}.", last_date)
            if abs(st_lib - max(0.0, st_fis - ped_p)) > 0.15:
                fail(f"INCOHERENCIA STOCK LIBRE [{ctx_it}]: stock_libre_disponible={st_lib} vs físico-pedidos={max(0.0, st_fis - ped_p):.2f}.", last_date)

            ex_qty = _req_finite(it.get("exceso_total_qty", 0), f"{ctx_it}.exceso_total_qty")
            ex_eur = _req_finite(it.get("exceso_total_eur", 0), f"{ctx_it}.exceso_total_eur")
            it_ing = _req_finite(it.get("ingreso_neto_eur", 0), f"{ctx_it}.ingreso_neto_eur")
            it_mrg = _req_finite(it.get("margen_neto_eur", 0), f"{ctx_it}.margen_neto_eur")
            it_aho = _req_finite(it.get("holding_cost_saved_eur", 0), f"{ctx_it}.holding_cost_saved_eur")
            it_fin = _req_finite(it.get("coste_fin_pago_eur", 0), f"{ctx_it}.coste_fin_pago_eur")
            it_cost_u = _req_finite(it.get("coste_unit", 0), f"{ctx_it}.coste_unit")
            exact_cost_u = _req_finite((sku_map.get(sku) or {}).get("precio_unitario", it_cost_u), f"{ctx_it}.exact_cost_unit")
            if abs(it_cost_u - exact_cost_u) > 0.015:
                fail(f"COSTE UNITARIO DESINCRONIZADO [{ctx_it}]: coste_unit={it_cost_u} vs base_sku={exact_cost_u}.", last_date)

            if ex_qty < -0.01 or ex_qty > st_lib + 0.15:
                fail(f"EXCESO SOBRE STOCK LIBRE [{ctx_it}]: exceso_total_qty={ex_qty} supera stock_libre_disponible={st_lib}.", last_date)

            sum_it_vol += ex_qty
            sum_it_exceso_eur += ex_eur
            sum_it_ing_neto += it_ing
            sum_it_margen += it_mrg
            sum_it_ahorro += it_aho
            sum_it_fin += it_fin

            offers = it.get("ofertas_clientes") or []
            if not offers:
                fail(f"INTEGRIDAD STOCK [{ctx_it}]: SKU sin ofertas_clientes.", last_date)

            sum_o_vol = 0.0
            sum_o_ing = 0.0
            sum_o_mrg = 0.0
            sum_o_aho = 0.0
            sum_o_fin = 0.0
            env_cap = _req_finite(it.get("envase_l", 0), f"{ctx_it}.envase_l")

            for o in offers:
                oid = str(o.get("id_oferta", "")).strip()
                ctx_o = f"{ctx_it}.oferta[{oid}]"
                vol = _req_finite(o.get("vol_asignado", 0), f"{ctx_o}.vol_asignado")
                if vol <= 0 or vol > st_lib + 0.05 or vol > st_fis + 0.05:
                    fail(f"RESERVA IMPOSIBLE [{ctx_o}]: vol_asignado={vol} frente a stock_libre={st_lib}, stock_actual={st_fis}.", last_date)

                if not _valid_dmy(o.get("fecha_limite")):
                    fail(f"FECHA INVÁLIDA [{ctx_o}]: fecha_limite={o.get('fecha_limite')!r}.", last_date)

                is_rework = bool(o.get("es_reprocesado_interno"))
                if not is_rework and env_cap > 0 and vol >= env_cap - 1e-6:
                    env_n = round(vol / env_cap)
                    if abs(vol - env_n * env_cap) > 0.05:
                        fail(f"ENVASE FRACCIONADO [{ctx_o}]: vol_asignado={vol} no es múltiplo de envase_l={env_cap}.", last_date)

                pvp_ref = _req_finite(o.get("pvp_tarifa_ref", 0), f"{ctx_o}.pvp_tarifa_ref")
                pvp_of = _req_finite(o.get("pvp_oferta", 0), f"{ctx_o}.pvp_oferta")
                dto_pct = _req_finite(o.get("dto_pct", 0), f"{ctx_o}.dto_pct")
                fact_n = _req_finite(o.get("facturacion_neta_eur", 0), f"{ctx_o}.facturacion_neta_eur")
                c_rap = _req_finite(o.get("coste_rappel_eur", 0), f"{ctx_o}.coste_rappel_eur")
                c_fin = _req_finite(o.get("coste_fin_eur", 0), f"{ctx_o}.coste_fin_eur")
                c_mrc = _req_finite(o.get("coste_merch_eur", 0), f"{ctx_o}.coste_merch_eur")
                c_reg = _req_finite(o.get("regalo_coste_eur", 0), f"{ctx_o}.regalo_coste_eur")
                cogs = _req_finite(o.get("cogs_eur", 0), f"{ctx_o}.cogs_eur")
                mrg_est = _req_finite(o.get("margen_estimado_eur", 0), f"{ctx_o}.margen_estimado_eur")
                hold_s = _req_finite(o.get("holding_cost_saved_eur", 0), f"{ctx_o}.holding_cost_saved_eur")

                if c_reg > c_mrc + 0.05:
                    fail(f"REGALO/MERCH INCONSISTENTE [{ctx_o}]: regalo_coste_eur={c_reg} > coste_merch_eur={c_mrc}.", last_date)

                if not is_rework:
                    expected_mrg = fact_n - cogs - c_rap - c_mrc - c_fin
                    if abs(mrg_est - expected_mrg) > 0.25:
                        fail(f"ARITMÉTICA MARGEN [{ctx_o}]: margen_estimado_eur={mrg_est} != fact-cogs-rap-merch-fin={expected_mrg:.2f}.", last_date)
                    if abs(fact_n - (vol * pvp_of)) > max(0.5, 0.002 * abs(fact_n)):
                        fail(f"ARITMÉTICA FACTURACIÓN [{ctx_o}]: facturacion_neta_eur={fact_n} != vol*pvp={vol * pvp_of:.2f}.", last_date)
                    if pvp_ref > 0 and abs(dto_pct - ((1.0 - (pvp_of / pvp_ref)) * 100.0)) > 0.25:
                        fail(f"DESCUENTO INCOHERENTE [{ctx_o}]: dto_pct={dto_pct} vs real={((1.0 - (pvp_of / pvp_ref)) * 100.0):.2f}%.", last_date)
                    if abs(cogs - (vol * exact_cost_u)) > max(0.5, 0.002 * abs(cogs)):
                        fail(f"COGS INCOHERENTE [{ctx_o}]: cogs_eur={cogs} != vol*coste_unit={vol * exact_cost_u:.2f}.", last_date)
                else:
                    mp_sav = _req_finite(o.get("compras_materia_prima_evitadas_eur", 0), f"{ctx_o}.compras_materia_prima_evitadas_eur")
                    tec_c = _req_finite(o.get("coste_tecnico_reprocesado_eur", 0), f"{ctx_o}.coste_tecnico_reprocesado_eur")
                    expected_rw_mrg = mp_sav - tec_c - cogs
                    if abs(mrg_est - expected_rw_mrg) > 0.25:
                        fail(f"ARITMÉTICA MARGEN REWORK [{ctx_o}]: margen_estimado_eur={mrg_est} != mp-tec-cogs={expected_rw_mrg:.2f}.", last_date)
                    if abs(fact_n) > 0.01 or abs(c_fin) > 0.01 or abs(hold_s) > 0.01:
                        fail(f"REPROCESADO CON AHORRO/COBRO FICTICIO [{ctx_o}]: fact={fact_n}, fin={c_fin}, holding={hold_s}.", last_date)
                    net_fin_g = _req_finite(o.get("net_financial_gain_eur", 0), f"{ctx_o}.net_financial_gain_eur")
                    if abs(net_fin_g) > 0.01:
                        fail(f"REPROCESADO CON GANANCIA FINANCIERA FICTICIA [{ctx_o}]: net_financial_gain_eur={net_fin_g}.", last_date)

                # Validación de regalo contra stock real y coste único en vivo
                reg_sku = o.get("regalo_sku")
                reg_uds = _req_finite(o.get("regalo_uds", 0), f"{ctx_o}.regalo_uds")
                if reg_sku and reg_sku != "-" and reg_uds > 0:
                    g_info = gift_stock_map.get(reg_sku)
                    g_sku_row = sku_map.get(reg_sku)
                    if not g_sku_row and not g_info:
                        fail(f"REGALO SIN SKU VIVO [{ctx_o}]: regalo_sku={reg_sku} no existe en catálogo vivo.", last_date)
                    g_env_l = _req_finite((g_info or {}).get("envase_l", (g_sku_row or {}).get("envase_litros", 1.0)) or 1.0, f"{ctx_o}.gift_env_l")
                    g_free = _req_finite((g_info or {}).get("stock_libre_disponible", (g_sku_row or {}).get("stock_libre_disponible", 0)), f"{ctx_o}.gift_stock")
                    g_req_qty = reg_uds * g_env_l
                    if g_req_qty > g_free + 0.05:
                        fail(f"REGALO SIN STOCK [{ctx_o}]: {reg_uds}u ({g_req_qty} L/Kg) de {reg_sku} > stock_libre={g_free}.", last_date)
                    g_cost_live = _req_finite((g_info or {}).get("coste_unit_sku", (g_sku_row or {}).get("precio_unitario", 0)), f"{ctx_o}.gift_cost")
                    if g_cost_live <= 0:
                        fail(f"REGALO SIN COSTE VIVO [{ctx_o}]: {reg_sku} coste={g_cost_live}.", last_date)
                    if abs(c_reg - (g_req_qty * g_cost_live)) > 0.25:
                        fail(f"COSTE DE REGALO DESACTUALIZADO [{ctx_o}]: regalo_coste_eur={c_reg} != {g_req_qty}*{g_cost_live:.4f}={g_req_qty * g_cost_live:.2f}.", last_date)

                # Cuadre de calendario de cobros
                cash_sum = (
                    _req_finite(o.get("caja_0_30d_eur", 0), f"{ctx_o}.caja_0_30d_eur") +
                    _req_finite(o.get("caja_31_60d_eur", 0), f"{ctx_o}.caja_31_60d_eur") +
                    _req_finite(o.get("caja_61_90d_eur", 0), f"{ctx_o}.caja_61_90d_eur") +
                    _req_finite(o.get("caja_90d_plus_eur", 0), f"{ctx_o}.caja_90d_plus_eur")
                )
                if abs(cash_sum - fact_n) > 0.25:
                    fail(f"CALENDARIO COBROS DESCUADRADO [{ctx_o}]: buckets={cash_sum:.2f} != facturacion_neta={fact_n:.2f}.", last_date)

                # Validación simétrica ACCA de alternativas económicas y recomendación
                alts = o.get("alternativas_economicas")
                if not isinstance(alts, dict) or not isinstance(alts.get("alt_a"), dict):
                    fail(f"FALTAN ALTERNATIVAS ECONÓMICAS [{ctx_o}].", last_date)
                alt_a = alts["alt_a"]
                a_ben = _req_finite(alt_a.get("beneficio_incremental_eur", alt_a.get("ganancia_neta_eur")), f"{ctx_o}.alt_a.beneficio_incremental_eur")
                a_mrg_c = _req_finite(alt_a.get("margen_contable_eur"), f"{ctx_o}.alt_a.margen_contable_eur")
                if not is_rework:
                    exp_a_ben = (fact_n - c_rap - c_mrc - c_fin) + hold_s
                else:
                    exp_a_ben = _req_finite(o.get("compras_materia_prima_evitadas_eur", 0), f"{ctx_o}.mp") - _req_finite(o.get("coste_tecnico_reprocesado_eur", 0), f"{ctx_o}.tec")
                if abs(a_ben - exp_a_ben) > 0.25:
                    fail(f"ALT A INCREMENTAL CORRUPTO [{ctx_o}]: {a_ben} != esperado={exp_a_ben:.2f}.", last_date)
                if abs(a_mrg_c - mrg_est) > 0.25:
                    fail(f"ALT A MARGEN CONTABLE CORRUPTO [{ctx_o}]: {a_mrg_c} != margen_estimado={mrg_est:.2f}.", last_date)

                cand_alts = []
                for ak in ("alt_a", "alt_b", "alt_c"):
                    av = alts.get(ak)
                    if isinstance(av, dict):
                        b_val = _req_finite(av.get("beneficio_incremental_eur", av.get("ganancia_neta_eur")), f"{ctx_o}.{ak}.beneficio_incremental_eur")
                        g_val = _req_finite(av.get("ganancia_neta_eur"), f"{ctx_o}.{ak}.ganancia_neta_eur")
                        if abs(b_val - g_val) > 0.05:
                            fail(f"DESINCRONIZACIÓN ALTERNATIVA [{ctx_o}.{ak}]: beneficio_incremental={b_val} != ganancia_neta={g_val}.", last_date)
                        m_cont = _req_finite(av.get("margen_contable_eur"), f"{ctx_o}.{ak}.margen_contable_eur")
                        m_eur = _req_finite(av.get("margen_eur"), f"{ctx_o}.{ak}.margen_eur")
                        if abs(m_cont - m_eur) > 0.05:
                            fail(f"DESINCRONIZACIÓN MARGEN ALT [{ctx_o}.{ak}]: margen_contable={m_cont} != margen_eur={m_eur}.", last_date)
                        a_hld = _req_finite(av.get("ahorro_holding_eur", 0), f"{ctx_o}.{ak}.ahorro_holding_eur")
                        a_fac = _req_finite(av.get("facturacion_eur", 0), f"{ctx_o}.{ak}.facturacion_eur")
                        a_com = _req_finite(av.get("coste_comercial_eur", 0), f"{ctx_o}.{ak}.coste_comercial_eur")
                        a_fin = _req_finite(av.get("coste_fin_eur", 0), f"{ctx_o}.{ak}.coste_fin_eur")
                        if "Inacción" in str(av.get("tipo", "")) or "Inmovilizado" in str(av.get("nombre", "")):
                            if abs(b_val) > 0.05 or abs(m_cont) > 0.05 or abs(a_hld) > 0.05:
                                fail(f"ALT BASE INACCIÓN NO CERO [{ctx_o}.{ak}]: ben={b_val}, mrg={m_cont}, hld={a_hld}.", last_date)
                        elif a_fac <= 0.01:
                            # Reprocesado / ahorro MP interno en alternativa
                            if abs(a_hld) > 0.01:
                                fail(f"ALT REPROCESADO CON AHORRO HOLDING FICTICIO [{ctx_o}.{ak}].", last_date)
                            if not av.get("requiere_validacion_tecnica"):
                                fail(f"ALT REPROCESADO SIN BANDERA TÉCNICA [{ctx_o}.{ak}].", last_date)
                        else:
                            if abs(b_val - (a_fac - a_com - a_fin + a_hld)) > 0.35:
                                fail(f"ALT INCREMENTAL CORRUPTO [{ctx_o}.{ak}]: {b_val} != fact-com-fin+holding={(a_fac - a_com - a_fin + a_hld):.2f}.", last_date)
                        cand_alts.append((ak, b_val))

                max_ben = max(v for _, v in cand_alts)
                rec_key = alts.get("recomendada")
                rec_vals = [v for k_alt, v in cand_alts if k_alt == rec_key]
                if not rec_vals or rec_vals[0] < max_ben - 0.05:
                    fail(f"RECOMENDACIÓN INCOHERENTE [{ctx_o}]: recomendada={rec_key!r} no alcanza el máximo beneficio incremental ({max_ben:.2f} €).", last_date)

                sum_o_vol += vol
                sum_o_ing += fact_n
                sum_o_mrg += mrg_est
                sum_o_aho += hold_s
                sum_o_fin += c_fin

                p_id = o.get("plan_id")
                if p_id:
                    pa = plan_acc.setdefault(p_id, {"n": 0, "ing": 0.0, "mrg": 0.0, "fin": 0.0, "aho": 0.0})
                    pa["n"] += 1
                    pa["ing"] += fact_n
                    pa["mrg"] += mrg_est
                    pa["fin"] += c_fin
                    pa["aho"] += hold_s

            if abs(sum_o_vol - ex_qty) > 0.35 or abs(sum_o_ing - it_ing) > 0.50 or abs(sum_o_mrg - it_mrg) > 0.50 or abs(sum_o_aho - it_aho) > 0.50 or abs(sum_o_fin - it_fin) > 0.50:
                fail(f"DESCUADRE SKU vs OFERTAS [{ctx_it}]: suma ofertas ({sum_o_vol:.2f}u, ing={sum_o_ing:.2f}, mrg={sum_o_mrg:.2f}, aho={sum_o_aho:.2f}, fin={sum_o_fin:.2f}) != SKU ({ex_qty}, {it_ing}, {it_mrg}, {it_aho}, {it_fin}).", last_date)

        # Cuadre KPIs globales vs suma de SKUs
        if (
            abs(sum_it_vol - _req_finite(kpis.get("exceso_total_qty", 0), f"{scen_name}.kpis.exceso_total_qty")) > 1.0 or
            abs(sum_it_exceso_eur - _req_finite(kpis.get("exceso_total_cost_eur", 0), f"{scen_name}.kpis.exceso_total_cost_eur")) > 1.0 or
            abs(sum_it_ing_neto - _req_finite(kpis.get("ingreso_neto_total_eur", 0), f"{scen_name}.kpis.ing")) > 1.0 or
            abs(sum_it_margen - _req_finite(kpis.get("margen_neto_total_eur", 0), f"{scen_name}.kpis.mrg")) > 1.0 or
            abs(sum_it_ahorro - _req_finite(kpis.get("holding_cost_saved_total_eur", 0), f"{scen_name}.kpis.ahorro")) > 1.0 or
            abs(sum_it_fin - _req_finite(kpis.get("coste_financiero_aplazamiento_total_eur", 0), f"{scen_name}.kpis.fin")) > 1.0
        ):
            fail(f"DESCUADRE KPIs GLOBALES [{scen_name}] frente a suma de ítems SKU.", last_date)

        # Cuadre tarjetas de plan vs suma de ofertas agrupadas
        for p_card in plans:
            p_id = p_card.get("plan_id")
            pa = plan_acc.get(p_id)
            if not pa:
                fail(f"PLAN SIN OFERTAS [{scen_name}.{p_id}].", last_date)
            if (
                int(p_card.get("ofertas_count", -1)) != pa["n"] or
                abs(_req_finite(p_card.get("ingreso_neto_eur", 0), f"{scen_name}.{p_id}.ing") - pa["ing"]) > 1.0 or
                abs(_req_finite(p_card.get("margen_neto_eur", 0), f"{scen_name}.{p_id}.mrg") - pa["mrg"]) > 1.0 or
                abs(_req_finite(p_card.get("coste_fin_pago_eur", 0), f"{scen_name}.{p_id}.fin") - pa["fin"]) > 1.0 or
                abs(_req_finite(p_card.get("holding_cost_saved_eur", 0), f"{scen_name}.{p_id}.aho") - pa["aho"]) > 1.0
            ):
                fail(f"DESCUADRE TARJETA DE PLAN [{scen_name}.{p_id}] frente a sus ofertas.", last_date)

    print("   -> Suite Comercial de Stock (integridad física, regalos vivos, ACCA, KPIs, planes y frescura) validada.")




    # Recálculo independiente de la coincidencia del RAI



    rai_mayor = round(float(meta.get("rai_contable_mayor", 0.0) or 0.0), 2)



    rai_analitico = round(float(meta.get("rai_analitico_cascada", 0.0) or 0.0), 2)



    diff_recalc = round(abs(rai_mayor - rai_analitico), 2)

    diff_meta = round(float(meta.get("diferencia_conciliacion", 0.0) or 0.0), 2)

    if abs(diff_meta - diff_recalc) > 0.01:

        fail(f"Incoherencia en metadatos: 'diferencia_conciliacion' declarada ({diff_meta:.2f} €) no coincide con recálculo independiente abs({rai_mayor:.2f} - {rai_analitico:.2f}) = {diff_recalc:.2f} €", last_date)



    if diff_recalc > 0.01:

        fail(f"Descuadre en conciliación del RAI: RAI contable del Mayor ({rai_mayor:.2f} €) != RAI analítico de la cascada ({rai_analitico:.2f} €). Desviación no permitida: {diff_recalc:.2f} €.", last_date)



    coincidencia_rai = "COINCIDENTE (0,00 €)" if diff_recalc <= 0.01 else f"DESVIACIÓN ({diff_recalc:.2f} €)"

    # Validaciones económicas bloqueantes del modelo Coste Completo 100% PyG y Cotizador Unificado
    cc_2026 = fc.get("coste_completo_2026") or {}
    if not cc_2026:
        fail("Bloque 'coste_completo_2026' ausente en fullcost_data.json", last_date)
    base_uds = _req_finite(cc_2026.get("base_unidades_2026"), "coste_completo_2026.base_unidades_2026")
    tot_no_cogs = _req_finite(cc_2026.get("total_gastos_no_cogs_2026"), "coste_completo_2026.total_gastos_no_cogs_2026")
    planta_anual = _req_finite(cc_2026.get("planta_anual_absorbida_2026"), "coste_completo_2026.planta_anual_absorbida_2026")
    bolsa_pend = _req_finite(cc_2026.get("bolsa_pendiente_2026"), "coste_completo_2026.bolsa_pendiente_2026")
    extra_exact = _req_finite(cc_2026.get("extra_por_unidad_exacto"), "coste_completo_2026.extra_por_unidad_exacto")
    extra_round = _req_finite(cc_2026.get("extra_por_unidad"), "coste_completo_2026.extra_por_unidad")
    recup_exact = _req_finite(cc_2026.get("recuperacion_anual_exacta"), "coste_completo_2026.recuperacion_anual_exacta")

    if base_uds <= 0:
        fail(f"base_unidades_2026 debe ser > 0 (actual: {base_uds})", last_date)
    if abs((planta_anual + bolsa_pend) - tot_no_cogs) > 0.05:
        fail(f"DESCUADRE PUENTE ANUAL 2026: planta ({planta_anual:.2f}) + bolsa ({bolsa_pend:.2f}) != total_gastos_no_cogs ({tot_no_cogs:.2f})", last_date)
    if abs(extra_exact - (bolsa_pend / base_uds)) > 1e-6:
        fail(f"DESCUADRE EXTRA EXACTO 2026: {extra_exact:.8f} != {bolsa_pend / base_uds:.8f}", last_date)
    if abs(extra_round - round(extra_exact, 4)) > 1e-6:
        fail(f"DESCUADRE EXTRA REDONDEADO 2026: {extra_round:.4f} != {round(extra_exact, 4):.4f}", last_date)
    if abs(recup_exact - tot_no_cogs) > 0.05:
        fail(f"DESCUADRE RECUPERACIÓN ANUAL EXACTA: {recup_exact:.2f} != {tot_no_cogs:.2f}", last_date)

    cad_pr = ((fc.get("pricing") or {}).get("cadencias_produccion") or {})
    for ck in ("h_sint_110", "h_sint_120", "h_sint_130", "h_env_1L", "h_env_5L", "h_env_20L", "h_env_200L", "h_env_1000L"):
        if _req_finite(cad_pr.get(ck), f"pricing.cadencias_produccion.{ck}") <= 0:
            fail(f"Cadencia inválida en pricing.cadencias_produccion.{ck}", last_date)

    # Contraste cadencias contra trazabilidad_parametros (parámetros de origen)
    cad_src_vals = None
    for tp in (fc.get("trazabilidad_parametros") or []):
        if isinstance(tp.get("valor"), dict) and "h_env_20L" in tp["valor"]:
            cad_src_vals = tp["valor"]
            break
    if not cad_src_vals:
        fail("No se encontró el parámetro de cadencias de origen en trazabilidad_parametros", last_date)
    for ck in ("h_sint_110", "h_sint_120", "h_sint_130", "h_env_1L", "h_env_5L", "h_env_20L", "h_env_200L", "h_env_1000L"):
        if abs(float(cad_pr[ck]) - float(cad_src_vals[ck])) > 1e-6:
            fail(f"DESCUADRE CADENCIA {ck}: pricing={cad_pr[ck]} != origen={cad_src_vals[ck]}", last_date)

    tar_meta = ((fc.get("metadata") or {}).get("tarifas_ceco_eur_h") or {})
    caps_map = {str(c.get("ceco")): float(c.get("tarifa_eur_h", 0.0) or 0.0) for c in (fc.get("capacidades") or [])}
    for ck in ("110", "120", "130", "140"):
        t_m = _req_finite(tar_meta.get(ck), f"metadata.tarifas_ceco_eur_h.{ck}")
        if ck in caps_map and abs(t_m - caps_map[ck]) > 0.02:
            fail(f"DESCUADRE TARIFA CECO {ck}: metadata={t_m} != capacidades={caps_map[ck]}", last_date)

    # Contraste presupuesto anual 2026 contra fuente PyG origen
    pnl26_src = get_source(os.path.join("Gastos", "2026-Forecast-Codiagro P&L con octubre simulación.xlsx"))
    if not os.path.exists(pnl26_src):
        fail(f"Fuente presupuestaria anual PyG 2026 ausente: {pnl26_src}", last_date)
    wb_p26 = openpyxl.load_workbook(pnl26_src, data_only=True, read_only=True)
    ws_t26 = wb_p26["TABLA"]
    t_map = {}
    for rt in ws_t26.iter_rows(values_only=True):
        lt = str(rt[1] or "").strip() if len(rt) > 1 else ""
        vt = rt[14] if len(rt) > 14 else None
        if lt and isinstance(vt, (int, float)):
            t_map[lt] = abs(float(vt))
    ws_pm26 = wb_p26["PL Mensual 2026"]
    rap_s = pp_s = ett_s = 0.0
    for rm in ws_pm26.iter_rows(values_only=True):
        v23 = rm[22] if len(rm) > 22 else None
        if not isinstance(v23, (int, float)):
            continue
        cs = str(rm[8] or "").strip() if len(rm) > 8 else ""
        ep_m = str(rm[6] or "").strip() if len(rm) > 6 else ""
        if cs.startswith("709"):
            rap_s += abs(float(v23))
        elif cs.startswith("706") and not ep_m.startswith("15.-"):
            pp_s += abs(float(v23))
        elif cs.startswith("607"):
            ett_s += abs(float(v23))
    wb_p26.close()
    src_no_cogs = round(
        t_map["6.- Gastos de personal"] + t_map["7.- Otros gastos de explotación"]
        + t_map["8.- Amortización del inmovilizado"] + t_map["15.- Gastos financieros"]
        + rap_s + pp_s + ett_s, 2
    )
    if abs(tot_no_cogs - src_no_cogs) > 0.05:
        fail(f"DESCUADRE PRESUPUESTO PYG 2026 vs FUENTE: json={tot_no_cogs:.2f} != fuente={src_no_cogs:.2f}", last_date)

    # Carga de escandallos maestros (Gastos/Costes SKU.xlsx) para validar COGS unitario
    costes_sku_src = get_source(os.path.join("Gastos", "Costes SKU.xlsx"))
    wb_csku = openpyxl.load_workbook(costes_sku_src, data_only=True, read_only=True)
    ws_csku = wb_csku["Costes SKU"]
    master_cogs_map = {}
    for rc in ws_csku.iter_rows(min_row=5, values_only=True):
        sk_c = str(rc[0]).strip().upper() if rc[0] else ""
        if not sk_c:
            continue
        fm_v = float(rc[6] or 0)
        pk_v = float(rc[7] or 0)
        tm_v = float(rc[8] or 0)
        master_cogs_map[sk_c] = round(tm_v if tm_v > 0 else (fm_v + pk_v), 4)
    wb_csku.close()
    eq_sku = {"A-50TR": "A-50BBTR", "SPCE20001CR": "SPCE0001CR", "SPCE20005CR": "SPCE0005CR", "A-50BB": "A-50BBTR", "AGXKAEP0001": "AGXKAE0001RO"}

    for prod in fc.get("productos", []):
        sku_p = str(prod.get("sku", "?")).strip().upper()
        uds_p = float(prod.get("uds", 0.0) or 0.0)
        mat_p = float(prod.get("cogs_materiales", 0.0) or 0.0)
        pla_p = float(prod.get("coste_planta_fullcost", 0.0) or 0.0)
        ind_u = prod.get("coste_industrial_u")
        comp_u = prod.get("coste_completo_u")
        if prod.get("coste_pendiente"):
            if ind_u is not None or comp_u is not None:
                fail(f"Producto {sku_p} marcado coste_pendiente=True pero tiene coste numérico ({ind_u}, {comp_u})", last_date)
            continue
        if ind_u is None or comp_u is None:
            fail(f"Producto {sku_p} sin coste_pendiente pero con coste_industrial_u/coste_completo_u nulo", last_date)
        ind_u_f = _req_finite(ind_u, f"productos[{sku_p}].coste_industrial_u")
        comp_u_f = _req_finite(comp_u, f"productos[{sku_p}].coste_completo_u")
        cogs_u = _req_finite(prod.get("cogs_u"), f"productos[{sku_p}].cogs_u")
        planta_u = _req_finite(prod.get("planta_u"), f"productos[{sku_p}].planta_u")

        # Contraste COGS maestro obligatorio usando sku_maestro_cogs registrado
        sku_lk = str(prod.get("sku_maestro_cogs") or eq_sku.get(sku_p, sku_p)).strip().upper()
        if sku_lk not in master_cogs_map or master_cogs_map[sku_lk] <= 0:
            fail(f"SKU {sku_p} sin correspondencia validada en Costes SKU.xlsx (sku_maestro_cogs={sku_lk})", last_date)
        if abs(cogs_u - master_cogs_map[sku_lk]) > 0.0002:
            fail(f"DESCUADRE COGS MAESTRO SKU {sku_p} (maestro {sku_lk}): cogs_u={cogs_u} != Costes_SKU={master_cogs_map[sku_lk]}", last_date)

        # Recálculo de planta_u desde cadencias + tarifas + ruta + formato
        ruta_s = str(prod.get("ruta", ""))
        fmt_s = str(prod.get("formato", ""))
        c_proc = "110" if "110" in ruta_s else ("130" if "130" in ruta_s else "120")
        k_env = "h_env_1000L" if "1000L" in fmt_s else ("h_env_200L" if "200L" in fmt_s else ("h_env_20L" if "20L" in fmt_s else ("h_env_1L" if "1L" in fmt_s else "h_env_5L")))
        planta_exp = round(float(cad_pr[f"h_sint_{c_proc}"]) * float(tar_meta[c_proc]) + float(cad_pr[k_env]) * float(tar_meta["140"]), 4)
        if abs(planta_u - planta_exp) > 0.0002:
            fail(f"DESCUADRE PLANTA CADENCIA SKU {sku_p} (CECO {c_proc}, {k_env}): planta_u={planta_u} != esperado={planta_exp}", last_date)

        if abs(ind_u_f - round(cogs_u + planta_u, 4)) > 0.0002:
            fail(f"DESCUADRE INDUSTRIAL SKU {sku_p}: {ind_u_f} != cogs_u({cogs_u}) + planta_u({planta_u})", last_date)
        if abs(comp_u_f - round(ind_u_f + extra_round, 4)) > 0.0002:
            fail(f"DESCUADRE COMPLETO SKU {sku_p}: {comp_u_f} != ind({ind_u_f}) + extra({extra_round})", last_date)
        if uds_p <= 0 and (abs(mat_p) > 0.01 or abs(pla_p) > 0.01) and ind_u_f <= 0:
            fail(f"DEVOLUCIÓN CON COSTE FICTICIO CERO/NO POSITIVO en SKU {sku_p}: uds={uds_p}, ind_u={ind_u_f}", last_date)

    for ref_it in (cc_2026.get("catalogo_espana_ref") or []):
        if ref_it.get("coste_pendiente"):
            continue
        r_ind = _req_finite(ref_it.get("coste_industrial_u"), f"catalogo_espana_ref[{ref_it.get('sku')}].ind")
        r_comp = _req_finite(ref_it.get("coste_completo_u"), f"catalogo_espana_ref[{ref_it.get('sku')}].comp")
        if abs(r_comp - round(r_ind + extra_round, 4)) > 0.0002:
            fail(f"DESCUADRE CATÁLOGO REF {ref_it.get('sku')}: {r_comp} != {r_ind} + {extra_round}", last_date)

    print("   -> Validación económica bloqueante Coste Completo 100% PyG + Devoluciones + Cotizador + Maestros OK.")




    



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



