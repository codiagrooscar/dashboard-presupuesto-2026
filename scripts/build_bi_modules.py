import json
import os
import sys
import openpyxl
from collections import defaultdict

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

print("=== Generando Datos para Mejoras 1 y 2 (Forecast Accuracy & Concentración/Stress Test) ===")

# 1. Cargar datos de rentabilidad
profit_path = get_source(os.path.join("web_dashboard_direccion", "profitability_data.json"))
with open(profit_path, "r", encoding="utf-8") as f:
    profit_data = json.load(f)

# 2. Cargar datos de dashboard
dash_path = get_source(os.path.join("web_dashboard_direccion", "dashboard_data.json"))
with open(dash_path, "r", encoding="utf-8") as f:
    dash_data = json.load(f)

# --- MEJORA 1: FORECAST ACCURACY & BIAS ---
p_oct = dash_data.get("periods", {}).get("2026-10", {})
resumen_com = p_oct.get("resumen_comerciales", [])
comerciales_dict = p_oct.get("comerciales", {})

forecast_comerciales = []
for c in resumen_com:
    com_name = c.get("comercial")
    b_uds = float(c.get("budget_uds", 0))
    p_uds = float(c.get("pedidos_uds", 0))
    imp_neto = float(c.get("importe_neto", 0))
    pct_consec = float(c.get("pct_consec", 0))
    
    # Error neto agregado (bias)
    bias_uds = p_uds - b_uds
    bias_tipo = "Subestimó (Ventas > Presupuesto)" if bias_uds > 0 else ("Sobreestimó (Ventas < Presupuesto)" if bias_uds < 0 else "Exacto")
    
    # Hallazgo H21 (Auditoría WMAPE): Sumar errores absolutos por SKU antes de dividir
    com_lines = comerciales_dict.get(com_name, {}).get("lineas", [])
    if com_lines:
        com_sku_abs_error = sum(abs(float(l.get("pedidos_uds", 0)) - float(l.get("budget_uds", 0))) for l in com_lines)
        com_denom_orders = sum(float(l.get("pedidos_uds", 0)) for l in com_lines)
        com_denom_budget = sum(float(l.get("budget_uds", 0)) for l in com_lines)
        denom_wmape = com_denom_orders if com_denom_orders > 0 else com_denom_budget
        
        com_wmape_err = (com_sku_abs_error / denom_wmape) if denom_wmape > 0 else 0.0
        acc_pct = round(max(0.0, 1.0 - com_wmape_err) * 100, 1) if denom_wmape > 0 else 100.0
        wmape_err_pct = round(com_wmape_err * 100, 1)
    else:
        com_sku_abs_error = abs(p_uds - b_uds)
        denom_wmape = p_uds if p_uds > 0 else b_uds
        com_wmape_err = (com_sku_abs_error / denom_wmape) if denom_wmape > 0 else 0.0
        acc_pct = round(max(0.0, 1.0 - com_wmape_err) * 100, 1) if denom_wmape > 0 else 100.0
        wmape_err_pct = round(com_wmape_err * 100, 1)

    forecast_comerciales.append({
        "comercial": com_name,
        "clientes_activos": c.get("clientes_activos", 0),
        "budget_uds": b_uds,
        "pedidos_uds": p_uds,
        "desv_uds": bias_uds,
        "error_abs_sku_uds": round(com_sku_abs_error, 2),
        "pct_consec": pct_consec,
        "forecast_accuracy_pct": acc_pct,  # WMAPE Accuracy a nivel SKU
        "wmape_error_pct": wmape_err_pct,
        "bias_tipo": bias_tipo,
        "importe_neto_eur": imp_neto
    })

# Totales forecast empresa a nivel SKU consolidado (Hallazgo H21)
tot_b_uds = sum(x["budget_uds"] for x in forecast_comerciales)
tot_p_uds = sum(x["pedidos_uds"] for x in forecast_comerciales)
tot_imp_eur = sum(x["importe_neto_eur"] for x in forecast_comerciales)
tot_bias = tot_p_uds - tot_b_uds

all_sku_lines = []
for c_data in comerciales_dict.values():
    all_sku_lines.extend(c_data.get("lineas", []))

if all_sku_lines:
    tot_sku_abs_error = sum(abs(float(l.get("pedidos_uds", 0)) - float(l.get("budget_uds", 0))) for l in all_sku_lines)
    denom_tot = tot_p_uds if tot_p_uds > 0 else tot_b_uds
    tot_wmape_err = (tot_sku_abs_error / denom_tot) if denom_tot > 0 else 0.0
    tot_acc = round(max(0.0, 1.0 - tot_wmape_err) * 100, 1) if denom_tot > 0 else 100.0
    tot_wmape_err_pct = round(tot_wmape_err * 100, 1)
else:
    tot_sku_abs_error = sum(x.get("error_abs_sku_uds", abs(x["pedidos_uds"] - x["budget_uds"])) for x in forecast_comerciales)
    denom_tot = tot_p_uds if tot_p_uds > 0 else tot_b_uds
    tot_wmape_err = (tot_sku_abs_error / denom_tot) if denom_tot > 0 else 0.0
    tot_acc = round(max(0.0, 1.0 - tot_wmape_err) * 100, 1) if denom_tot > 0 else 100.0
    tot_wmape_err_pct = round(tot_wmape_err * 100, 1)

tot_consec_pct = round(tot_p_uds / tot_b_uds * 100, 1) if tot_b_uds > 0 else 0.0

forecast_resumen = {
    "tot_budget_uds": tot_b_uds,
    "tot_pedidos_uds": tot_p_uds,
    "tot_bias_uds": tot_bias,
    "tot_error_abs_sku_uds": round(tot_sku_abs_error, 2),
    "tot_importe_neto_eur": tot_imp_eur,
    "tot_consecucion_pct": tot_consec_pct,
    "tot_consecucion_volumen_pct": tot_consec_pct,
    "tot_forecast_accuracy_pct": tot_acc,
    "tot_forecast_accuracy_sku_wmape_pct": tot_acc,
    "tot_wmape_error_pct": tot_wmape_err_pct,
    "comerciales": forecast_comerciales
}

# --- MEJORA 2: CONCENTRACIÓN DE NEGOCIO & STRESS TEST ---
clientes = profit_data.get("clientes", [])
total_v26 = sum(c.get("ventas_2026", 0) for c in clientes)
total_m26 = sum((c.get("ventas_2026", 0) - c.get("cogs_2026", 0)) for c in clientes)

clientes_sorted = sorted(clientes, key=lambda x: x.get("ventas_2026", 0), reverse=True)
cum_v = 0
cum_m = 0
pareto_clientes = []
hhi_sum = 0 # Herfindahl-Hirschman Index

for idx, c in enumerate(clientes_sorted, start=1):
    v = float(c.get("ventas_2026", 0))
    m = v - float(c.get("cogs_2026", 0))
    cum_v += v
    cum_m += m
    pct_v = round((v / total_v26 * 100), 2) if total_v26 > 0 else 0
    pct_v_cum = round((cum_v / total_v26 * 100), 1) if total_v26 > 0 else 0
    pct_m_cum = round((cum_m / total_m26 * 100), 1) if total_m26 > 0 else 0
    m_pct = round((m / v * 100), 1) if v > 0 else 0
    
    # HHI component
    hhi_sum += (pct_v ** 2)

    pareto_clientes.append({
        "ranking": idx,
        "cliente": c.get("cliente"),
        "pais": c.get("pais"),
        "comercial": c.get("comercial"),
        "ventas_eur": round(v, 2),
        "margen_eur": round(m, 2),
        "margen_pct": m_pct,
        "pct_s_ventas": pct_v,
        "pct_acum_ventas": pct_v_cum,
        "pct_acum_margen": pct_m_cum,
        "es_top_5": idx <= 5,
        "es_top_10": idx <= 10
    })

# Concentración Países
paises = profit_data.get("paises", [])
total_v_pais = sum(p.get("ventas_2026", 0) for p in paises)
paises_sorted = sorted(paises, key=lambda x: x.get("ventas_2026", 0), reverse=True)
pareto_paises = []
cum_vp = 0
for idx, p in enumerate(paises_sorted, start=1):
    v = float(p.get("ventas_2026", 0))
    cum_vp += v
    pareto_paises.append({
        "ranking": idx,
        "pais": p.get("pais"),
        "ventas_eur": round(v, 2),
        "pct_s_ventas": round(v / total_v_pais * 100, 1) if total_v_pais > 0 else 0,
        "pct_acum_ventas": round(cum_vp / total_v_pais * 100, 1) if total_v_pais > 0 else 0
    })

# Concentración Familias / Productos
productos = profit_data.get("productos", [])
total_v_prod = sum(p.get("ventas_2026", 0) for p in productos)
prod_sorted = sorted(productos, key=lambda x: x.get("ventas_2026", 0), reverse=True)
pareto_productos = []
cum_vpr = 0
for idx, pr in enumerate(prod_sorted[:20], start=1):
    v = float(pr.get("ventas_2026", 0))
    cum_vpr += v
    pareto_productos.append({
        "ranking": idx,
        "sku": pr.get("sku"),
        "desc": pr.get("desc"),
        "familia": pr.get("familia"),
        "ventas_eur": round(v, 2),
        "pct_s_ventas": round(v / total_v_prod * 100, 1) if total_v_prod > 0 else 0,
        "pct_acum_ventas": round(cum_vpr / total_v_prod * 100, 1) if total_v_prod > 0 else 0
    })

concentracion_resumen = {
    "total_ventas_2026": round(total_v26, 2),
    "total_margen_2026": round(total_m26, 2),
    "hhi_index": round(hhi_sum, 1), # HHI > 1500 indica concentración moderada; > 2500 alta
    "hhi_diagnostico": "Moderadamente Concentrado (Riesgo Medio-Alto)" if hhi_sum > 1000 else "Baja Concentración",
    "top_1_pct": pareto_clientes[0]["pct_acum_ventas"] if pareto_clientes else 0,
    "top_5_pct": pareto_clientes[4]["pct_acum_ventas"] if len(pareto_clientes) >= 5 else 0,
    "top_10_pct": pareto_clientes[9]["pct_acum_ventas"] if len(pareto_clientes) >= 10 else 0,
    "clientes_pareto": pareto_clientes,
    "paises_pareto": pareto_paises,
    "productos_top20": pareto_productos
}

# Cargar datos existentes para no sobrescribir ni borrar los demás módulos (H22)
out_json = get_dest(os.path.join("web_dashboard_direccion", "bi_intelligence_data.json"))
src_bi_json = get_source(os.path.join("web_dashboard_direccion", "bi_intelligence_data.json"))
output_bi = {}
if os.path.exists(src_bi_json):
    with open(src_bi_json, "r", encoding="utf-8") as f:
        output_bi = json.load(f)

# Unificación de definición de fiabilidad (H21):
# Separar Consecución de Volumen Global frente a Forecast Accuracy Oficial WMAPE
consec_pct = forecast_resumen.get("tot_consecucion_pct", 81.5)
forecast_resumen["tot_consecucion_volumen_pct"] = consec_pct
forecast_resumen["tot_forecast_accuracy_sku_wmape_pct"] = tot_acc  # Calculado dinámicamente según fórmula WMAPE
forecast_resumen["nota_metodologia"] = f"La consecución de volumen total ({consec_pct}%) mide cumplimiento presupuestario agregado. La fiabilidad oficial SKU a SKU WMAPE ({tot_acc}%) mide precisión ponderada por volumen."

output_bi["mejora_1_forecast_accuracy"] = forecast_resumen
output_bi["mejora_2_concentracion_stress_test"] = concentracion_resumen

# Panel de Antigüedad de Deuda y Situación de Cobros (Ageing Codiagro Sep.xlsx)
output_bi["antiguedad_deuda"] = {
  "fecha_corte": "30/09/2026",
  "fuente": "Ageing Codiagro Sep.xlsx (Hoja SALDOS VENCIDOS)",
  "total_deuda_gestionada_eur": 718769.55,
  "totales_tramos": {
    "t0_30": 214321.29,
    "t31_60": 140363.27,
    "t61_90": 83501.43,
    "t91_120": 37387.11,
    "t121_plus": 243196.45,
    "total": 718769.55
  },
  "clientes": [
    {
      "cliente": "CODIMED EXPORT, S.L.",
      "pais": "España",
      "t0_30": 86922.03,
      "t31_60": 70138.04,
      "t61_90": 41670.28,
      "t91_120": 37387.11,
      "t121_plus": 151890.82,
      "total_vencido": 388008.28,
      "pct_total": 53.98,
      "badge": "Plan Oct-Feb / Contado",
      "badge_cls": "success",
      "observaciones": "Está pagando al contado los pedidos que hace, nos ha enviado un calendario de pago, desde octubre hasta febrero para ponerse al día."
    },
    {
      "cliente": "CAMPOEJIDO S.C.A.",
      "pais": "España",
      "t0_30": 122073.44,
      "t31_60": 0.0,
      "t61_90": 0.0,
      "t91_120": 0.0,
      "t121_plus": 0.0,
      "total_vencido": 122073.44,
      "pct_total": 16.98,
      "badge": "Abono Pendiente",
      "badge_cls": "warning",
      "observaciones": "Discrepancias en un abono, nada más esté claro nos hará el pago de inmediato. (Deuda 100% en tramo corriente 0-30d)."
    },
    {
      "cliente": "EUROFERTIL, S.A.",
      "pais": "Costa Rica",
      "t0_30": 0.0,
      "t31_60": 63959.63,
      "t61_90": 41831.15,
      "t91_120": 0.0,
      "t121_plus": 0.0,
      "total_vencido": 105790.78,
      "pct_total": 14.72,
      "badge": "Cobro Activo 8.000 €/sem",
      "badge_cls": "info",
      "observaciones": "Seguimos cobrando a razón de 8.000 euros semanales, aunque en el mes de agosto sólo han podido enviar 2 transferencias."
    },
    {
      "cliente": "DOGA TARIM TIC.SAN.VE TUR. LTD. STI",
      "pais": "Turquía",
      "t0_30": 0.0,
      "t31_60": 0.0,
      "t61_90": 0.0,
      "t91_120": 0.0,
      "t121_plus": 89504.67,
      "total_vencido": 89504.67,
      "pct_total": 12.45,
      "badge": "Concordato Turquía",
      "badge_cls": "danger",
      "observaciones": "Está negociando su salida del Concordato (suspensión de pagos) para poder negociar un plan de pagos. Neutralizada de caída comercial."
    },
    {
      "cliente": "T.J AGRICULTURA SL",
      "pais": "España",
      "t0_30": 3382.83,
      "t31_60": 6265.6,
      "t61_90": 0.0,
      "t91_120": 0.0,
      "t121_plus": 0.0,
      "total_vencido": 9648.43,
      "pct_total": 1.34,
      "badge": "Importes Reclamados",
      "badge_cls": "warning",
      "observaciones": "Importes reclamados. En breve pagará."
    },
    {
      "cliente": "DAVID BARAHONA HERNAEZ",
      "pais": "España",
      "t0_30": 1942.99,
      "t31_60": 0.0,
      "t61_90": 0.0,
      "t91_120": 0.0,
      "t121_plus": 0.0,
      "total_vencido": 1942.99,
      "pct_total": 0.27,
      "badge": "Importes Reclamados",
      "badge_cls": "warning",
      "observaciones": "Importes reclamados. En breve pagará."
    },
    {
      "cliente": "EURL BOUAZZA FILAHA",
      "pais": "Argelia",
      "t0_30": 0.0,
      "t31_60": 0.0,
      "t61_90": 0.0,
      "t91_120": 0.0,
      "t121_plus": 1800.96,
      "total_vencido": 1800.96,
      "pct_total": 0.25,
      "badge": "Discrepancia Flete",
      "badge_cls": "info",
      "observaciones": "Cliente Argelia. Discrepancias en el último pedido (flete y análisis). Se arreglará en el próximo."
    }
  ]
}

os.makedirs(os.path.dirname(out_json), exist_ok=True)
with open(out_json, "w", encoding="utf-8") as f:
    json.dump(output_bi, f, ensure_ascii=False, indent=2)

out_js = get_dest(os.path.join("web_dashboard_direccion", "bi_intelligence_data.js"))
with open(out_js, "w", encoding="utf-8") as f:
    f.write("window.BI_INTELLIGENCE_DATA = " + json.dumps(output_bi, ensure_ascii=False, indent=2) + ";\n")

print(f"Datos exportados a:\n  - {out_json}\n  - {out_js}")
print(f"Forecast Accuracy Global: {tot_acc}%")
print(f"Top 5 Clientes acumulan: {concentracion_resumen['top_5_pct']}% de las ventas")
print(f"Top 10 Clientes acumulan: {concentracion_resumen['top_10_pct']}% de las ventas")
