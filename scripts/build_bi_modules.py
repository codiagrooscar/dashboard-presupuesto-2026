import json
import os
import openpyxl
from collections import defaultdict

print("=== Generando Datos para Mejoras 1 y 2 (Forecast Accuracy & Concentración/Stress Test) ===")

# 1. Cargar datos de rentabilidad
profit_path = os.path.join("web_dashboard_direccion", "profitability_data.json")
with open(profit_path, "r", encoding="utf-8") as f:
    profit_data = json.load(f)

# 2. Cargar datos de dashboard
dash_path = os.path.join("web_dashboard_direccion", "dashboard_data.json")
with open(dash_path, "r", encoding="utf-8") as f:
    dash_data = json.load(f)

# --- MEJORA 1: FORECAST ACCURACY & BIAS ---
p_oct = dash_data.get("periods", {}).get("2026-10", {})
resumen_com = p_oct.get("resumen_comerciales", [])

forecast_comerciales = []
for c in resumen_com:
    com_name = c.get("comercial")
    b_uds = float(c.get("budget_uds", 0))
    p_uds = float(c.get("pedidos_uds", 0))
    imp_neto = float(c.get("importe_neto", 0))
    pct_consec = float(c.get("pct_consec", 0))
    
    # Error absoluto y error con signo (bias)
    bias_uds = p_uds - b_uds
    bias_tipo = "Subestimó (Ventas > Presupuesto)" if bias_uds > 0 else ("Sobreestimó (Ventas < Presupuesto)" if bias_uds < 0 else "Exacto")
    
    # Forecast accuracy standard: 1 - (|Actual - Forecast| / max(Actual, Forecast))
    denom = max(p_uds, b_uds)
    acc_pct = round(max(0, 1 - (abs(p_uds - b_uds) / denom)) * 100, 1) if denom > 0 else 100.0

    forecast_comerciales.append({
        "comercial": com_name,
        "clientes_activos": c.get("clientes_activos", 0),
        "budget_uds": b_uds,
        "pedidos_uds": p_uds,
        "desv_uds": bias_uds,
        "pct_consec": pct_consec,
        "forecast_accuracy_pct": acc_pct,
        "bias_tipo": bias_tipo,
        "importe_neto_eur": imp_neto
    })

# Totales forecast empresa
tot_b_uds = sum(x["budget_uds"] for x in forecast_comerciales)
tot_p_uds = sum(x["pedidos_uds"] for x in forecast_comerciales)
tot_imp_eur = sum(x["importe_neto_eur"] for x in forecast_comerciales)
tot_bias = tot_p_uds - tot_b_uds
tot_denom = max(tot_b_uds, tot_p_uds)
tot_acc = round(max(0, 1 - (abs(tot_p_uds - tot_b_uds) / tot_denom)) * 100, 1) if tot_denom > 0 else 100.0

forecast_resumen = {
    "tot_budget_uds": tot_b_uds,
    "tot_pedidos_uds": tot_p_uds,
    "tot_bias_uds": tot_bias,
    "tot_importe_neto_eur": tot_imp_eur,
    "tot_consecucion_pct": round(tot_p_uds / tot_b_uds * 100, 1) if tot_b_uds > 0 else 0,
    "tot_forecast_accuracy_pct": tot_acc,
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

output_bi = {
    "mejora_1_forecast_accuracy": forecast_resumen,
    "mejora_2_concentracion_stress_test": concentracion_resumen
}

out_json = os.path.join("web_dashboard_direccion", "bi_intelligence_data.json")
with open(out_json, "w", encoding="utf-8") as f:
    json.dump(output_bi, f, ensure_ascii=False, indent=2)

out_js = os.path.join("web_dashboard_direccion", "bi_intelligence_data.js")
with open(out_js, "w", encoding="utf-8") as f:
    f.write("window.BI_INTELLIGENCE_DATA = " + json.dumps(output_bi, ensure_ascii=False, indent=2) + ";\n")

print(f"Datos exportados a:\n  - {out_json}\n  - {out_js}")
print(f"Forecast Accuracy Global: {tot_acc}%")
print(f"Top 5 Clientes acumulan: {concentracion_resumen['top_5_pct']}% de las ventas")
print(f"Top 10 Clientes acumulan: {concentracion_resumen['top_10_pct']}% de las ventas")
