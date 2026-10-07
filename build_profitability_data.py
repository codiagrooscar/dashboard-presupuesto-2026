"""
Motor de Cálculo y Generación de Datos de Rentabilidad y Márgenes (COGS)
========================================================================
Cruza los costes unitarios de componentes y envasado (Costes SKU.xlsx)
con las ventas reales 2026 y el presupuesto 2026-2027 para calcular:
- COGS medio (Formulación media + Empaquetado media = Total medio)
- Margen Bruto (€ y %)
- Semáforo de Margen:
    * ROJO: < 60% (o Venta a pérdida < 0%)
    * AMARILLO: 60% a 66%
    * VERDE: > 66%
- Agregaciones por Cliente, Producto, Formato, Mercado/País y Comercial.
"""

import os
import json
import openpyxl
from collections import defaultdict

def build_profitability_data():
    print("=== 1. Cargando Costes SKU (Costes SKU.xlsx) ===")
    costes_path = os.path.join("Gastos", "Costes SKU.xlsx")
    wb_c = openpyxl.load_workbook(costes_path, data_only=True)
    ws_c = wb_c['Costes SKU']

    costes_map = {}
    for r in ws_c.iter_rows(min_row=5, values_only=True):
        sku = str(r[0]).strip().upper() if r[0] else ''
        if not sku:
            continue
        
        fam = str(r[1]).strip() if r[1] is not None else ''
        pais = str(r[2]).strip().upper() if r[2] is not None else ''
        form_max = float(r[3] or 0)
        pkg_max = float(r[4] or 0)
        tot_max = float(r[5] or 0)
        form_med = float(r[6] or 0)
        pkg_med = float(r[7] or 0)
        tot_med = float(r[8] or 0)
        estado = str(r[9]).strip() if r[9] else 'Completo'

        costes_map[sku] = {
            "sku": sku,
            "familia": fam,
            "pais": pais,
            "form_med": round(form_med, 4),
            "pkg_med": round(pkg_med, 4),
            "cogs_medio": round(tot_med, 4),
            "form_max": round(form_max, 4),
            "pkg_max": round(pkg_max, 4),
            "cogs_max": round(tot_max, 4),
            "estado": estado
        }

    print(f"Total SKUs con costes cargados: {len(costes_map)}")

    # Mapeos directos de variantes conocidas
    equivalencias = {
        "A-50TR": "A-50BBTR",
        "SPCE20001CR": "SPCE0001CR",
        "SPCE20005CR": "SPCE0005CR",
        "A-50BB": "A-50BBTR",
    }

    def get_cogs(sku_code, pais_code=""):
        sku_clean = str(sku_code).strip().upper()
        if sku_clean in equivalencias:
            sku_clean = equivalencias[sku_clean]
        
        if sku_clean in costes_map:
            return costes_map[sku_clean]

        # Si no está directo, buscar combinación SKU base + Sufijo país
        p_up = pais_code.strip().upper()
        # Intentar buscar variaciones que comiencen con sku_clean
        candidatos = [c for c in costes_map.keys() if c.startswith(sku_clean) or sku_clean.startswith(c)]
        if candidatos:
            # Si hay uno cuyo país coincida
            for c in candidatos:
                if costes_map[c]["pais"] == p_up:
                    return costes_map[c]
            return costes_map[candidatos[0]]
            
        return None

    print("\n=== 2. Procesando Ventas Reales y Presupuesto (Presupuesto_Ventas_2027_Definitivo.xlsx) ===")
    p_path = "Presupuesto_Ventas_2027_Definitivo.xlsx"
    wb_p = openpyxl.load_workbook(p_path, data_only=True)
    ws_p = wb_p['Detalle Líneas Clasificadas']

    header = [cell.value for cell in ws_p[1]]
    idx_year = header.index('Año')
    idx_mes = header.index('Mes')
    idx_com = header.index('Comercial')
    idx_pais = header.index('País')
    idx_client = header.index('Cliente')
    idx_cl_tipo = header.index('Clasificación Cliente')
    idx_sku = header.index('Código Artículo')
    idx_desc = header.index('Descripción Artículo')
    idx_env = header.index('Tamaño Envase') if 'Tamaño Envase' in header else None
    
    idx_uds = None
    idx_pr_med = None
    for i, h in enumerate(header):
        if h and 'unid' in str(h).lower(): idx_uds = i
        if h and ('precio' in str(h).lower() or 'pr. medio' in str(h).lower()): idx_pr_med = i

    print(f"Col Uds: {idx_uds}, Col Precio: {idx_pr_med}")

    lineas = []
    ventas_por_cliente = defaultdict(lambda: {
        "cliente": "", "pais": "", "comercial": "", "clasif": "",
        "ventas_2026": 0.0, "cogs_2026": 0.0, "uds_2026": 0.0,
        "ventas_2027": 0.0, "cogs_2027": 0.0, "uds_2027": 0.0,
        "skus_comprados": set(), "lineas_peligro": 0
    })

    ventas_por_producto = defaultdict(lambda: {
        "sku": "", "desc": "", "familia": "", "envase": 0.0,
        "cogs_medio": 0.0, "form_med": 0.0, "pkg_med": 0.0,
        "ventas_2026": 0.0, "cogs_2026": 0.0, "uds_2026": 0.0,
        "ventas_2027": 0.0, "cogs_2027": 0.0, "uds_2027": 0.0,
        "clientes": set(), "paises": set()
    })

    ventas_por_comercial = defaultdict(lambda: {
        "comercial": "",
        "ventas_2026": 0.0, "cogs_2026": 0.0, "uds_2026": 0.0,
        "ventas_2027": 0.0, "cogs_2027": 0.0, "uds_2027": 0.0,
        "clientes": set(), "skus": set(), "lineas_rojas": 0
    })

    ventas_por_pais = defaultdict(lambda: {
        "pais": "", "ambito": "",
        "ventas_2026": 0.0, "cogs_2026": 0.0, "uds_2026": 0.0,
        "ventas_2027": 0.0, "cogs_2027": 0.0, "uds_2027": 0.0,
        "clientes": set(), "skus": set()
    })

    ventas_por_formato = defaultdict(lambda: {
        "formato": "",
        "ventas_2026": 0.0, "cogs_2026": 0.0, "uds_2026": 0.0,
        "ventas_2027": 0.0, "cogs_2027": 0.0, "uds_2027": 0.0
    })

    ventas_por_sku_cliente = defaultdict(lambda: {
        "sku": "", "desc": "", "cliente": "", "pais": "", "comercial": "",
        "uds_2027": 0.0, "ventas_2027": 0.0, "cogs_2027": 0.0, "cogs_medio": 0.0
    })

    radar_alertas = []

    for row_idx in range(2, ws_p.max_row + 1):
        r = [ws_p.cell(row_idx, c).value for c in range(1, len(header) + 1)]
        sku = str(r[idx_sku]).strip().upper() if r[idx_sku] else ''
        if not sku: continue

        year = int(r[idx_year]) if r[idx_year] else 2026
        mes = str(r[idx_mes]).strip() if r[idx_mes] else ''
        com = str(r[idx_com]).strip() if r[idx_com] else 'General'
        pais = str(r[idx_pais]).strip().upper() if r[idx_pais] else 'ESPAÑA'
        client = str(r[idx_client]).strip() if r[idx_client] else 'Cliente General'
        cl_tipo = str(r[idx_cl_tipo]).strip() if r[idx_cl_tipo] else 'Estándar'
        desc = str(r[idx_desc]).strip() if r[idx_desc] else sku
        env_val = float(r[idx_env] or 1.0) if idx_env is not None and r[idx_env] is not None else 1.0
        uds = float(r[idx_uds] or 0.0)
        pr_unit = float(r[idx_pr_med] or 0.0)

        # Importe venta
        imp_venta = uds * pr_unit
        if uds <= 0 and imp_venta <= 0:
            continue

        cogs_info = get_cogs(sku, pais)
        cogs_u = cogs_info["cogs_medio"] if cogs_info else 0.0
        form_u = cogs_info["form_med"] if cogs_info else 0.0
        pkg_u = cogs_info["pkg_med"] if cogs_info else 0.0
        cogs_total = uds * cogs_u
        margen_eur = imp_venta - cogs_total
        margen_pct = (margen_eur / imp_venta * 100.0) if imp_venta > 0 else 0.0

        # Semáforo de Margen (Regla Oficial del Usuario):
        # Rojo: < 60%
        # Amarillo: 60% a 66%
        # Verde: > 66%
        if margen_pct < 60.0 or margen_eur < 0:
            semaforo = "ROJO"
        elif 60.0 <= margen_pct <= 66.0:
            semaforo = "AMARILLO"
        else:
            semaforo = "VERDE"

        es_perdida = (pr_unit < cogs_u) and (cogs_u > 0)

        # Formato agrupado (20L/Kg, 5L/Kg, 1L/Kg, IBC/Granel, etc.)
        if env_val >= 900:
            fmt_grupo = "Depósito 1000L / Granel"
        elif env_val >= 200:
            fmt_grupo = "Bidón 200L"
        elif env_val >= 20:
            fmt_grupo = "20 L / Kg"
        elif env_val >= 5:
            fmt_grupo = "5 L / Kg"
        elif env_val >= 1:
            fmt_grupo = "1 L / Kg"
        else:
            fmt_grupo = "Menor de 1 L/Kg"

        # Acumular cliente
        c_entry = ventas_por_cliente[client]
        c_entry["cliente"] = client
        c_entry["pais"] = pais
        c_entry["comercial"] = com
        c_entry["clasif"] = cl_tipo
        c_entry["skus_comprados"].add(sku)
        if year == 2026:
            c_entry["ventas_2026"] += imp_venta
            c_entry["cogs_2026"] += cogs_total
            c_entry["uds_2026"] += uds
        else:
            c_entry["ventas_2027"] += imp_venta
            c_entry["cogs_2027"] += cogs_total
            c_entry["uds_2027"] += uds
        if semaforo == "ROJO":
            c_entry["lineas_peligro"] += 1

        # Acumular producto
        p_entry = ventas_por_producto[sku]
        p_entry["sku"] = sku
        p_entry["desc"] = desc
        p_entry["familia"] = cogs_info["familia"] if cogs_info else ""
        p_entry["envase"] = env_val
        p_entry["cogs_medio"] = cogs_u
        p_entry["form_med"] = form_u
        p_entry["pkg_med"] = pkg_u
        p_entry["clientes"].add(client)
        p_entry["paises"].add(pais)
        if year == 2026:
            p_entry["ventas_2026"] += imp_venta
            p_entry["cogs_2026"] += cogs_total
            p_entry["uds_2026"] += uds
        else:
            p_entry["ventas_2027"] += imp_venta
            p_entry["cogs_2027"] += cogs_total
            p_entry["uds_2027"] += uds

        # Acumular comercial
        com_entry = ventas_por_comercial[com]
        com_entry["comercial"] = com
        com_entry["clientes"].add(client)
        com_entry["skus"].add(sku)
        if year == 2026:
            com_entry["ventas_2026"] += imp_venta
            com_entry["cogs_2026"] += cogs_total
            com_entry["uds_2026"] += uds
        else:
            com_entry["ventas_2027"] += imp_venta
            com_entry["cogs_2027"] += cogs_total
            com_entry["uds_2027"] += uds
        if semaforo == "ROJO":
            com_entry["lineas_rojas"] += 1

        # Acumular país
        pais_entry = ventas_por_pais[pais]
        pais_entry["pais"] = pais
        pais_entry["ambito"] = "Nacional" if pais in ["ESPAÑA", "ESP"] else "Exportación"
        pais_entry["clientes"].add(client)
        pais_entry["skus"].add(sku)
        if year == 2026:
            pais_entry["ventas_2026"] += imp_venta
            pais_entry["cogs_2026"] += cogs_total
            pais_entry["uds_2026"] += uds
        else:
            pais_entry["ventas_2027"] += imp_venta
            pais_entry["cogs_2027"] += cogs_total
            pais_entry["uds_2027"] += uds

        # Acumular formato
        fmt_entry = ventas_por_formato[fmt_grupo]
        fmt_entry["formato"] = fmt_grupo
        if year == 2026:
            fmt_entry["ventas_2026"] += imp_venta
            fmt_entry["cogs_2026"] += cogs_total
            fmt_entry["uds_2026"] += uds
        else:
            fmt_entry["ventas_2027"] += imp_venta
            fmt_entry["cogs_2027"] += cogs_total
            fmt_entry["uds_2027"] += uds

        # Acumular detalle por combinación (SKU, Cliente) para 2027
        if year == 2027:
            pair_key = (sku, client)
            pair_entry = ventas_por_sku_cliente[pair_key]
            pair_entry["sku"] = sku
            pair_entry["desc"] = desc
            pair_entry["cliente"] = client
            pair_entry["pais"] = pais
            pair_entry["comercial"] = com
            pair_entry["uds_2027"] += uds
            pair_entry["ventas_2027"] += imp_venta
            pair_entry["cogs_2027"] += cogs_total
            pair_entry["cogs_medio"] = cogs_u

        # Alerta para el Radar (solo si es pérdida o margen < 60%)
        if semaforo == "ROJO" and imp_venta > 0:
            radar_alertas.append({
                "year": year,
                "mes": mes,
                "sku": sku,
                "desc": desc,
                "cliente": client,
                "pais": pais,
                "comercial": com,
                "unidades": round(uds, 1),
                "precio_venta": round(pr_unit, 2),
                "cogs_medio": round(cogs_u, 2),
                "importe_venta": round(imp_venta, 2),
                "cogs_total": round(cogs_total, 2),
                "margen_eur": round(margen_eur, 2),
                "margen_pct": round(margen_pct, 1),
                "es_perdida": es_perdida,
                "semaforo": semaforo
            })

    print(f"Total líneas con margen rojo / alerta crítica: {len(radar_alertas)}")

    # Formatear estructuras finales para JSON
    clientes_list = []
    for c_name, d in ventas_por_cliente.items():
        v26 = round(d["ventas_2026"], 2)
        c26 = round(d["cogs_2026"], 2)
        m26_eur = round(v26 - c26, 2)
        m26_pct = round((m26_eur / v26 * 100), 1) if v26 > 0 else 0.0

        v27 = round(d["ventas_2027"], 2)
        c27 = round(d["cogs_2027"], 2)
        m27_eur = round(v27 - c27, 2)
        m27_pct = round((m27_eur / v27 * 100), 1) if v27 > 0 else 0.0

        # Clasificación matriz estratégica
        # Alto volumen: > 50.000 €; Alto margen: > 66%
        v_ref = v27 if v27 > 0 else v26
        m_ref = m27_pct if v27 > 0 else m26_pct
        if v_ref >= 50000 and m_ref >= 66.0:
            matriz = "Estrella (Alto Vol. / Alto Margen)"
        elif v_ref >= 50000 and m_ref < 60.0:
            matriz = "Volumen de Riesgo (Alto Vol. / Bajo Margen)"
        elif v_ref >= 50000:
            matriz = "Core Estable (Alto Vol. / Margen Medio)"
        elif v_ref < 50000 and m_ref >= 66.0:
            matriz = "Boutique / Potencial (Bajo Vol. / Alto Margen)"
        elif v_ref < 50000 and m_ref < 60.0:
            matriz = "Crítico / Déficit (Bajo Vol. / Bajo Margen)"
        else:
            matriz = "Desarrollo (Margen Medio)"

        sem_cl = "ROJO" if m_ref < 60.0 else ("AMARILLO" if m_ref <= 66.0 else "VERDE")

        clientes_list.append({
            "cliente": c_name,
            "pais": d["pais"],
            "comercial": d["comercial"],
            "clasif_cliente": d["clasif"],
            "skus_count": len(d["skus_comprados"]),
            "ventas_2026": v26,
            "cogs_2026": c26,
            "margen_2026_eur": m26_eur,
            "margen_2026_pct": m26_pct,
            "ventas_2027": v27,
            "cogs_2027": c27,
            "margen_2027_eur": m27_eur,
            "margen_2027_pct": m27_pct,
            "matriz_estrategica": matriz,
            "semaforo": sem_cl,
            "lineas_peligro": d["lineas_peligro"]
        })

    # Ordenar clientes por facturación 2027 descendente
    clientes_list.sort(key=lambda x: x["ventas_2027"], reverse=True)

    # Formatear productos
    productos_list = []
    for p_sku, d in ventas_por_producto.items():
        v26 = round(d["ventas_2026"], 2)
        c26 = round(d["cogs_2026"], 2)
        m26_eur = round(v26 - c26, 2)
        m26_pct = round((m26_eur / v26 * 100), 1) if v26 > 0 else 0.0

        v27 = round(d["ventas_2027"], 2)
        c27 = round(d["cogs_2027"], 2)
        m27_eur = round(v27 - c27, 2)
        m27_pct = round((m27_eur / v27 * 100), 1) if v27 > 0 else 0.0

        m_ref = m27_pct if v27 > 0 else m26_pct
        sem_p = "ROJO" if m_ref < 60.0 else ("AMARILLO" if m_ref <= 66.0 else "VERDE")

        pct_envase = round((d["pkg_med"] / d["cogs_medio"] * 100), 1) if d["cogs_medio"] > 0 else 0.0

        productos_list.append({
            "sku": p_sku,
            "desc": d["desc"],
            "familia": d["familia"],
            "envase": d["envase"],
            "cogs_medio": d["cogs_medio"],
            "form_med": d["form_med"],
            "pkg_med": d["pkg_med"],
            "pct_envase_sobre_cogs": pct_envase,
            "ventas_2026": v26,
            "cogs_2026": c26,
            "uds_2026": round(d["uds_2026"], 0),
            "margen_2026_eur": m26_eur,
            "margen_2026_pct": m26_pct,
            "ventas_2027": v27,
            "cogs_2027": c27,
            "uds_2027": round(d["uds_2027"], 0),
            "margen_2027_eur": m27_eur,
            "margen_2027_pct": m27_pct,
            "semaforo": sem_p,
            "clientes_count": len(d["clientes"]),
            "paises_count": len(d["paises"])
        })
    productos_list.sort(key=lambda x: x["ventas_2027"], reverse=True)

    # Formatear comerciales
    comerciales_list = []
    for com_name, d in ventas_por_comercial.items():
        v26 = round(d["ventas_2026"], 2)
        c26 = round(d["cogs_2026"], 2)
        m26_eur = round(v26 - c26, 2)
        m26_pct = round((m26_eur / v26 * 100), 1) if v26 > 0 else 0.0

        v27 = round(d["ventas_2027"], 2)
        c27 = round(d["cogs_2027"], 2)
        m27_eur = round(v27 - c27, 2)
        m27_pct = round((m27_eur / v27 * 100), 1) if v27 > 0 else 0.0

        comerciales_list.append({
            "comercial": com_name,
            "clientes_count": len(d["clientes"]),
            "skus_count": len(d["skus"]),
            "ventas_2026": v26,
            "cogs_2026": c26,
            "margen_2026_eur": m26_eur,
            "margen_2026_pct": m26_pct,
            "ventas_2027": v27,
            "cogs_2027": c27,
            "margen_2027_eur": m27_eur,
            "margen_2027_pct": m27_pct,
            "lineas_rojas": d["lineas_rojas"]
        })
    comerciales_list.sort(key=lambda x: x["ventas_2027"], reverse=True)

    # Formatear países
    paises_list = []
    for p_name, d in ventas_por_pais.items():
        v26 = round(d["ventas_2026"], 2)
        c26 = round(d["cogs_2026"], 2)
        m26_eur = round(v26 - c26, 2)
        m26_pct = round((m26_eur / v26 * 100), 1) if v26 > 0 else 0.0

        v27 = round(d["ventas_2027"], 2)
        c27 = round(d["cogs_2027"], 2)
        m27_eur = round(v27 - c27, 2)
        m27_pct = round((m27_eur / v27 * 100), 1) if v27 > 0 else 0.0

        paises_list.append({
            "pais": p_name,
            "ambito": d["ambito"],
            "clientes_count": len(d["clientes"]),
            "skus_count": len(d["skus"]),
            "ventas_2026": v26,
            "cogs_2026": c26,
            "margen_2026_eur": m26_eur,
            "margen_2026_pct": m26_pct,
            "ventas_2027": v27,
            "cogs_2027": c27,
            "margen_2027_eur": m27_eur,
            "margen_2027_pct": m27_pct
        })
    paises_list.sort(key=lambda x: x["ventas_2027"], reverse=True)

    # Formatear formatos
    formatos_list = []
    for f_name, d in ventas_por_formato.items():
        v26 = round(d["ventas_2026"], 2)
        c26 = round(d["cogs_2026"], 2)
        m26_eur = round(v26 - c26, 2)
        m26_pct = round((m26_eur / v26 * 100), 1) if v26 > 0 else 0.0

        v27 = round(d["ventas_2027"], 2)
        c27 = round(d["cogs_2027"], 2)
        m27_eur = round(v27 - c27, 2)
        m27_pct = round((m27_eur / v27 * 100), 1) if v27 > 0 else 0.0

        formatos_list.append({
            "formato": f_name,
            "ventas_2026": v26,
            "cogs_2026": c26,
            "margen_2026_eur": m26_eur,
            "margen_2026_pct": m26_pct,
            "ventas_2027": v27,
            "cogs_2027": c27,
            "margen_2027_eur": m27_eur,
            "margen_2027_pct": m27_pct
        })
    formatos_list.sort(key=lambda x: x["ventas_2027"], reverse=True)

    # Ordenar alertas de radar: primero ventas a pérdida, luego por menor margen
    radar_alertas.sort(key=lambda x: (not x["es_perdida"], x["margen_pct"]))

    # Construir sku_comparador y cliente_detalle para drilldown y comparativa
    sku_comparador = defaultdict(list)
    cliente_detalle = defaultdict(list)
    for (s, c), d in ventas_por_sku_cliente.items():
        v27 = round(d["ventas_2027"], 2)
        u27 = round(d["uds_2027"], 1)
        c27 = round(d["cogs_2027"], 2)
        pr_med = round(v27 / u27, 2) if u27 > 0 else 0.0
        m_eur = round(v27 - c27, 2)
        m_pct = round((m_eur / v27 * 100), 1) if v27 > 0 else 0.0
        cogs_u = d["cogs_medio"]
        es_perdida = (pr_med < cogs_u) and (cogs_u > 0)
        sem = "ROJO" if (m_pct < 60.0 or m_eur < 0) else ("AMARILLO" if m_pct <= 66.0 else "VERDE")
        
        sku_comparador[s].append({
            "cliente": c,
            "pais": d["pais"],
            "comercial": d["comercial"],
            "uds": u27,
            "precio_medio": pr_med,
            "cogs_medio": cogs_u,
            "ventas_eur": v27,
            "cogs_eur": c27,
            "margen_eur": m_eur,
            "margen_pct": m_pct,
            "es_perdida": es_perdida,
            "semaforo": sem
        })
        
        cliente_detalle[c].append({
            "sku": s,
            "desc": d["desc"],
            "uds": u27,
            "precio_medio": pr_med,
            "cogs_medio": cogs_u,
            "ventas_eur": v27,
            "cogs_eur": c27,
            "margen_eur": m_eur,
            "margen_pct": m_pct,
            "es_perdida": es_perdida,
            "semaforo": sem
        })

    for s in sku_comparador:
        sku_comparador[s].sort(key=lambda x: x["ventas_eur"], reverse=True)
    for c in cliente_detalle:
        cliente_detalle[c].sort(key=lambda x: x["ventas_eur"], reverse=True)

    # Totales globales de la empresa
    tot_v26 = sum(c["ventas_2026"] for c in clientes_list)
    tot_c26 = sum(c["cogs_2026"] for c in clientes_list)
    tot_m26_eur = tot_v26 - tot_c26
    tot_m26_pct = round((tot_m26_eur / tot_v26 * 100), 1) if tot_v26 > 0 else 0.0

    tot_v27 = sum(c["ventas_2027"] for c in clientes_list)
    tot_c27 = sum(c["cogs_2027"] for c in clientes_list)
    tot_m27_eur = tot_v27 - tot_c27
    tot_m27_pct = round((tot_m27_eur / tot_v27 * 100), 1) if tot_v27 > 0 else 0.0

    count_rojo = sum(1 for c in clientes_list if c["semaforo"] == "ROJO")
    count_amarillo = sum(1 for c in clientes_list if c["semaforo"] == "AMARILLO")
    count_verde = sum(1 for c in clientes_list if c["semaforo"] == "VERDE")

    output_data = {
        "metadata": {
            "title": "Codiagro - Análisis de Rentabilidad y Márgenes COGS",
            "cogs_coste_usado": "Total Medio (Formulación Media + Empaquetado Media)",
            "regla_semaforo": {
                "rojo": "< 60%",
                "amarillo": "60% - 66%",
                "verde": "> 66%"
            },
            "skus_costes_count": len(costes_map),
            "alertas_radar_count": len(radar_alertas),
            "ventas_perdida_count": sum(1 for a in radar_alertas if a["es_perdida"])
        },
        "totales_empresa": {
            "ventas_2026": round(tot_v26, 2),
            "cogs_2026": round(tot_c26, 2),
            "margen_2026_eur": round(tot_m26_eur, 2),
            "margen_2026_pct": tot_m26_pct,
            "ventas_2027": round(tot_v27, 2),
            "cogs_2027": round(tot_c27, 2),
            "margen_2027_eur": round(tot_m27_eur, 2),
            "margen_2027_pct": tot_m27_pct,
            "clientes_rojo": count_rojo,
            "clientes_amarillo": count_amarillo,
            "clientes_verde": count_verde
        },
        "alertas_radar": radar_alertas,
        "clientes": clientes_list,
        "productos": productos_list,
        "comerciales": comerciales_list,
        "paises": paises_list,
        "formatos": formatos_list,
        "sku_comparador": sku_comparador,
        "cliente_detalle": cliente_detalle
    }

    # Guardar en raíz y en web_dashboard
    out_json = os.path.join("web_dashboard", "profitability_data.json")
    out_js = os.path.join("web_dashboard", "profitability_data.js")

    with open(out_json, "w", encoding="utf-8") as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)

    with open(out_js, "w", encoding="utf-8") as f:
        f.write("window.PROFITABILITY_DATA = " + json.dumps(output_data, ensure_ascii=False) + ";\n")

    print(f"\nDatos guardados exitosamente en:\n  - {out_json}\n  - {out_js}")
    print(f"Resumen Empresa 2027:")
    print(f"  * Facturación: {tot_v27:,.2f} €")
    print(f"  * COGS Medio:  {tot_c27:,.2f} €")
    print(f"  * Margen Bruto: {tot_m27_eur:,.2f} € ({tot_m27_pct}%)")
    print(f"  * Ventas a Pérdida detectadas: {output_data['metadata']['ventas_perdida_count']}")
    print(f"  * Clientes en Rojo (<60%): {count_rojo} | Amarillo (60-66%): {count_amarillo} | Verde (>66%): {count_verde}")

if __name__ == "__main__":
    build_profitability_data()
