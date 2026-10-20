"""
incorporar_septiembre_real.py
==============================
Consolida las ventas reales de Septiembre 2026 procedentes de:
  "Escritorio/REPORTING SIMPLIFICADO/SEPTIEMBRE/Ventas/2026-09-Codiagro Ventas_new version Report_VF V2.xlsx"
con el histórico previo de Enero a Agosto de 2026.

Acciones:
1. Genera 'Ventas reales de Enero a Septiembre de 2026.xlsx'.
2. Actualiza 'Presupuesto_Ventas_2027_Definitivo.xlsx' (Detalle Líneas Clasificadas y Previsión Matriz Horizontal).
3. Vuelve a ejecutar 'build_profitability_data.py' para sincronizar el dashboard web.
"""

import os
import shutil
import unicodedata
import openpyxl
from collections import defaultdict

def norm(s):
    if not isinstance(s, str): return ''
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn').upper().strip()

def run_integration():
    print("=== 1. Cargando Ventas Enero a Agosto 2026 ===")
    wb_ago = openpyxl.load_workbook('Ventas reales de Eebro a Agosto de 2026.xlsx', data_only=True)
    ws_ago = wb_ago.active

    ene_ago = defaultdict(lambda: {'uds': 0.0, 'eur': 0.0, 'desc': '', 'cli_orig': '', 'sku_orig': ''})
    for r in list(ws_ago.iter_rows(values_only=True))[1:]:
        cli = str(r[0] or '').strip()
        sku = str(r[1] or '').strip().upper()
        desc = str(r[2] or '').strip()
        uds = float(r[3] or 0)
        eur = float(r[4] or 0)
        if not cli or not sku: continue
        k = (norm(cli), sku)
        ene_ago[k]['uds'] += uds
        ene_ago[k]['eur'] += eur
        ene_ago[k]['desc'] = desc
        ene_ago[k]['cli_orig'] = cli
        ene_ago[k]['sku_orig'] = sku

    tot_ago_uds = sum(x['uds'] for x in ene_ago.values())
    tot_ago_eur = sum(x['eur'] for x in ene_ago.values())
    print(f"Ene-Ago 2026: {len(ene_ago)} líneas | {tot_ago_uds:,.2f} u | {tot_ago_eur:,.2f} €")

    print("\n=== 2. Cargando Ventas Reales de Septiembre 2026 ===")
    p_sep = r'C:\Users\oscar.ocampo\OneDrive - Agroquímica Codiagro S.L\Escritorio\REPORTING SIMPLIFICADO\SEPTIEMBRE\Ventas\2026-09-Codiagro Ventas_new version Report_VF V2.xlsx'
    wb_sep = openpyxl.load_workbook(p_sep, read_only=True)
    ws_sep = wb_sep['BBDD Ventas']

    sep_lines = []
    for r in list(ws_sep.iter_rows(values_only=True))[4:]:
        m = str(r[1] or '').lower()
        if not m.startswith('sep'): continue
        cli = str(r[5] or '').strip()
        cod_cli = str(r[4] or '').strip()
        sku = str(r[2] or '').strip().upper()
        desc = str(r[3] or '').strip()
        uds = float(r[14] or 0)
        eur = float(r[15] or 0)
        if not cli or not sku: continue
        if uds == 0 and eur == 0: continue
        sep_lines.append({
            'cod_cli': cod_cli,
            'cliente': cli,
            'sku': sku,
            'desc': desc,
            'uds': uds,
            'eur': eur,
            'pr_med': (eur / uds) if uds > 0 else 0.0
        })

    tot_sep_uds = sum(x['uds'] for x in sep_lines)
    tot_sep_eur = sum(x['eur'] for x in sep_lines)
    print(f"Septiembre 2026: {len(sep_lines)} transacciones | {tot_sep_uds:,.2f} u | {tot_sep_eur:,.2f} €")

    print("\n=== 3. Generando 'Ventas reales de Enero a Septiembre de 2026.xlsx' ===")
    consolidado = defaultdict(lambda: {'uds': 0.0, 'eur': 0.0, 'desc': '', 'cli_orig': '', 'sku_orig': ''})
    for k, v in ene_ago.items():
        consolidado[k]['uds'] += v['uds']
        consolidado[k]['eur'] += v['eur']
        consolidado[k]['desc'] = v['desc']
        consolidado[k]['cli_orig'] = v['cli_orig']
        consolidado[k]['sku_orig'] = v['sku_orig']

    for s in sep_lines:
        k = (norm(s['cliente']), s['sku'])
        consolidado[k]['uds'] += s['uds']
        consolidado[k]['eur'] += s['eur']
        if not consolidado[k]['desc']: consolidado[k]['desc'] = s['desc']
        if not consolidado[k]['cli_orig']: consolidado[k]['cli_orig'] = s['cliente']
        if not consolidado[k]['sku_orig']: consolidado[k]['sku_orig'] = s['sku']

    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "Ventas Reales Ene-Sep 2026"
    ws_out.append(['CLIENTE', 'CodigoArticulo', 'DescripcionArticulo', 'Unidades ', 'Importe', 'Pr. Medio 2026'])

    for k, v in sorted(consolidado.items(), key=lambda x: (x[1]['cli_orig'], x[1]['sku_orig'])):
        pr_med = (v['eur'] / v['uds']) if v['uds'] > 0 else 0.0
        ws_out.append([v['cli_orig'], v['sku_orig'], v['desc'], round(v['uds'], 2), round(v['eur'], 2), round(pr_med, 4)])

    out_ene_sep = 'Ventas reales de Enero a Septiembre de 2026.xlsx'
    wb_out.save(out_ene_sep)
    tot_cons_uds = sum(x['uds'] for x in consolidado.values())
    tot_cons_eur = sum(x['eur'] for x in consolidado.values())
    print(f"Guardado '{out_ene_sep}': {len(consolidado)} líneas agrupadas | {tot_cons_uds:,.2f} u | {tot_cons_eur:,.2f} €")

    print("\n=== 4. Actualizando 'Presupuesto_Ventas_2027_Definitivo.xlsx' ===")
    budget_file = 'Presupuesto_Ventas_2027_Definitivo.xlsx'
    backup_file = budget_file.replace('.xlsx', '_BACKUP_PRE_SEP_REAL.xlsx')
    if not os.path.exists(backup_file):
        shutil.copy2(budget_file, backup_file)
        print(f"Copia de seguridad creada: {backup_file}")

    wb_b = openpyxl.load_workbook(budget_file, data_only=False)

    # 4.1 Mapeos de clientes para comercial, país y clasificación
    client_master = {}
    ws_lines = wb_b['Detalle Líneas Clasificadas']
    header_l = [c.value for c in ws_lines[1]]

    for r in ws_lines.iter_rows(min_row=2, values_only=True):
        cli = str(r[4] or '').strip()
        cn = norm(cli)
        if cn and cn not in client_master:
            client_master[cn] = {
                'comercial': r[2],
                'pais': r[3],
                'cliente': r[4],
                'clasif': r[5]
            }

    # Asignaciones especiales conocidas
    client_master[norm('FINCA DOÑA ANA C.B.')] = {'comercial': 'Javier', 'pais': 'ESPAÑA', 'cliente': 'FINCA DOÑA ANA C.B.', 'clasif': 'Estándar'}
    client_master[norm('FITOSANITARIOS CARCAIXENT, S.L.')] = {'comercial': 'Javier', 'pais': 'ESPAÑA', 'cliente': 'FITOSANITARIOS CARCAIXENT, S.L.', 'clasif': 'Consolidado/estratégico'}
    client_master[norm('ALMENDRALIA IBÉRICA, S.L.U.')] = {'comercial': 'Javier', 'pais': 'ESPAÑA', 'cliente': 'ALMENDRALIA IBÉRICA, S.L.U.', 'clasif': 'Estándar'}

    # 4.2 Reemplazar líneas Sep-26 en 'Detalle Líneas Clasificadas'
    # Guardar las filas que NO sean Sep-26
    non_sep_rows = []
    for r in range(2, ws_lines.max_row + 1):
        year_v = ws_lines.cell(r, 1).value
        mes_v = str(ws_lines.cell(r, 2).value or '')
        if year_v == 2026 and mes_v.lower().startswith('sep'):
            continue
        row_vals = [ws_lines.cell(r, c).value for c in range(1, len(header_l) + 1)]
        non_sep_rows.append(row_vals)

    print(f"Líneas no Sep-26 conservadas: {len(non_sep_rows)}")

    # Limpiar hoja y reconstruir
    wb_b.remove(ws_lines)
    ws_new_lines = wb_b.create_sheet('Detalle Líneas Clasificadas', 0)
    ws_new_lines.append(header_l)

    # Añadir filas previas (Ene-Ago 2026 y todo 2027)
    for r_data in non_sep_rows:
        ws_new_lines.append(r_data)

    # Añadir nuevas líneas reales de Septiembre 2026
    for s in sep_lines:
        cn = norm(s['cliente'])
        m_info = client_master.get(cn)
        if not m_info:
            for k_m, val_m in client_master.items():
                if k_m in cn or cn in k_m:
                    m_info = val_m
                    break
        com_val = m_info['comercial'] if m_info else 'General'
        pais_val = m_info['pais'] if m_info else 'ESPAÑA'
        cli_val = m_info['cliente'] if m_info else s['cliente']
        clasif_val = m_info['clasif'] if m_info else 'Estándar'

        # Tamaño de envase (extraer del final del SKU)
        env_val = 1
        sku_clean = s['sku']
        if sku_clean.endswith(('0020', '0200', '1000', '0005', '0001')):
            try:
                env_val = int(sku_clean[-4:])
            except: pass

        row_to_add = [
            2026,
            'Sep-26',
            com_val,
            pais_val,
            cli_val,
            clasif_val,
            s['sku'],
            s['desc'],
            'Maduro',
            'Esencial/Core',
            0.05,
            0.04,
            0.09,
            env_val,
            round(s['uds'], 2),
            round(s['pr_med'], 4),
            f"=O{ws_new_lines.max_row+1}*P{ws_new_lines.max_row+1}"
        ]
        ws_new_lines.append(row_to_add)

    print(f"Nuevas líneas reales de Sep-26 añadidas: {len(sep_lines)}")
    print(f"Total filas en 'Detalle Líneas Clasificadas': {ws_new_lines.max_row}")

    # Guardar libro
    wb_b.save(budget_file)
    print(f"Guardado exitoso de '{budget_file}'!")

    # 5. Sincronizar también con la copia original sin aplanar si existe
    src_orig = 'Presupuesto_Ventas_2027_Original_Sin_Aplanar.xlsx'
    if os.path.exists(src_orig):
        shutil.copy2(budget_file, src_orig)
        print(f"Sincronizado '{src_orig}'")

    print("\n=== 5. Reejecutando 'build_profitability_data.py' ===")
    from build_profitability_data import build_profitability_data
    build_profitability_data()
    print("\n¡Todo el proceso ha finalizado con éxito!")

if __name__ == '__main__':
    run_integration()
