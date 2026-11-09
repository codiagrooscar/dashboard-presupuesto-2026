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

import openpyxl
import json
import re
import shutil
import subprocess
import copy
from datetime import datetime
from collections import defaultdict as _dd

print("=== Generating Production Planning Data (Dual Scenario: Oficial vs Revisado) ===")

def safe_load_workbook(filepath, data_only=True, read_only=False):
    try:
        return openpyxl.load_workbook(filepath, data_only=data_only, read_only=read_only)
    except Exception:
        base_name = os.path.basename(filepath)
        temp_file = f"temp_safe_{base_name}"
        subprocess.run(['powershell', '-Command', f'Copy-Item -LiteralPath "{filepath}" -Destination "{temp_file}" -Force'], check=True)
        return openpyxl.load_workbook(temp_file, data_only=data_only, read_only=read_only)

def get_latest_stock_file():
    stock_dir = r'C:\Users\oscar.ocampo\OneDrive - Agroquímica Codiagro S.L\Escritorio\Estado de pedidos\Aplicacion web'
    if not os.path.exists(stock_dir):
        return None
    candidates = [f for f in os.listdir(stock_dir) if f.startswith('Stock') and f.endswith('.xlsx') and not f.startswith(('temp_', '~$'))]
    if not candidates:
        return None
    def sort_key(f):
        m = re.search(r'(\d{1,2})\.(\d{1,2})', f)
        if m:
            return (int(m.group(2)), int(m.group(1)))
        return (0, 0)
    candidates.sort(key=sort_key, reverse=True)
    return os.path.join(stock_dir, candidates[0])

def get_active_order_files():
    order_files = []
    for f in ['Pedidos02.11 comerciales.xlsx', 'Pedidos 02.11 comerciales.xlsx', 'Pedidos30.10 comerciales.xlsx']:
        if os.path.exists(f):
            order_files.append(f)
            break
        elif os.path.exists(os.path.join('Pedidos', f)):
            order_files.append(os.path.join('Pedidos', f))
            break
    candidates = [
        f for f in os.listdir('.')
        if f.startswith('Pedidos') and f.endswith('.xlsx') and not f.startswith(('temp_', '~$'))
    ]
    if os.path.exists('Pedidos') and os.path.isdir('Pedidos'):
        candidates += [
            os.path.join('Pedidos', f) for f in os.listdir('Pedidos')
            if f.startswith('Pedidos') and f.endswith('.xlsx') and not f.startswith(('temp_', '~$'))
        ]
    nov_candidates = []
    for f in candidates:
        base_f = os.path.basename(f)
        m = re.search(r'(\d{1,2})\.11', base_f)
        if m:
            day = int(m.group(1))
            if day >= 3:
                nov_candidates.append((day, f))
    if nov_candidates:
        nov_candidates.sort(key=lambda x: x[0], reverse=True)
        order_files.append(nov_candidates[0][1])
    elif os.path.exists('Pedidos03.11 comerciales.xlsx'):
        order_files.append('Pedidos03.11 comerciales.xlsx')
    return list(dict.fromkeys(order_files))

def detect_current_cutoff_month(order_files, default="2026-11"):
    latest_date = None
    for f in order_files:
        base_f = os.path.basename(f)
        m = re.search(r'Pedidos\s*(\d{1,2})\.(\d{1,2})', base_f, re.IGNORECASE)
        if m:
            day = int(m.group(1))
            mo = int(m.group(2))
            yr = 2026 if mo in [10, 11, 12] else 2027
            d_val = (yr, mo, day)
            if latest_date is None or d_val > latest_date:
                latest_date = d_val
    if latest_date:
        return f"{latest_date[0]}-{latest_date[1]:02d}"
    return default

# 1. Load Stock from latest stock file
stock_path = get_latest_stock_file() or r'C:\Users\oscar.ocampo\OneDrive - Agroquímica Codiagro S.L\Escritorio\Estado de pedidos\Aplicacion web\Stock 02.10.xlsx'
print(f"Loading stock from: {stock_path}")
wb_stock = safe_load_workbook(stock_path, data_only=True)
ws_stock = wb_stock.active

stock_map = {}
granel_map = {}
envasado_map = {}

for row in ws_stock.iter_rows(min_row=4, values_only=True):
    fam = row[2]
    code = row[3]
    desc = row[4]
    qty = row[5]
    price = row[6]
    if code:
        c_clean = str(code).strip().upper()
        item_info = {
            'codigo': c_clean,
            'desc': str(desc or '').strip(),
            'stock': float(qty or 0),
            'precio': float(price or 0),
            'familia': int(fam) if fam is not None else None
        }
        stock_map[c_clean] = item_info
        if fam in [38, 39, 40]:
            granel_map[c_clean] = item_info
        else:
            envasado_map[c_clean] = item_info

print(f"Loaded {len(stock_map)} stock items ({len(granel_map)} Graneles = {sum(x['stock'] for x in granel_map.values()):,.0f} L/Kg, {len(envasado_map)} Envasados = {sum(x['stock'] for x in envasado_map.values()):,.0f} Uds).")

# 2. Load Pending Orders from dashboard_data.json and active orders files
active_order_files = get_active_order_files()
pending_orders = {}

dash_src = get_source('web_dashboard/dashboard_data.json')
if os.path.exists(dash_src):
    try:
        with open(dash_src, 'r', encoding='utf-8') as f:
            dash_data = json.load(f)
        p_sep = dash_data.get('periods', {}).get('2026-09', {})
        for com, cdata in p_sep.get('comerciales', {}).items():
            for l in cdata.get('lineas', []):
                pend = float(l.get('pendientes_uds', 0))
                if pend > 0:
                    sku = str(l.get('sku', '')).strip().upper()
                    if sku not in pending_orders:
                        pending_orders[sku] = {
                            'sku': sku,
                            'desc': l.get('descripcion', ''),
                            'pendientes_uds': 0.0,
                            'is_nacional': bool(l.get('is_nacional')),
                            'clientes': set()
                        }
                    pending_orders[sku]['pendientes_uds'] += pend
                    if l.get('cliente'):
                        pending_orders[sku]['clientes'].add(str(l.get('cliente')).strip())
    except Exception as e:
        print("Aviso leyendo dashboard_data.json:", e)

seen_pending_lines = set()
for ord_file in active_order_files:
    try:
        wb_ped = safe_load_workbook(ord_file, data_only=True)
        ws_ped = wb_ped.active
        for row in ws_ped.iter_rows(min_row=2, values_only=True):
            p_uds = row[10]
            if p_uds and float(p_uds) > 0:
                ped_num = str(row[3] or '').strip()
                c_art = str(row[6] or '').strip().upper()
                d_art = str(row[7] or '').strip()
                cli = str(row[5] or '').strip()
                line_key = (ped_num, c_art)
                if line_key in seen_pending_lines:
                    continue
                seen_pending_lines.add(line_key)
                if c_art not in pending_orders:
                    pending_orders[c_art] = {
                        'sku': c_art,
                        'desc': d_art,
                        'pendientes_uds': 0.0,
                        'is_nacional': True,
                        'clientes': set()
                    }
                pending_orders[c_art]['pendientes_uds'] += float(p_uds)
                if cli:
                    pending_orders[c_art]['clientes'].add(cli)
    except Exception as e:
        print(f"Aviso leyendo {ord_file}:", e)

for s in pending_orders:
    pending_orders[s]['clientes'] = sorted(list(pending_orders[s]['clientes']))

print(f"Loaded {len(pending_orders)} pending order SKUs. Total pending units: {sum(x['pendientes_uds'] for x in pending_orders.values()):,.0f}")

# 3. Setup Temporal Horizon & Dynamic Month N
month_keys = [
    '2026-10', '2026-11', '2026-12',
    '2027-01', '2027-02', '2027-03', '2027-04', '2027-05', '2027-06',
    '2027-07', '2027-08', '2027-09', '2027-10', '2027-11', '2027-12'
]
month_col_idx = [27, 29, 31, 35, 37, 39, 41, 43, 45, 47, 49, 51, 53, 55, 57]

current_month = detect_current_cutoff_month(active_order_files, default="2026-11")
print(f"Active Dynamic Cut-Off Month (N): {current_month}")

idx_n = month_keys.index(current_month) if current_month in month_keys else 1

presets = {
    "N": month_keys[idx_n : idx_n + 1],
    "N1": month_keys[idx_n : idx_n + 2],
    "N2": month_keys[idx_n : idx_n + 3],
    "N5": month_keys[idx_n : idx_n + 6],
    "REST_2026": [m for m in month_keys if m >= current_month and m.startswith('2026')],
    "Q1_2027": ["2027-01", "2027-02", "2027-03"],
    "H1_2027": ["2027-01", "2027-02", "2027-03", "2027-04", "2027-05", "2027-06"],
    "ALL_2027": [m for m in month_keys if m.startswith('2027')],
    "ALL_15": month_keys
}

es_months = {1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun", 7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic"}
en_months = {1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun", 7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"}

months_metadata = []
for m_key in month_keys:
    y_str, mo_str = m_key.split('-')
    y_int = int(y_str)
    mo_int = int(mo_str)
    yy = y_str[2:]
    if m_key < current_month:
        label_es = f"{es_months[mo_int]} {yy} (Cerrado)"
        label_en = f"{en_months[mo_int]} {yy} (Closed)"
    elif m_key == current_month:
        label_es = f"{es_months[mo_int]} {yy} (N)"
        label_en = f"{en_months[mo_int]} {yy} (N)"
    else:
        label_es = f"{es_months[mo_int]} {yy}"
        label_en = f"{en_months[mo_int]} {yy}"
    months_metadata.append({
        "key": m_key,
        "label_es": label_es,
        "label_en": label_en,
        "year": y_int,
        "month": mo_int
    })

def format_preset_label(m_list):
    if not m_list: return ""
    if len(m_list) == 1:
        y, mo = m_list[0].split('-')
        return f"{es_months[int(mo)]} {y[2:]}"
    y1, mo1 = m_list[0].split('-')
    y2, mo2 = m_list[-1].split('-')
    if y1 == y2:
        return f"{es_months[int(mo1)]}-{es_months[int(mo2)]} {y1[2:]}"
    else:
        return f"{es_months[int(mo1)]} {y1[2:]} - {es_months[int(mo2)]} {y2[2:]}"

preset_labels = {
    "N": f"N ({format_preset_label(presets['N'])})",
    "N1": f"N+1 ({format_preset_label(presets['N1'])})",
    "N2": f"N+2 ({format_preset_label(presets['N2'])})",
    "N5": f"N+5 ({format_preset_label(presets['N5'])})",
    "REST_2026": f"Resto 2026 ({format_preset_label(presets['REST_2026'])})",
    "Q1_2027": "Q1 2027 (Ene-Mar)",
    "H1_2027": "1er Semestre 2027",
    "ALL_2027": "Todo 2027 (12 meses)",
    "ALL_15": "Todo el Horizonte (15 meses)"
}

country_code_map = {
    'TURQUIA': 'TR', 'COLOMBIA': 'CO', 'PERU': 'PE', 'EGIPTO': 'EG',
    'MARRUECOS': 'MA', 'GRECIA': 'GR', 'ECUADOR': 'EC', 'PORTUGAL': 'PT',
    'CHILE': 'CL', 'ARGELIA': 'DZ', 'MEXICO': 'MX', 'ITALIA': 'IT',
    'JORDANIA': 'JO', 'LIBANO': 'LB', 'TUNEZ': 'TN', 'GUATEMALA': 'GT'
}

# 4. Load Forecast Sources: Oficial (Presupuesto_Ventas_2027_Definitivo.xlsx) and Revisado (Revision_Previsiones_Q4_2026_CONSOLIDADO.xlsx)
wb_prev = safe_load_workbook('Presupuesto_Ventas_2027_Definitivo.xlsx', read_only=False, data_only=True)
ws_prev = wb_prev['Previsión Matriz Horizontal']

# Cache historical real sales LTM from Consolidado_Ventas_y_Previsiones_2026.xlsx
real_sales_by_sku_cli = _dd(lambda: {'vol': 0.0, 'eur': 0.0, 'pais': '', 'comercial': ''})
real_sales_by_granel_cli = _dd(lambda: {'vol': 0.0, 'eur': 0.0, 'pais': '', 'comercial': ''})

if os.path.exists('Consolidado_Ventas_y_Previsiones_2026.xlsx'):
    try:
        wb_c = openpyxl.load_workbook('Consolidado_Ventas_y_Previsiones_2026.xlsx', read_only=True)
        if 'Detalle Líneas 2026' in wb_c.sheetnames:
            ws_det = wb_c['Detalle Líneas 2026']
            for row in ws_det.iter_rows(min_row=2, values_only=True):
                per = str(row[0] or '')
                if 'REAL' not in per.upper(): continue
                com = str(row[1] or '').strip()
                pais = str(row[2] or '').strip().upper()
                cli = str(row[3] or '').strip()
                sku = str(row[4] or '').strip().upper()
                u = float(row[7] or 0)
                e = float(row[8] or 0)
                if not sku or not cli or u <= 0: continue
                if 'CODIMED' in cli.upper(): continue
                m_b = re.match(r'^([A-Z0-9\-\(\)]+?\d{4})([A-Z]{2})$', sku)
                b_sku = m_b.group(1) if m_b else sku
                k = (b_sku, cli)
                real_sales_by_sku_cli[k]['vol'] += u
                real_sales_by_sku_cli[k]['eur'] += e
                real_sales_by_sku_cli[k]['pais'] = pais
                real_sales_by_sku_cli[k]['comercial'] = com
    except Exception as _e_sales:
        print(f"Aviso cargando ventas reales de consolidado: {_e_sales}")

# Carga de la Revisión Q4 2026 de los Comerciales (Revision_Previsiones_Q4_2026_CONSOLIDADO.xlsx)
rev_by_sku = _dd(lambda: {'oct': 0.0, 'nov': 0.0, 'dic': 0.0, 'is_nacional': True, 'desc': '', 'clientes': set(), 'comerciales': set()})
rev_lines_list = []

candidates_rev = [
    'Revision_Previsiones_Q4_2026_CONSOLIDADO.xlsx',
    os.path.join('RevQ4Noviembre', 'Revision_Previsiones_Q4_2026_CONSOLIDADO.xlsx'),
    os.path.join('antiguos', 'Revision_Previsiones_Q4_2026', 'Revision_Previsiones_Q4_2026_CONSOLIDADO.xlsx')
]
f_path_rev = next((f for f in candidates_rev if os.path.exists(f)), None)

if f_path_rev:
    wb_rev = safe_load_workbook(f_path_rev, data_only=True)
    for com in ['Alfonso', 'García', 'Irene', 'Javier', 'Mehmet', 'Pedro', 'Ricardo']:
        if com not in wb_rev.sheetnames: continue
        ws = wb_rev[com]
        for r in range(10, ws.max_row + 1):
            pais = str(ws.cell(r, 2).value or '').strip().upper()
            cli = str(ws.cell(r, 3).value or '').strip()
            sku_raw = str(ws.cell(r, 4).value or '').strip().upper()
            desc = str(ws.cell(r, 5).value or '').strip()
            oct_u = ws.cell(r, 6).value
            nov_u = ws.cell(r, 12).value
            dic_u = ws.cell(r, 14).value
            if not cli or not sku_raw: continue
            is_nac = (pais in ['ESPAÑA', 'ESP', ''])
            has_suffix = bool(re.search(r'[A-Z]{2}$', sku_raw)) and len(sku_raw) >= 8
            if not is_nac and not has_suffix:
                ctry_code = country_code_map.get(pais, pais[:2] if len(pais) >= 2 else 'EX')
                sku = f"{sku_raw}{ctry_code}"
            else:
                sku = sku_raw
            try: nov_v = float(nov_u or 0)
            except: nov_v = 0.0
            try: dic_v = float(dic_u or 0)
            except: dic_v = 0.0
            try: oct_v = float(oct_u or 0)
            except: oct_v = 0.0

            rev_by_sku[sku]['oct'] += oct_v
            rev_by_sku[sku]['nov'] += nov_v
            rev_by_sku[sku]['dic'] += dic_v
            rev_by_sku[sku]['is_nacional'] = is_nac
            if desc: rev_by_sku[sku]['desc'] = desc
            rev_by_sku[sku]['clientes'].add(cli)
            rev_by_sku[sku]['comerciales'].add(com)

            m_b = re.match(r'^([A-Z0-9\-\(\)]+?\d{4})([A-Z]{2})$', sku)
            b_sku = m_b.group(1) if m_b else sku
            rev_lines_list.append({
                'comercial': com,
                'pais': pais,
                'is_nacional': is_nac,
                'cliente': cli,
                'sku': sku,
                'base_sku': b_sku,
                'desc': desc,
                'oct_26': oct_v,
                'nov_26': nov_v,
                'dic_26': dic_v
            })

print(f"Revision Q4 loaded: {len(rev_by_sku)} SKUs, Nov sum = {sum(x['nov'] for x in rev_by_sku.values()):,.0f} u, Dic sum = {sum(x['dic'] for x in rev_by_sku.values()):,.0f} u")

# Helper functions for packaging and resolution
def resolve_envasado_stock(base_sku, sku):
    if sku in envasado_map:
        return envasado_map[sku]['stock'], envasado_map[sku]['desc'], 'exact_envasado'
    if base_sku in envasado_map:
        return envasado_map[base_sku]['stock'], envasado_map[base_sku]['desc'], 'base_envasado'
    if sku in stock_map and stock_map[sku].get('familia') not in [38, 39, 40]:
        return stock_map[sku]['stock'], stock_map[sku]['desc'], 'exact'
    if base_sku in stock_map and stock_map[base_sku].get('familia') not in [38, 39, 40]:
        return stock_map[base_sku]['stock'], stock_map[base_sku]['desc'], 'base'
    return 0.0, '', 'none'

def find_granel_smart(base_sku, desc=''):
    sku = base_sku.strip().upper()
    m = re.match(r'^(.+?)(0001|0005|0020|0200|1000|0008)(\S*)$', sku)
    if m:
        root, pkg, sfx = m.group(1), m.group(2), m.group(3)
        if root in granel_map:
            return granel_map[root]
        if root + sfx in granel_map:
            return granel_map[root + sfx]
        if root.startswith('AN') or root == 'A' or root == 'AO':
            if 'A' in granel_map: return granel_map['A']
        if root == 'S' and 'S' in granel_map:
            return granel_map['S']
        if root == 'SS' and 'SS' in granel_map:
            return granel_map['SS']
        if root == 'PHD' and 'PHD00' in granel_map:
            return granel_map['PHD00']
        if root == 'GRE1' and 'GRE1' in granel_map:
            return granel_map['GRE1']
        for end_sfx in ['AE', 'EG', 'CR', 'CO', 'TR', 'P']:
            if root.endswith(end_sfx) and root[:-len(end_sfx)] in granel_map:
                return granel_map[root[:-len(end_sfx)]]
        r_norm = re.sub(r'[\-\(\)\s]', '', root)
        for g_cod, g in granel_map.items():
            g_norm = re.sub(r'[\-\(\)\s]', '', g_cod)
            if r_norm == g_norm:
                return g
    elif sku in granel_map:
        return granel_map[sku]
    return None

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
        if l_val >= 500:
            return f"Big Bag {int(l_val) if l_val.is_integer() else l_val} Kg"
        elif l_val == int(l_val):
            return f"Bolsas {int(l_val)} Kg"
        else:
            return f"Bolsas {l_val} Kg"
    else:
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

# Catálogo Auditado de Regalos Técnicos Verificados (Stock real excedentario en almacén)
CATALOGO_REGALOS = {
    'BIO0005': {'sku': 'BIO0005', 'desc': 'BIORAD-20, 5 L', 'formato': 'Garrafas 5L', 'envase_l': 5.0, 'coste_unit': 0.7328, 'stock_disponible': 1840.0, 'sinergia': 'Enraizamiento y Asimilación'},
    'AN0005': {'sku': 'AN0005', 'desc': 'ALCAPLANT NEW, 5 KG', 'formato': 'Bolsas 5 Kg', 'envase_l': 5.0, 'coste_unit': 0.2191, 'stock_disponible': 1660.0, 'sinergia': 'Corrección de Calcio en Suelo'},
    'AGFEL0001': {'sku': 'AGFEL0001', 'desc': 'AGROXILATO-FE-L, 1 L', 'formato': 'Botellas 1L', 'envase_l': 1.0, 'coste_unit': 1.6931, 'stock_disponible': 732.0, 'sinergia': 'Corrección Rápida Clorosis Férrica'},
    'AGRMG0005': {'sku': 'AGRMG0005', 'desc': 'AGROXIGREEN-MG, 5 L', 'formato': 'Garrafas 5L', 'envase_l': 5.0, 'coste_unit': 1.2216, 'stock_disponible': 775.0, 'sinergia': 'Activación Fotosintética Magnesio'},
    'BROTA0005': {'sku': 'BROTA0005', 'desc': 'BROTAMEC, 5 L', 'formato': 'Garrafas 5L', 'envase_l': 5.0, 'coste_unit': 1.2668, 'stock_disponible': 1100.0, 'sinergia': 'Brotación y Crecimiento Vegetativo'},
    'BR-590005': {'sku': 'BR-590005', 'desc': 'BR-59, 5 KG', 'formato': 'Bolsas 5 Kg', 'envase_l': 5.0, 'coste_unit': 0.5247, 'stock_disponible': 1760.0, 'sinergia': 'Desbloqueo Salino y Fertilidad'},
    'AQUOM0005': {'sku': 'AQUOM0005', 'desc': 'AQUOM, 5 L', 'formato': 'Garrafas 5L', 'envase_l': 5.0, 'coste_unit': 2.3703, 'stock_disponible': 850.0, 'sinergia': 'Regulación Infiltración Hídrica'},
    'PT0005': {'sku': 'PT0005', 'desc': 'AGROPOTASION, 5 L', 'formato': 'Garrafas 5L', 'envase_l': 5.0, 'coste_unit': 1.1395, 'stock_disponible': 880.0, 'sinergia': 'Potasio Puro para Maduración'},
    'B0005': {'sku': 'B0005', 'desc': 'BOROCAL, 5 L', 'formato': 'Garrafas 5L', 'envase_l': 5.0, 'coste_unit': 0.9928, 'stock_disponible': 475.0, 'sinergia': 'Floración, Cuajado y Movilidad Boro'},
    'FU0001': {'sku': 'FU0001', 'desc': 'FULVIGIN, 1 L', 'formato': 'Botellas 1L', 'envase_l': 1.0, 'coste_unit': 1.4470, 'stock_disponible': 588.0, 'sinergia': 'Ácidos Fúlvicos y Complejación'}
}

def _get_verified_gift(b_sku, desc, g_code):
    text = f"{b_sku} {desc} {g_code or ''}".upper()
    if any(k in text for k in ['ALCAPLANT', 'SAL-WAX', 'SALWAX', 'CALCI', 'BITTER', 'SS']):
        return CATALOGO_REGALOS['BIO0005']
    if any(k in text for k in ['RADISOL', 'START', 'PGR', 'FLORAMEC', 'AVANCE', 'SPRINTENE', 'CODIBROT']):
        return CATALOGO_REGALOS['B0005']
    if any(k in text for k in ['AGROXILATO', 'CODIFOL', 'FERRO', 'FE-L', 'ZINC', 'MN', 'C-M-P']):
        return CATALOGO_REGALOS['AN0005']
    if any(k in text for k in ['AGROWHITE', 'FROST', 'STRESS', 'AQUOM', 'FITOKONTROL']):
        return CATALOGO_REGALOS['BR-590005']
    if any(k in text for k in ['POTAS', 'PT', 'K', 'MADUR']):
        return CATALOGO_REGALOS['AGRMG0005']
    return CATALOGO_REGALOS['BROTA0005']

def _get_complementary_product(b_sku, desc, g_code):
    g = _get_verified_gift(b_sku, desc, g_code)
    return f"Obsequio Técnico: {g['desc']} ({g['sku']}) - {g['sinergia']} (Coste real: {g['coste_unit']:.2f}€/L-Kg)"

# Official Balance / P&L constants (Inputs CEO / CFO)
INVENTARIO_TOTAL_BALANCE_EUR = 1323814.0
COGS_LTM_TOTAL_EUR = 4935072.0
VENTAS_NETAS_LTM_EUR = 13780257.0
CLIENTES_BALANCE_EUR = 3907712.0
PROVEEDORES_BALANCE_EUR = 871084.0

COGS_DIARIO_TOTAL = COGS_LTM_TOTAL_EUR / 365.0      # 13.520,745 €/día
VENTAS_DIARIAS_LTM = VENTAS_NETAS_LTM_EUR / 365.0   # 37.754,129 €/día

DIO_BASE_OFICIAL = round((INVENTARIO_TOTAL_BALANCE_EUR / COGS_LTM_TOTAL_EUR) * 365.0, 1)  # 97.9 días (~98d)
DSO_BASE_OFICIAL = round((CLIENTES_BALANCE_EUR / VENTAS_NETAS_LTM_EUR) * 365.0, 1)         # 103.5 días
DPO_BASE_OFICIAL = round((PROVEEDORES_BALANCE_EUR / COGS_LTM_TOTAL_EUR) * 365.0, 1)       # 64.4 días
CCC_BASE_OFICIAL = round(DIO_BASE_OFICIAL + DSO_BASE_OFICIAL - DPO_BASE_OFICIAL, 1)        # 137.0 días

COGS_DIARIO_PT_NAC = 4068.49  # Solo Almacén PT Nacional Planta

# 5. Pipeline Builder for a specific Scenario ('oficial' or 'revisado')
def build_scenario_production_data(scenario_name='revisado'):
    print(f"\n--- Building Scenario: {scenario_name.upper()} ---")

    # A. Build Detailed SKUs
    detailed_skus = {}
    ws_prev_entries = []

    for row in ws_prev.iter_rows(min_row=2, values_only=True):
        comercial = str(row[0] or '').strip()
        if not comercial or 'TOTAL' in comercial.upper():
            continue
        pais = str(row[1] or '').strip().upper()
        cliente = str(row[2] or '').strip()
        codigo = row[4]
        desc = str(row[5] or '').strip()
        envase = str(row[6] or '').strip()
        pvp_col = row[7]
        try: pvp_val = float(pvp_col) if pvp_col is not None else 0.0
        except: pvp_val = 0.0

        if not codigo:
            continue

        raw_sku = str(codigo).strip().upper()
        is_nacional = (pais in ['ESPAÑA', 'ESP', ''])

        has_suffix = bool(re.search(r'[A-Z]{2}$', raw_sku)) and len(raw_sku) >= 8
        if not is_nacional and not has_suffix:
            ctry_code = country_code_map.get(pais, pais[:2] if len(pais) >= 2 else 'EX')
            sku = f"{raw_sku}{ctry_code}"
        else:
            sku = raw_sku

        m_base = re.match(r'^([A-Z0-9\-\(\)]+?\d{4})([A-Z]{2})$', sku)
        base_sku = m_base.group(1) if m_base else sku

        if sku not in detailed_skus:
            detailed_skus[sku] = {
                'sku': sku,
                'base_sku': base_sku,
                'desc': desc,
                'envase': envase,
                'is_nacional': is_nacional,
                'pais': pais if not is_nacional else 'ESPAÑA',
                'ambito': 'Nacional' if is_nacional else 'Exportación',
                'monthly_forecast': {m: 0.0 for m in month_keys},
                'clientes': set(),
                'comerciales': set()
            }

        if cliente:
            detailed_skus[sku]['clientes'].add(cliente)
        if comercial:
            detailed_skus[sku]['comerciales'].add(comercial)

        vol_prev_total = 0.0
        eur_prev_total = 0.0
        for m_key, c_idx in zip(month_keys, month_col_idx):
            if c_idx < len(row):
                val = row[c_idx]
                try:
                    if val:
                        v_flt = float(val)
                        detailed_skus[sku]['monthly_forecast'][m_key] += v_flt
                        vol_prev_total += v_flt
                except:
                    pass
            if c_idx + 1 < len(row):
                e_val = row[c_idx + 1]
                try:
                    if e_val: eur_prev_total += float(e_val)
                except:
                    pass

        if cliente and 'CODIMED' not in cliente.upper():
            pvp_calc = (eur_prev_total / vol_prev_total) if (vol_prev_total > 0 and eur_prev_total > 0) else pvp_val
            ws_prev_entries.append({
                'comercial': comercial,
                'pais': pais if pais else 'ESPAÑA',
                'is_nacional': is_nacional,
                'cliente': cliente,
                'codigo': raw_sku,
                'sku': sku,
                'base_sku': base_sku,
                'vol': vol_prev_total,
                'eur': eur_prev_total,
                'pvp': pvp_calc
            })

    # If scenario is 'revisado', substitute Q4 (2026-10, 2026-11, 2026-12) with commercial revision
    if scenario_name == 'revisado' and rev_by_sku:
        for sku, d in detailed_skus.items():
            d['monthly_forecast']['2026-10'] = 0.0
            d['monthly_forecast']['2026-11'] = 0.0
            d['monthly_forecast']['2026-12'] = 0.0

        for r_sku, r_info in rev_by_sku.items():
            if r_sku in detailed_skus:
                d = detailed_skus[r_sku]
                d['monthly_forecast']['2026-10'] = r_info['oct']
                d['monthly_forecast']['2026-11'] = r_info['nov']
                d['monthly_forecast']['2026-12'] = r_info['dic']
                d['clientes'].update(r_info['clientes'])
                d['comerciales'].update(r_info['comerciales'])
            else:
                m_base = re.match(r'^([A-Z0-9\-\(\)]+?\d{4})([A-Z]{2})$', r_sku)
                base_sku = m_base.group(1) if m_base else r_sku
                detailed_skus[r_sku] = {
                    'sku': r_sku,
                    'base_sku': base_sku,
                    'desc': r_info['desc'],
                    'envase': '',
                    'is_nacional': r_info['is_nacional'],
                    'pais': 'ESPAÑA' if r_info['is_nacional'] else 'EXPORT',
                    'ambito': 'Nacional' if r_info['is_nacional'] else 'Exportación',
                    'monthly_forecast': {m: 0.0 for m in month_keys},
                    'clientes': set(r_info['clientes']),
                    'comerciales': set(r_info['comerciales'])
                }
                detailed_skus[r_sku]['monthly_forecast']['2026-10'] = r_info['oct']
                detailed_skus[r_sku]['monthly_forecast']['2026-11'] = r_info['nov']
                detailed_skus[r_sku]['monthly_forecast']['2026-12'] = r_info['dic']

    # Integrate Pending Orders not in forecast
    for p_sku, p_data in pending_orders.items():
        if p_sku not in detailed_skus:
            m_base = re.match(r'^([A-Z0-9\-\(\)]+?\d{4})([A-Z]{2})$', p_sku)
            base_sku = m_base.group(1) if m_base else p_sku
            detailed_skus[p_sku] = {
                'sku': p_sku,
                'base_sku': base_sku,
                'desc': p_data['desc'],
                'envase': '',
                'is_nacional': p_data['is_nacional'],
                'pais': 'ESPAÑA' if p_data['is_nacional'] else 'EXPORT',
                'ambito': 'Nacional' if p_data['is_nacional'] else 'Exportación',
                'monthly_forecast': {m: 0.0 for m in month_keys},
                'clientes': set(p_data['clientes']),
                'comerciales': set()
            }

    for sku, data in detailed_skus.items():
        p_info = pending_orders.get(sku, {})
        data['pendientes_uds'] = p_info.get('pendientes_uds', 0.0)
        data['clientes'] = sorted(list(data['clientes']))
        data['comerciales'] = sorted(list(data['comerciales']))

    # B. Build Physical Base SKUs (Packaging Level)
    base_skus_dict = {}

    for sku, d in detailed_skus.items():
        b_sku = d['base_sku']
        if b_sku not in base_skus_dict:
            stock_qty, stock_desc, match_type = resolve_envasado_stock(b_sku, sku)
            g_info = find_granel_smart(b_sku, d['desc'])
            g_cod = g_info['codigo'] if g_info else None
            g_desc = g_info['desc'] if g_info else None
            g_stock = g_info['stock'] if g_info else 0.0
            g_fam = g_info['familia'] if g_info else None

            env_l = parse_envase_litros(d['envase'], d['desc'] or stock_desc)
            fmt_label = get_packaging_format(d['envase'], d['desc'] or stock_desc, b_sku, g_fam)
            precio_val = stock_map.get(b_sku, {}).get('precio', 0.0) or stock_map.get(sku, {}).get('precio', 0.0)
            is_solid = is_kg_product(d['desc'] or stock_desc, b_sku, g_fam)

            base_skus_dict[b_sku] = {
                'base_sku': b_sku,
                'desc': d['desc'] or stock_desc,
                'envase': d['envase'],
                'envase_litros': env_l,
                'unidad_medida': 'Kg' if is_solid else 'L',
                'formato_label': fmt_label,
                'precio_unitario': precio_val,
                'stock_actual': stock_qty,
                'stock_envasado': stock_qty,
                'stock_envases': round(stock_qty / env_l, 1) if env_l > 0 else stock_qty,
                'stock_litros': stock_qty,
                'stock_match_type': match_type,
                'granel_code': g_cod,
                'granel_desc': g_desc,
                'granel_stock': g_stock,
                'granel_family': g_fam,
                'pedidos_pendientes': 0.0,
                'pedidos_pendientes_nacional': 0.0,
                'pedidos_pendientes_export': 0.0,
                'monthly_forecast': {m: 0.0 for m in month_keys},
                'monthly_forecast_nacional': {m: 0.0 for m in month_keys},
                'monthly_forecast_export': {m: 0.0 for m in month_keys},
                'destinos': [],
                'has_nacional': False,
                'has_export': False
            }

        b_item = base_skus_dict[b_sku]
        p_uds = float(d.get('pendientes_uds', 0) or 0)
        b_item['pedidos_pendientes'] += p_uds
        if d['is_nacional']:
            b_item['has_nacional'] = True
            b_item['pedidos_pendientes_nacional'] += p_uds
        else:
            b_item['has_export'] = True
            b_item['pedidos_pendientes_export'] += p_uds

        for m in month_keys:
            f_val = float(d['monthly_forecast'].get(m, 0) or 0)
            b_item['monthly_forecast'][m] += f_val
            if d['is_nacional']:
                b_item['monthly_forecast_nacional'][m] += f_val
            else:
                b_item['monthly_forecast_export'][m] += f_val

        b_item['destinos'].append({
            'sku': d['sku'],
            'desc': d['desc'],
            'pais': d['pais'],
            'ambito': d['ambito'],
            'is_nacional': d['is_nacional'],
            'fabricacion_tipo': 'Para Stock / Previsión Nacional' if d['is_nacional'] else 'Sobre Pedido (Exportación)',
            'pendientes_uds': d['pendientes_uds'],
            'monthly_forecast': d['monthly_forecast']
        })

    for b_sku, b_item in base_skus_dict.items():
        b_item['sub_items'] = b_item['destinos']

    # C. Calculate Rotation (SlowMovers / NoMovers on Free Stock)
    n2_months = presets['N2'] # ['2026-11', '2026-12', '2027-01']
    n5_months = presets['N5'] # 6 dynamic months N+5
    mid_months = [m for m in n5_months if m not in n2_months]

    for b_sku, b_item in base_skus_dict.items():
        stock_u = b_item['stock_actual']
        pend_total = b_item['pedidos_pendientes']
        env_l = b_item['envase_litros']
        p_eur = b_item['precio_unitario']

        stock_comprometido = min(stock_u, pend_total) if stock_u > 0 else 0.0
        stock_libre = max(0.0, stock_u - pend_total) if stock_u > 0 else 0.0
        b_item['stock_comprometido_pedidos'] = round(stock_comprometido, 2)
        b_item['stock_libre_disponible'] = round(stock_libre, 2)

        d_nac_n2 = sum(b_item['monthly_forecast_nacional'].get(m, 0.0) for m in n2_months)
        d_nac_mid = sum(b_item['monthly_forecast_nacional'].get(m, 0.0) for m in mid_months)
        d_nac_6m = sum(b_item['monthly_forecast_nacional'].get(m, 0.0) for m in n5_months)

        no_mover_u = max(0.0, stock_libre - d_nac_6m) if stock_libre > 0 else 0.0
        s_rem_n2 = max(0.0, stock_libre - d_nac_n2) if stock_libre > 0 else 0.0
        slow_mover_u = min(s_rem_n2, d_nac_mid) if stock_libre > 0 else 0.0
        fast_mover_u = min(stock_libre, d_nac_n2) if stock_libre > 0 else 0.0

        b_item['demanda_nac_n2'] = round(d_nac_n2, 2)
        b_item['demanda_nac_q4'] = round(d_nac_n2, 2)
        b_item['demanda_nac_mid'] = round(d_nac_mid, 2)
        b_item['demanda_nac_q1_27'] = round(d_nac_mid, 2)
        b_item['demanda_nac_6m'] = round(d_nac_6m, 2)
        b_item['no_mover_qty'] = round(no_mover_u, 2)
        b_item['slow_mover_qty'] = round(slow_mover_u, 2)
        b_item['fast_mover_qty'] = round(fast_mover_u, 2)
        b_item['no_mover_litros'] = round(no_mover_u, 2)
        b_item['slow_mover_litros'] = round(slow_mover_u, 2)
        b_item['fast_mover_litros'] = round(fast_mover_u, 2)
        b_item['no_mover_envases'] = round(no_mover_u / env_l, 1) if env_l > 0 else no_mover_u
        b_item['slow_mover_envases'] = round(slow_mover_u / env_l, 1) if env_l > 0 else slow_mover_u
        b_item['fast_mover_envases'] = round(fast_mover_u / env_l, 1) if env_l > 0 else fast_mover_u
        b_item['no_mover_eur'] = round(no_mover_u * p_eur, 2)
        b_item['slow_mover_eur'] = round(slow_mover_u * p_eur, 2)
        b_item['fast_mover_eur'] = round(fast_mover_u * p_eur, 2)
        b_item['stock_eur'] = round(stock_u * p_eur, 2)
        b_item['is_nomover'] = no_mover_u > 0
        b_item['is_slowmover'] = slow_mover_u > 0
        b_item['is_fastmover'] = stock_u > 0 and no_mover_u == 0 and slow_mover_u == 0

    for sku, d in detailed_skus.items():
        b_sku = d['base_sku']
        b_item = base_skus_dict[b_sku]
        d['pedidos_pendientes'] = d.get('pendientes_uds', 0.0)
        d['stock_actual'] = b_item['stock_actual']
        d['stock_envasado'] = b_item['stock_envasado']
        d['stock_litros'] = b_item['stock_litros']
        d['stock_envases'] = b_item.get('stock_envases', 0)
        d['stock_match_type'] = b_item['stock_match_type']
        d['granel_code'] = b_item['granel_code']
        d['granel_desc'] = b_item['granel_desc']
        d['granel_stock'] = b_item['granel_stock']
        d['granel_family'] = b_item['granel_family']
        d['envase_litros'] = b_item['envase_litros']
        d['unidad_medida'] = b_item.get('unidad_medida', 'L')
        d['formato_label'] = b_item['formato_label']
        d['precio_unitario'] = b_item['precio_unitario']
        d['demanda_nac_n2'] = b_item['demanda_nac_n2']
        d['demanda_nac_q4'] = b_item['demanda_nac_q4']
        d['demanda_nac_6m'] = b_item['demanda_nac_6m']
        d['no_mover_qty'] = b_item['no_mover_qty']
        d['slow_mover_qty'] = b_item['slow_mover_qty']
        d['fast_mover_qty'] = b_item['fast_mover_qty']
        d['no_mover_litros'] = b_item['no_mover_litros']
        d['slow_mover_litros'] = b_item['slow_mover_litros']
        d['is_nomover'] = b_item['is_nomover']
        d['is_slowmover'] = b_item['is_slowmover']
        d['is_fastmover'] = b_item['is_fastmover']

    # D. Graneles Summary
    graneles_summary = []
    for g_code, g in sorted(granel_map.items()):
        linked_items = [b for b in base_skus_dict.values() if b.get('granel_code') == g_code]
        linked_skus = [b['base_sku'] for b in linked_items]
        total_envasado = sum(b.get('stock_envasado', 0.0) for b in linked_items)
        total_pending = sum(b.get('pedidos_pendientes', 0.0) for b in linked_items)
        monthly_dem = {m: 0.0 for m in month_keys}
        for b in linked_items:
            for m in month_keys:
                monthly_dem[m] += b.get('monthly_forecast', {}).get(m, 0.0)
        graneles_summary.append({
            'granel_code': g_code,
            'granel_desc': g['desc'],
            'granel_stock': g['stock'],
            'granel_family': g['familia'],
            'formatos': linked_skus,
            'stock_envasado_total': total_envasado,
            'pedidos_pendientes_total': total_pending,
            'monthly_forecast_total': monthly_dem
        })

    # E. Build Stock Suggestions (CEO & CFO Audited Model)
    _sku_clients = _dd(list)
    _granel_clients = _dd(list)
    _family_clients = _dd(list)

    # Populate client pools from scenario entries
    for e in ws_prev_entries:
        _sku_clients[e['base_sku']].append(e)
        g_c = base_skus_dict.get(e['base_sku'], {}).get('granel_code')
        if g_c: _granel_clients[g_c].append(e)
        fam = base_skus_dict.get(e['base_sku'], {}).get('granel_family')
        if fam: _family_clients[fam].append(e)

    def _consolidate_buyers(entries, b_sku=None, g_code=None):
        by_cli = {}
        for e in entries:
            c = e['cliente']
            if 'CODIMED' in c.upper(): continue
            if c not in by_cli:
                by_cli[c] = {
                    'cliente': c,
                    'pais': e['pais'],
                    'is_nacional': e['is_nacional'],
                    'comercial': e['comercial'],
                    'prev_vol': 0.0,
                    'prev_eur': 0.0,
                    'real_ltm_vol': 0.0,
                    'real_ltm_eur': 0.0,
                    'skus': set()
                }
            by_cli[c]['prev_vol'] += e['vol']
            by_cli[c]['prev_eur'] += e['eur']
            by_cli[c]['skus'].add(e['base_sku'])

        if b_sku:
            for (sku_k, cli_k), rdata in real_sales_by_sku_cli.items():
                if sku_k == b_sku:
                    if 'CODIMED' in cli_k.upper(): continue
                    if cli_k not in by_cli:
                        by_cli[cli_k] = {
                            'cliente': cli_k,
                            'pais': rdata['pais'] or 'ESPAÑA',
                            'is_nacional': rdata['pais'] in ['ESPAÑA', 'ESP', ''],
                            'comercial': rdata['comercial'] or 'Comercial',
                            'prev_vol': 0.0,
                            'prev_eur': 0.0,
                            'real_ltm_vol': 0.0,
                            'real_ltm_eur': 0.0,
                            'skus': set([b_sku])
                        }
                    by_cli[cli_k]['real_ltm_vol'] += rdata['vol']
                    by_cli[cli_k]['real_ltm_eur'] += rdata['eur']

        res = []
        for c, d in by_cli.items():
            es_real = d['real_ltm_vol'] > 0
            if es_real and d['real_ltm_vol'] > 0:
                pvp_calc = d['real_ltm_eur'] / d['real_ltm_vol']
                pvp_tipo = "PVP Real Histórico Cliente (Ene-Ago 2026)"
            elif d['prev_vol'] > 0:
                pvp_calc = d['prev_eur'] / d['prev_vol']
                pvp_tipo = "Precio Tarifa Vigente (Previsión)"
            else:
                pvp_calc = 0.0
                pvp_tipo = "Hipótesis Comercial (Pendiente de Validar)"

            d['vol'] = d['real_ltm_vol'] if es_real else d['prev_vol']
            d['eur'] = d['real_ltm_eur'] if es_real else d['prev_eur']
            d['pvp'] = round(pvp_calc, 2)
            d['pvp_tipo'] = pvp_tipo
            d['es_comprador_real'] = es_real
            d['tipo_relacion'] = "Real Cerrado Ene-Ago 2026 (8 meses)" if es_real else "Previsión Planificada"
            d['skus'] = sorted(list(d['skus']))
            vol_anual = (d['real_ltm_vol'] / 8.0) * 12.0 if es_real else d['prev_vol']
            d['vol_anualizado'] = round(vol_anual, 1)
            d['capacidad_absorcion'] = round(max(vol_anual * 0.5, d['prev_vol'], 10.0), 1)
            res.append(d)

        res.sort(key=lambda x: (1 if x['es_comprador_real'] else 0, x['real_ltm_vol'], x['prev_vol']), reverse=True)
        return res

    suggestions_items = []
    for b_sku, b in base_skus_dict.items():
        stock_libre = b.get('stock_libre_disponible', 0.0)
        nm_q = float(b.get('no_mover_qty', 0) or 0)
        sm_q = float(b.get('slow_mover_qty', 0) or 0)
        if nm_q <= 0 and sm_q <= 0: continue
        if stock_libre <= 0: continue

        desc = b['desc']
        fmt = b['formato_label']
        env_l = float(b.get('envase_litros', 1) or 1)
        cost = float(b.get('precio_unitario', 0) or 0)
        g_code = b.get('granel_code')
        nm_eur = float(b.get('no_mover_eur', 0) or 0)
        sm_eur = float(b.get('slow_mover_eur', 0) or 0)
        tot_ex_qty = round(nm_q + sm_q, 2)
        tot_ex_eur = round(nm_eur + sm_eur, 2)

        exp_6m = sum(float(b.get('monthly_forecast_export', {}).get(m, 0) or 0) for m in presets['N5'])
        nac_6m = float(b.get('demanda_nac_6m', 0) or 0)

        direct_buyers = _consolidate_buyers(_sku_clients.get(b_sku, []), b_sku=b_sku, g_code=g_code)
        granel_buyers = _consolidate_buyers(_granel_clients.get(g_code, []), b_sku=None, g_code=g_code) if g_code else []

        pvp_ref = 0.0
        origen_pvp = "Hipótesis Comercial (Pendiente de Validar)"
        if direct_buyers and direct_buyers[0]['pvp'] > 0:
            pvp_ref = direct_buyers[0]['pvp']
            origen_pvp = direct_buyers[0].get('pvp_tipo', 'PVP Real Histórico Cliente')
        elif granel_buyers and granel_buyers[0]['pvp'] > 0:
            pvp_ref = granel_buyers[0]['pvp']
            origen_pvp = granel_buyers[0].get('pvp_tipo', 'PVP Real Histórico Granel')

        if pvp_ref <= 0 or pvp_ref < cost * 1.10:
            pvp_ref = round(max(cost * 1.60, 1.20), 2)
            origen_pvp = "Hipótesis Comercial Prudente (+37,5% Margen)"

        if nm_q > 0 and sm_q > 0: seg_type = "NO MOVER + SLOW MOVER"
        elif nm_q > 0: seg_type = "NO MOVER PURO (>6m)"
        else: seg_type = "SLOW MOVER (3-6m)"

        exp_buyers = [x for x in direct_buyers if not x['is_nacional']]
        nac_direct = [x for x in direct_buyers if x['is_nacional']]
        nac_granel = [x for x in granel_buyers if x['is_nacional'] and x['cliente'] not in [y['cliente'] for y in nac_direct]]

        gift_info = _get_verified_gift(b_sku, desc, g_code)

        if nm_q > 0 and (exp_6m >= nm_q * 0.5 or (exp_buyers and not nac_direct)):
            plan_id = "PLAN 1"
            plan_name = "P1: Fast-Track Exportación & Re-etiquetado"
            dto_pct = 0.08 if exp_6m < nm_q else 0.0
            dto_label = "0% Dto (Asignación directa pedido en curso)" if dto_pct == 0 else "-8% Dto Pronto Embarque Contenedor"
            pvp_neto = max(cost * 1.25, pvp_ref * (1.0 - dto_pct))
            pago_label = "Confirming / Carta Crédito +30 días extra (90d)"
            promo_label = f"Prioridad embarque + Re-etiquetado en planta sin cargo + Obsequio: {gift_info['desc']}"
            motivo = f"Demanda exportación activa ({exp_6m:,.0f} L/Kg en 6m) o cliente internacional recurrente del SKU."
        elif tot_ex_qty < 150 or tot_ex_eur < 220:
            plan_id = "PLAN 5"
            plan_name = "P5: Re-work en Reactor / Liquidación Outlet Fin de Año"
            dto_pct = 0.35
            dto_label = "-35% Liquidación Flash Lote Cerrado (o Re-work Tanque)"
            pvp_neto = max(cost * 1.15, pvp_ref * (1.0 - dto_pct))
            pago_label = "Contado -5% adicional o 60 días estándar"
            promo_label = f"Vaciado a tanque nodriza ({g_code or 'Base'}) en próxima fabricación o lote outlet cerrado con portes pagados"
            motivo = f"Cola residual de stock ({tot_ex_qty:,.0f} L/Kg); óptimo re-volcar a reactor nodriza ({g_code or 'Base'}) o liquidar en bloque."
        elif nac_granel and (env_l <= 5.0 or env_l == 200.0) and nac_6m < tot_ex_qty * 0.8:
            plan_id = "PLAN 2"
            plan_name = "P2: Switch de Formato (Precio 20L en envase 1L/5L/200L)"
            dto_pct = 0.18
            dto_label = "-18% (Igualar €/L de Garrafa 20L / IBC)"
            pvp_neto = max(cost * 1.35, pvp_ref * (1.0 - dto_pct))
            pago_label = "Vencimiento a 90 días sin recargo financiero"
            promo_label = f"Expositor punto de venta + Merchandising técnico + Obsequio: {gift_info['desc']} ({gift_info['sinergia']})"
            motivo = f"Clientes nacionales compran el mismo granel ({g_code}) en 20L/1000L; alta aceptación al recibir formato alternativo a precio de grande."
        elif nac_direct and (nac_6m >= 1500 or tot_ex_qty >= 2500):
            plan_id = "PLAN 3"
            plan_name = "P3: Bundle Cross-Selling & Bonificación 8+2 (En Especie)"
            dto_pct = 0.20
            dto_label = "Bonif. 8+2 (20% Gratis en especie) + 5% Rappel 2027"
            pvp_neto = max(cost * 1.30, pvp_ref * (1.0 - dto_pct))
            pago_label = "Pago fraccionado 60 / 90 / 120 días en pedido conjunto"
            promo_label = f"Rappel diferido 5% en nota de crédito Q1-2027 + Obsequio: {gift_info['desc']} ({gift_info['sinergia']})"
            motivo = "Clientes históricos activos en España; la bonificación en especie limpia almacén sin erosionar el precio facial de tarifa."
        else:
            plan_id = "PLAN 4"
            plan_name = "P4: Kit Captación con Garantía de Rotación (Canje 90d)"
            dto_pct = 0.25
            dto_label = "-25% Dto Lanzamiento Pack Captación + Ensayo Técnico"
            pvp_neto = max(cost * 1.20, pvp_ref * (1.0 - dto_pct))
            pago_label = "Pago a 120 días con Garantía de Canje Técnico (90d)"
            promo_label = f"Obsequio: {gift_info['desc']} ({gift_info['sinergia']}) + Asistencia técnica en campo"
            motivo = "Ideal como puerta de entrada en nuevos distribuidores regionales o reactivación de cuentas con riesgo comercial mitigado mediante canje a 90 días."

        rec_clients = []
        for x in (nac_direct + exp_buyers)[:3]:
            rel_label = "Real Cerrado Ene-Ago 2026 (8 meses)" if x.get('es_comprador_real') else "Previsión Planificada"
            vol_info = f"{x['vol']:,.0f} u ({rel_label})"
            pvp_info = f"@{x['pvp']:.2f}€" if x['pvp'] > 0 else ""
            rec_clients.append(f"{x['cliente']} [{x['pais']}: {vol_info} {pvp_info}]")
        for x in nac_granel[:2]:
            if len(rec_clients) < 3:
                rel_label = "Real Cerrado Ene-Ago 2026 (8 meses)" if x.get('es_comprador_real') else "Previsión Planificada"
                rec_clients.append(f"{x['cliente']} [Compra granel {g_code}: {x['vol']:,.0f} u ({rel_label}) @{x['pvp']:.2f}€]")
        if not rec_clients:
            fam_b = _consolidate_buyers(_family_clients.get(b.get('granel_family'), []))
            for x in [f for f in fam_b if f['is_nacional']][:2]:
                rec_clients.append(f"{x['cliente']} [Potencial Familia: {x['vol']:,.0f} u]")
        if not rec_clients:
            rec_clients.append("FITOMURCIA S.L. / CAMPOEJIDO S.C.A. / AGROTÉCNICA DEL NORTE (Distribuidores Tractores)")

        clientes_str = " | ".join(rec_clients[:3])
        cap_palet = 1000.0 if env_l >= 500 else (800.0 if env_l >= 150 else (640.0 if env_l >= 15 else (480.0 if env_l >= 4 else 360.0)))
        palets_lib = round(max(0.15, tot_ex_qty / cap_palet), 2)

        TASA_FINANCIERA_ANUAL = 0.055
        TASA_HOLDING_INVENTARIO = 0.15

        dias_extra_map = { "PLAN 1": 30, "PLAN 2": 30, "PLAN 3": 30, "PLAN 4": 60, "PLAN 5": 0 }
        dias_extra = dias_extra_map.get(plan_id, 30)

        # Distribuir excedente físico por cliente concreto
        rem_stock = tot_ex_qty
        ofertas_clientes = []
        ordered_buyers = []
        seen_c = set()
        for x in (nac_direct + exp_buyers):
            if x['cliente'] not in seen_c:
                ordered_buyers.append(x)
                seen_c.add(x['cliente'])
        for x in nac_granel:
            if x['cliente'] not in seen_c:
                ordered_buyers.append(x)
                seen_c.add(x['cliente'])
        if not ordered_buyers:
            fam_b = _consolidate_buyers(_family_clients.get(b.get('granel_family'), []))
            for x in fam_b:
                if x['cliente'] not in seen_c and x['is_nacional']:
                    ordered_buyers.append(x)
                    seen_c.add(x['cliente'])

        for buy in ordered_buyers:
            if rem_stock <= 0:
                break
            cap = buy.get('capacidad_absorcion', 10.0)
            alloc_q = min(rem_stock, cap)
            if alloc_q < 5.0 and rem_stock >= 5.0:
                alloc_q = min(rem_stock, 5.0)
            alloc_q = round(alloc_q, 1)
            if alloc_q <= 0:
                continue

            if not buy['is_nacional']:
                c_plan_id = "PLAN 1"
                c_plan_name = "P1: Fast-Track Exportación & Re-etiquetado"
                c_dto_pct = 0.08 if exp_6m < alloc_q else 0.0
                c_pago_label = "Confirming / Carta de Crédito (90d)"
                c_dias_extra = 30
            elif plan_id == "PLAN 2":
                c_plan_id = "PLAN 2"
                c_plan_name = "P2: Switch de Formato (Precio 20L en 1L/5L/200L)"
                c_dto_pct = 0.18
                c_pago_label = "Vencimiento a 90 días sin recargo"
                c_dias_extra = 30
            elif buy['es_comprador_real'] or nac_6m >= 1500 or alloc_q >= 1000:
                c_plan_id = "PLAN 3"
                c_plan_name = "P3: Bonificación 8+2 (En Especie) + 5% Rappel"
                c_dto_pct = 0.20
                c_pago_label = "Pago fraccionado 60 / 90 / 120 días"
                c_dias_extra = 30
            else:
                c_plan_id = "PLAN 4"
                c_plan_name = "P4: Kit Captación con Garantía de Rotación (90d)"
                c_dto_pct = 0.25
                c_pago_label = "Pago a 120 días (Garantía Canje 90d)"
                c_dias_extra = 60

            c_pvp_ref = buy['pvp'] if buy['pvp'] > cost * 1.10 else pvp_ref
            c_pvp_oferta = round(max(cost * 1.20, c_pvp_ref * (1.0 - c_dto_pct)), 2)
            c_facturacion = round(alloc_q * c_pvp_oferta, 2)
            c_cogs = round(alloc_q * cost, 2)
            c_dto_eur = round(max(0.0, (alloc_q * c_pvp_ref) - c_facturacion), 2)
            c_rappel_eur = round(c_facturacion * 0.05, 2) if c_plan_id == "PLAN 3" else 0.0

            gift_units = max(1, min(5, int(alloc_q / (200.0 if env_l <= 5 else 400.0)) + 1))
            c_gift_cost = round(gift_units * gift_info['coste_unit'] * gift_info['envase_l'], 2)
            c_merch_eur = round(c_gift_cost + (alloc_q * 0.04), 2)
            c_fin_eur = round(c_facturacion * TASA_FINANCIERA_ANUAL * (c_dias_extra / 365.0), 2)
            c_margen_est = round(c_facturacion - c_cogs - c_rappel_eur - c_merch_eur - c_fin_eur, 2)
            c_margen_pct = round((c_margen_est / c_facturacion) * 100, 1) if c_facturacion > 0 else 0.0

            if c_plan_id in ["PLAN 1", "PLAN 2"]:
                c_caja_0_30 = 0.0
                c_caja_31_60 = 0.0
                c_caja_61_90 = c_facturacion
                c_caja_90_plus = 0.0
            elif c_plan_id == "PLAN 3":
                c_caja_0_30 = 0.0
                c_caja_31_60 = round(c_facturacion / 3.0, 2)
                c_caja_61_90 = round(c_facturacion / 3.0, 2)
                c_caja_90_plus = round(c_facturacion - c_caja_31_60 - c_caja_61_90, 2)
            elif c_plan_id == "PLAN 4":
                c_caja_0_30 = 0.0
                c_caja_31_60 = 0.0
                c_caja_61_90 = 0.0
                c_caja_90_plus = c_facturacion
            else:
                c_caja_0_30 = round(c_facturacion * 0.70, 2)
                c_caja_31_60 = round(c_facturacion - c_caja_0_30, 2)
                c_caja_61_90 = 0.0
                c_caja_90_plus = 0.0

            ofertas_clientes.append({
                'cliente': buy['cliente'],
                'pais': buy['pais'],
                'comercial': buy['comercial'],
                'tipo_relacion': buy['tipo_relacion'],
                'es_comprador_real': buy['es_comprador_real'],
                'vol_historico_o_prev': buy['vol'],
                'capacidad_absorcion': cap,
                'vol_asignado': alloc_q,
                'envases_asignados': round(alloc_q / env_l, 1) if env_l > 0 else alloc_q,
                'plan_id': c_plan_id,
                'plan_name': c_plan_name,
                'pvp_tarifa_ref': round(c_pvp_ref, 2),
                'pvp_oferta': round(c_pvp_oferta, 2),
                'dto_pct': round(c_dto_pct * 100, 1),
                'facturacion_neta_eur': c_facturacion,
                'cogs_eur': c_cogs,
                'coste_dto_eur': c_dto_eur,
                'coste_rappel_eur': c_rappel_eur,
                'coste_merch_eur': c_merch_eur,
                'coste_fin_eur': c_fin_eur,
                'margen_estimado_eur': c_margen_est,
                'margen_estimado_pct': c_margen_pct,
                'regalo_sku': gift_info['sku'],
                'regalo_desc': gift_info['desc'],
                'regalo_uds': gift_units,
                'regalo_coste_eur': c_gift_cost,
                'regalo_sinergia': gift_info['sinergia'],
                'plazo_pago': c_pago_label,
                'caja_0_30d_eur': c_caja_0_30,
                'caja_31_60d_eur': c_caja_31_60,
                'caja_61_90d_eur': c_caja_61_90,
                'caja_90d_plus_eur': c_caja_90_plus
            })
            rem_stock = round(rem_stock - alloc_q, 1)

        # Si queda excedente no absorbido, asignar a Plan 5 (Re-work reactor / Outlet)
        if rem_stock > 0:
            c_plan_id = "PLAN 5"
            c_plan_name = "P5: Re-work en Reactor / Liquidación Outlet"
            c_dto_pct = 35.0
            c_pvp_oferta = round(max(cost * 1.15, pvp_ref * (1.0 - 0.35)), 2)
            c_facturacion = round(rem_stock * c_pvp_oferta, 2)
            c_cogs = round(rem_stock * cost, 2)
            c_dto_eur = round(max(0.0, (rem_stock * pvp_ref) - c_facturacion), 2)
            c_rappel_eur = 0.0
            c_merch_eur = round(rem_stock * 0.06, 2)
            c_fin_eur = 0.0
            c_margen_est = round(c_facturacion - c_cogs - c_merch_eur, 2)
            c_margen_pct = round((c_margen_est / c_facturacion) * 100, 1) if c_facturacion > 0 else 0.0
            c_caja_0_30 = round(c_facturacion * 0.70, 2)
            c_caja_31_60 = round(c_facturacion - c_caja_0_30, 2)

            ofertas_clientes.append({
                'cliente': f"Almacén Planta / Reactor Nodriza ({g_code or 'Base'})" if g_code else "Lote Outlet Liquidación Fin de Año",
                'pais': "ESPAÑA",
                'comercial': "Dirección Técnica / Operaciones",
                'tipo_relacion': "Reincorporación Formulación / Outlet",
                'es_comprador_real': False,
                'vol_historico_o_prev': 0.0,
                'capacidad_absorcion': rem_stock,
                'vol_asignado': rem_stock,
                'envases_asignados': round(rem_stock / env_l, 1) if env_l > 0 else rem_stock,
                'plan_id': c_plan_id,
                'plan_name': c_plan_name,
                'pvp_tarifa_ref': round(pvp_ref, 2),
                'pvp_oferta': round(c_pvp_oferta, 2),
                'dto_pct': 35.0,
                'facturacion_neta_eur': c_facturacion,
                'cogs_eur': c_cogs,
                'coste_dto_eur': c_dto_eur,
                'coste_rappel_eur': 0.0,
                'coste_merch_eur': c_merch_eur,
                'coste_fin_eur': 0.0,
                'margen_estimado_eur': c_margen_est,
                'margen_estimado_pct': c_margen_pct,
                'regalo_sku': "-",
                'regalo_desc': "Vaciado a tanque nodriza / Sin obsequio",
                'regalo_uds': 0,
                'regalo_coste_eur': 0.0,
                'regalo_sinergia': "Recuperación de principio activo",
                'plazo_pago': "Contado -5% o Inmediato en Reactor",
                'caja_0_30d_eur': c_caja_0_30,
                'caja_31_60d_eur': c_caja_31_60,
                'caja_61_90d_eur': 0.0,
                'caja_90d_plus_eur': 0.0
            })
            rem_stock = 0.0

        # Totales exactos del SKU derivados de sus ofertas
        ingreso_tot = round(sum(o['facturacion_neta_eur'] for o in ofertas_clientes), 2)
        coste_dto_eur = round(sum(o['coste_dto_eur'] for o in ofertas_clientes), 2)
        coste_rappel_eur = round(sum(o['coste_rappel_eur'] for o in ofertas_clientes), 2)
        coste_merch_eur = round(sum(o['coste_merch_eur'] for o in ofertas_clientes), 2)
        coste_fin_pago_eur = round(sum(o['coste_fin_eur'] for o in ofertas_clientes), 2)
        coste_comercial_total_eur = round(coste_dto_eur + coste_rappel_eur + coste_merch_eur + coste_fin_pago_eur, 2)
        margen_bruto_eur = round(ingreso_tot - tot_ex_eur, 2)
        margen_neto_real_eur = round(sum(o['margen_estimado_eur'] for o in ofertas_clientes), 2)
        margen_neto_real_pct = round((margen_neto_real_eur / ingreso_tot) * 100, 1) if ingreso_tot > 0 else 0.0

        caja_0_30d_eur = round(sum(o['caja_0_30d_eur'] for o in ofertas_clientes), 2)
        caja_31_60d_eur = round(sum(o['caja_31_60d_eur'] for o in ofertas_clientes), 2)
        caja_61_90d_eur = round(sum(o['caja_61_90d_eur'] for o in ofertas_clientes), 2)
        caja_90d_plus_eur = round(sum(o['caja_90d_plus_eur'] for o in ofertas_clientes), 2)

        dio_saving_empresa_days = round(tot_ex_eur / COGS_DIARIO_TOTAL, 3)
        dio_nm_empresa_days = round(nm_eur / COGS_DIARIO_TOTAL, 3)
        dio_tot_pt = round(tot_ex_eur / COGS_DIARIO_PT_NAC, 2)

        dso_impact_empresa_days = round((ingreso_tot * dias_extra) / VENTAS_NETAS_LTM_EUR, 4)
        ccc_net_impact_days = round(-dio_saving_empresa_days + dso_impact_empresa_days, 3)

        # Horizonte de rotación dinámico según velocidad del SKU
        if nm_q > 0 and sm_q == 0:
            holding_days = 270.0  # No-mover sin demanda: previene 9 meses de inmovilización / caducidad
        elif sm_q > 0 and nac_6m > 0:
            holding_days = max(90.0, min(180.0, round((tot_ex_qty / nac_6m) * 180.0, 0)))
        else:
            holding_days = 180.0
        holding_cost_saved_eur = round(tot_ex_eur * TASA_HOLDING_INVENTARIO * (holding_days / 365.0), 2)
        net_financial_gain_eur = round(holding_cost_saved_eur - coste_fin_pago_eur, 2)

        suggestions_items.append({
            'base_sku': b_sku,
            'desc': desc,
            'formato': fmt,
            'envase_l': env_l,
            'granel_code': g_code or '-',
            'segmento': seg_type,
            'is_nomover': nm_q > 0,
            'stock_actual': float(b.get('stock_actual', 0) or 0),
            'no_mover_qty': nm_q,
            'slow_mover_qty': sm_q,
            'exceso_total_qty': tot_ex_qty,
            'exceso_envases': round(tot_ex_qty / env_l, 1) if env_l > 0 else tot_ex_qty,
            'coste_unit': round(cost, 2),
            'no_mover_eur': round(nm_eur, 2),
            'slow_mover_eur': round(sm_eur, 2),
            'exceso_total_eur': tot_ex_eur,
            'pvp_historico': round(pvp_ref, 2),
            'origen_pvp': origen_pvp,
            'pvp_objetivo': round(pvp_neto, 2),
            'dto_pct': round(dto_pct * 100, 1),
            'dto_label': dto_label,
            'coste_dto_eur': coste_dto_eur,
            'rappel_pct': 5.0 if plan_id == "PLAN 3" else 0.0,
            'coste_rappel_eur': coste_rappel_eur,
            'coste_merch_eur': coste_merch_eur,
            'coste_fin_pago_eur': coste_fin_pago_eur,
            'coste_comercial_total_eur': coste_comercial_total_eur,
            'dias_aplazamiento_extra': dias_extra,
            'ingreso_nm_eur': round(sum(o['facturacion_neta_eur'] for o in ofertas_clientes if o['vol_asignado'] <= nm_q), 2) if nm_q > 0 else 0.0,
            'ingreso_neto_eur': ingreso_tot,
            'margen_bruto_eur': margen_bruto_eur,
            'margen_neto_eur': margen_neto_real_eur,
            'margen_neto_pct': margen_neto_real_pct,
            'caja_0_30d_eur': caja_0_30d_eur,
            'caja_31_60d_eur': caja_31_60d_eur,
            'caja_61_90d_eur': caja_61_90d_eur,
            'caja_90d_plus_eur': caja_90d_plus_eur,
            'dio_saving_days': dio_saving_empresa_days,
            'dio_nm_days': dio_nm_empresa_days,
            'dio_pt_saving_days': dio_tot_pt,
            'dso_impact_empresa_days': dso_impact_empresa_days,
            'ccc_net_impact_days': ccc_net_impact_days,
            'holding_days': holding_days,
            'holding_cost_saved_eur': holding_cost_saved_eur,
            'net_financial_gain_eur': net_financial_gain_eur,
            'ratio_dio_vs_dso': round(dio_saving_empresa_days / dso_impact_empresa_days, 1) if dso_impact_empresa_days > 0 else 999.0,
            'palets_liberados': palets_lib,
            'plan_id': plan_id,
            'plan_name': plan_name,
            'clientes_diana': clientes_str,
            'motivo_coherente': motivo,
            'condiciones_pago': pago_label,
            'incentivos_adicionales': promo_label,
            'total_ofertas': len(ofertas_clientes),
            'ofertas_clientes': ofertas_clientes
        })

    suggestions_items.sort(key=lambda x: (0 if x['is_nomover'] else 1, -x['no_mover_eur'], -x['exceso_total_eur']))

    plan_meta_defs = {
        "PLAN 1": {
            "plan_id": "PLAN 1",
            "title": "Fast-Track Exportación & Re-etiquetado",
            "target_segment": "NoMovers en España con Demanda Internacional Activa",
            "price_lever": "0% a -8% Dto Pronto Embarque Contenedor",
            "payment_terms": "Confirming / Carta Crédito 90d (+30d extra sobre 60d)",
            "dias_extra": 30,
            "cost_merch_rate": 0.10,
            "non_price_incentives": "Prioridad de embarque + Re-etiquetado en planta (0,04€/u) + Muestras nueva línea (0,06€/u)",
            "rationale": "Artículos sin previsión en España que sí tienen rotación en exportación. Asignación inmediata a pedidos en curso aliviando turnos de planta."
        },
        "PLAN 2": {
            "plan_id": "PLAN 2",
            "title": "Switch de Formato (Upgrade Envase 1L/5L/200L a Precio de 20L/IBC)",
            "target_segment": "Envases pequeños/alternativos de Graneles de Alta Rotación",
            "price_lever": "-18% s/Tarifa Envase Pequeño (Igualar €/L de Garrafa 20L)",
            "payment_terms": "Vencimiento a 90 días sin recargo (+30d extra)",
            "dias_extra": 30,
            "cost_merch_rate": 0.18,
            "non_price_incentives": "Expositor punto de venta distribuidor + Obsequio técnico verificado de almacén",
            "rationale": "Clientes nacionales compran el mismo producto químico en 20L o 1000L. Al entregarles 1L/5L al precio por litro del 20L perciben un alto ahorro preservando un margen industrial objetivo superior al 50%."
        },
        "PLAN 3": {
            "plan_id": "PLAN 3",
            "title": "Bundles Cross-Selling & Bonificación 8+2 (En Especie + Rappel 2027)",
            "target_segment": "Lotes de Volumen en Referencias Tractoras Nacionales",
            "price_lever": "Bonificación 8+2 (20% en especie sin cargo) + 5% Rappel Q1-2027",
            "payment_terms": "Pago fraccionado 60 / 90 / 120 días (promedio 90d, +30d extra)",
            "dias_extra": 30,
            "cost_merch_rate": 0.22,
            "non_price_incentives": "Nota de crédito 5% diferida a compras 2027 + Obsequio técnico verificado de stock",
            "rationale": "Evita rebajar el precio facial de tarifa en el mercado nacional. La entrega de unidades bonificadas acelera la absorción física del stock y fideliza la campaña 2027."
        },
        "PLAN 4": {
            "plan_id": "PLAN 4",
            "title": "Kit Captación Nuevos Distribuidores con Cláusula de Rotación (Canje 90d)",
            "target_segment": "Especialidades Agronómicas en Formato Mediano (5L / 5Kg / 20L)",
            "price_lever": "-25% Dto Lanzamiento Pack Bienvenida Multirreferencia",
            "payment_terms": "Pago a 120 días (+60d extra) con Cláusula de Rotación Garantizada (Canje 90d)",
            "dias_extra": 60,
            "cost_merch_rate": 0.35,
            "non_price_incentives": "Obsequio técnico complementario verificado + Asistencia agronómica en campo",
            "rationale": "Facilita la apertura de distribuidores o reactivación de cuentas, mitigando el riesgo comercial mediante compromiso de canje a 90 días si no rota."
        },
        "PLAN 5": {
            "plan_id": "PLAN 5",
            "title": "Re-work Técnico en Reactor (Vaciado a Tanque) & Liquidación Outlet",
            "target_segment": "Colas Residuales (<150 L/Kg) y Obsoletos de Almacén",
            "price_lever": "Re-work interno en tanque nodriza o -35% Liquidación Flash Lote Cerrado",
            "payment_terms": "Contado (-5% pronto pago adicional) o 60 días estándar (+0d extra)",
            "dias_extra": 0,
            "cost_merch_rate": 0.08,
            "non_price_incentives": "Reincorporación a tanque nodriza en formulación activa o portes pagados lote cerrado",
            "rationale": "Limpia referencias residuales antes del cierre del 31/12, recuperando el principio activo en reactor según formulación o liquidando en bloque."
        }
    }

    plans_summary = []
    for pid in sorted(plan_meta_defs.keys()):
        p_items = [x for x in suggestions_items if x['plan_id'] == pid]
        meta_p = dict(plan_meta_defs[pid])
        meta_p['skus_count'] = len(p_items)
        meta_p['nomover_skus_count'] = sum(1 for x in p_items if x['is_nomover'])
        meta_p['nomover_qty'] = round(sum(x['no_mover_qty'] for x in p_items), 1)
        meta_p['slowmover_qty'] = round(sum(x['slow_mover_qty'] for x in p_items), 1)
        meta_p['total_qty'] = round(sum(x['exceso_total_qty'] for x in p_items), 1)
        meta_p['nomover_cost_eur'] = round(sum(x['no_mover_eur'] for x in p_items), 2)
        meta_p['slowmover_cost_eur'] = round(sum(x['slow_mover_eur'] for x in p_items), 2)
        meta_p['total_cost_eur'] = round(sum(x['exceso_total_eur'] for x in p_items), 2)
        meta_p['ingreso_nm_eur'] = round(sum(x['ingreso_nm_eur'] for x in p_items), 2)
        meta_p['ingreso_neto_eur'] = round(sum(x['ingreso_neto_eur'] for x in p_items), 2)
        meta_p['coste_dto_eur'] = round(sum(x['coste_dto_eur'] for x in p_items), 2)
        meta_p['coste_rappel_eur'] = round(sum(x.get('coste_rappel_eur', 0) for x in p_items), 2)
        meta_p['coste_merch_eur'] = round(sum(x['coste_merch_eur'] for x in p_items), 2)
        meta_p['coste_fin_pago_eur'] = round(sum(x['coste_fin_pago_eur'] for x in p_items), 2)
        meta_p['coste_comercial_total_eur'] = round(sum(x['coste_comercial_total_eur'] for x in p_items), 2)
        meta_p['margen_bruto_eur'] = round(sum(x['margen_bruto_eur'] for x in p_items), 2)
        meta_p['margen_neto_eur'] = round(sum(x['margen_neto_eur'] for x in p_items), 2)
        meta_p['margen_neto_pct'] = round((meta_p['margen_neto_eur'] / meta_p['ingreso_neto_eur']) * 100, 1) if meta_p['ingreso_neto_eur'] > 0 else 0.0
        meta_p['caja_0_30d_eur'] = round(sum(x.get('caja_0_30d_eur', 0) for x in p_items), 2)
        meta_p['caja_31_60d_eur'] = round(sum(x.get('caja_31_60d_eur', 0) for x in p_items), 2)
        meta_p['caja_61_90d_eur'] = round(sum(x.get('caja_61_90d_eur', 0) for x in p_items), 2)
        meta_p['caja_90d_plus_eur'] = round(sum(x.get('caja_90d_plus_eur', 0) for x in p_items), 2)
        meta_p['dio_nm_days'] = round(sum(x['dio_nm_days'] for x in p_items), 2)
        meta_p['dio_saving_days'] = round(sum(x['dio_saving_days'] for x in p_items), 2)
        meta_p['dso_impact_days'] = round(sum(x['dso_impact_empresa_days'] for x in p_items), 3)
        meta_p['ccc_net_impact_days'] = round(-meta_p['dio_saving_days'] + meta_p['dso_impact_days'], 2)
        meta_p['holding_cost_saved_eur'] = round(sum(x['holding_cost_saved_eur'] for x in p_items), 2)
        meta_p['net_financial_gain_eur'] = round(sum(x['net_financial_gain_eur'] for x in p_items), 2)
        meta_p['ratio_dio_vs_dso'] = round(meta_p['dio_saving_days'] / meta_p['dso_impact_days'], 1) if meta_p['dso_impact_days'] > 0 else 999.0
        meta_p['palets_liberados'] = round(sum(x['palets_liberados'] for x in p_items), 1)

        dio_sav = meta_p['dio_saving_days']
        dso_imp = meta_p['dso_impact_days']
        c_fin = meta_p['coste_fin_pago_eur']
        c_merch = meta_p['coste_merch_eur']
        c_rap = meta_p['coste_rappel_eur']
        meta_p['dio_vs_dso_analysis'] = (
            f"DIO Almacén se reduce en -{dio_sav:.2f}d. El DSO estimado varía en +{dso_imp:.2f}d. "
            f"Coste financiero aplazamiento: {c_fin:,.2f} € (5,5% anual). "
            f"Costes de incentivos: Merchandising/Muestras={c_merch:,.2f} €, Rappel={c_rap:,.2f} €. "
            f"Impacto neto en ciclo de caja (CCC): {-dio_sav + dso_imp:+.2f} días."
        )
        plans_summary.append(meta_p)

    tot_stock_pt_eur = round(sum(b.get('stock_eur', 0) for b in base_skus_dict.values()), 2)
    dio_actual_pt = round(tot_stock_pt_eur / COGS_DIARIO_PT_NAC, 1)
    dio_nm_total = round(sum(x['dio_nm_days'] for x in suggestions_items), 2)
    dio_ex_total = round(sum(x['dio_saving_days'] for x in suggestions_items), 2)
    dso_impact_total = round(sum(x['dso_impact_empresa_days'] for x in suggestions_items), 2)
    ccc_net_impact = round(-dio_ex_total + dso_impact_total, 2)
    coste_dto_total = round(sum(x['coste_dto_eur'] for x in suggestions_items), 2)
    coste_rappel_total = round(sum(x.get('coste_rappel_eur', 0) for x in suggestions_items), 2)
    coste_merch_total = round(sum(x['coste_merch_eur'] for x in suggestions_items), 2)
    coste_fin_total = round(sum(x['coste_fin_pago_eur'] for x in suggestions_items), 2)
    coste_comercial_tot = round(sum(x['coste_comercial_total_eur'] for x in suggestions_items), 2)
    margen_bruto_tot = round(sum(x['margen_bruto_eur'] for x in suggestions_items), 2)
    margen_neto_real_tot = round(sum(x['margen_neto_eur'] for x in suggestions_items), 2)
    holding_saved_tot = round(sum(x['holding_cost_saved_eur'] for x in suggestions_items), 2)
    net_fin_gain_tot = round(sum(x['net_financial_gain_eur'] for x in suggestions_items), 2)
    ratio_dio_dso_tot = round(dio_ex_total / dso_impact_total, 1) if dso_impact_total > 0 else 999.0

    # Resumen de regalos técnicos reservados de stock
    regalos_reservados = _dd(lambda: {'sku': '', 'desc': '', 'uds': 0, 'coste_eur': 0.0, 'sinergia': ''})
    for it in suggestions_items:
        for o in it.get('ofertas_clientes', []):
            g_sku = o.get('regalo_sku')
            if g_sku and g_sku != '-':
                regalos_reservados[g_sku]['sku'] = g_sku
                regalos_reservados[g_sku]['desc'] = o.get('regalo_desc', '')
                regalos_reservados[g_sku]['sinergia'] = o.get('regalo_sinergia', '')
                regalos_reservados[g_sku]['uds'] += o.get('regalo_uds', 0)
                regalos_reservados[g_sku]['coste_eur'] = round(regalos_reservados[g_sku]['coste_eur'] + o.get('regalo_coste_eur', 0.0), 2)
    regalos_reservados_list = sorted(list(regalos_reservados.values()), key=lambda x: -x['coste_eur'])

    stock_suggestions = {
        "kpis": {
            "scenario": scenario_name,
            "total_skus_accionables": len(suggestions_items),
            "nomover_skus": sum(1 for x in suggestions_items if x['is_nomover']),
            "slowmover_skus": sum(1 for x in suggestions_items if x['slow_mover_qty'] > 0),
            "nomover_qty": round(sum(x['no_mover_qty'] for x in suggestions_items), 1),
            "slowmover_qty": round(sum(x['slow_mover_qty'] for x in suggestions_items), 1),
            "exceso_total_qty": round(sum(x['exceso_total_qty'] for x in suggestions_items), 1),
            "nomover_cost_eur": round(sum(x['no_mover_eur'] for x in suggestions_items), 2),
            "slowmover_cost_eur": round(sum(x['slow_mover_eur'] for x in suggestions_items), 2),
            "exceso_total_cost_eur": round(sum(x['exceso_total_eur'] for x in suggestions_items), 2),
            "stock_total_pt_eur": tot_stock_pt_eur,
            "inventario_total_base_eur": INVENTARIO_TOTAL_BALANCE_EUR,
            "cogs_ltm_total_eur": COGS_LTM_TOTAL_EUR,
            "ventas_netas_ltm_eur": VENTAS_NETAS_LTM_EUR,
            "clientes_balance_eur": CLIENTES_BALANCE_EUR,
            "proveedores_balance_eur": PROVEEDORES_BALANCE_EUR,
            "dio_base_total_days": DIO_BASE_OFICIAL,
            "dso_base_total_days": DSO_BASE_OFICIAL,
            "dpo_base_total_days": DPO_BASE_OFICIAL,
            "ccc_base_total_days": CCC_BASE_OFICIAL,
            "dio_actual_total_days": DIO_BASE_OFICIAL,
            "dio_actual_pt_days": dio_actual_pt,
            "dio_saving_nm_days": dio_nm_total,
            "dio_saving_total_days": dio_ex_total,
            "dio_post_nm_days": round(DIO_BASE_OFICIAL - dio_nm_total, 1),
            "dio_post_total_days": round(DIO_BASE_OFICIAL - dio_ex_total, 1),
            "dso_impact_total_days": dso_impact_total,
            "dso_post_total_days": round(DSO_BASE_OFICIAL + dso_impact_total, 1),
            "ccc_net_impact_days": ccc_net_impact,
            "ccc_post_total_days": round(CCC_BASE_OFICIAL + ccc_net_impact, 1),
            "ingreso_nm_eur": round(sum(x['ingreso_nm_eur'] for x in suggestions_items), 2),
            "ingreso_neto_total_eur": round(sum(x['ingreso_neto_eur'] for x in suggestions_items), 2),
            "coste_descuentos_total_eur": coste_dto_total,
            "coste_rappel_diferido_total_eur": coste_rappel_total,
            "coste_merchandising_total_eur": coste_merch_total,
            "coste_financiero_aplazamiento_total_eur": coste_fin_total,
            "coste_comercial_total_eur": coste_comercial_tot,
            "margen_bruto_total_eur": margen_bruto_tot,
            "margen_neto_total_eur": margen_neto_real_tot,
            "margen_neto_pct": round((margen_neto_real_tot / sum(x['ingreso_neto_eur'] for x in suggestions_items)) * 100, 1) if sum(x['ingreso_neto_eur'] for x in suggestions_items) > 0 else 0.0,
            "caja_0_30d_total_eur": round(sum(x['caja_0_30d_eur'] for x in suggestions_items), 2),
            "caja_31_60d_total_eur": round(sum(x['caja_31_60d_eur'] for x in suggestions_items), 2),
            "caja_61_90d_total_eur": round(sum(x['caja_61_90d_eur'] for x in suggestions_items), 2),
            "caja_90d_plus_total_eur": round(sum(x['caja_90d_plus_eur'] for x in suggestions_items), 2),
            "holding_cost_saved_total_eur": holding_saved_tot,
            "net_financial_gain_total_eur": net_fin_gain_tot,
            "ratio_dio_vs_dso": ratio_dio_dso_tot,
            "tasa_financiera_anual_pct": 5.5,
            "tasa_holding_anual_pct": 15.0,
            "palets_liberados_total": round(sum(x['palets_liberados'] for x in suggestions_items), 1),
            "palets_planta_sop": 747.9,
            "regalos_reservados_catalogo": regalos_reservados_list
        },
        "plans": plans_summary,
        "items": suggestions_items
    }
    nov_sum = sum(d['monthly_forecast']['2026-11'] for d in detailed_skus.values())
    dic_sum = sum(d['monthly_forecast']['2026-12'] for d in detailed_skus.values())
    print(f"[{scenario_name.upper()}] Forecast Nov = {nov_sum:,.0f} u, Dic = {dic_sum:,.0f} u | Base SKUs = {len(base_skus_dict)}, Actionable SKUs = {len(suggestions_items)}, Excess Cost = {exceso_total_cost_eur:,.2f} €" if 'exceso_total_cost_eur' in locals() else f"[{scenario_name.upper()}] Forecast Nov = {nov_sum:,.0f} u, Dic = {dic_sum:,.0f} u | Base SKUs = {len(base_skus_dict)}, Actionable SKUs = {len(suggestions_items)}")

    return {
        "graneles": graneles_summary,
        "base_skus": list(base_skus_dict.values()),
        "detailed_skus": list(detailed_skus.values()),
        "stock_suggestions": stock_suggestions
    }

# Build both scenarios
data_revisado = build_scenario_production_data('revisado')
data_oficial = build_scenario_production_data('oficial')

# 6. Construct Final Output Structure
output_data = {
    "metadata": {
        "generated_at": datetime.now().isoformat(),
        "stock_source": os.path.basename(stock_path) if stock_path else "Stock.xlsx",
        "orders_source": ", ".join(os.path.basename(f) for f in active_order_files) if active_order_files else "Pedidos.xlsx",
        "forecast_source": "Revision_Previsiones_Q4_2026_CONSOLIDADO.xlsx (Revisado) / Presupuesto_Ventas_2027_Definitivo.xlsx (Oficial)",
        "current_month": current_month,
        "months": months_metadata,
        "presets": presets,
        "preset_labels": preset_labels,
        "default_preset": "REST_2026",
        "active_scenario": "revisado",
        "total_graneles_stock": sum(x['granel_stock'] for x in data_revisado['graneles']),
        "total_envasado_stock": sum(b['stock_actual'] for b in data_revisado['base_skus'])
    },
    "scenarios": {
        "revisado": data_revisado,
        "oficial": data_oficial
    },
    # Backwards compatibility top-level keys default to 'revisado':
    "graneles": data_revisado["graneles"],
    "base_skus": data_revisado["base_skus"],
    "detailed_skus": data_revisado["detailed_skus"],
    "stock_suggestions": data_revisado["stock_suggestions"]
}

out_file = get_dest(os.path.join('web_dashboard', 'production_planning_data.json'))
with open(out_file, 'w', encoding='utf-8') as f:
    json.dump(output_data, f, ensure_ascii=False, indent=2)

out_js_file = get_dest(os.path.join('web_dashboard', 'production_planning_data.js'))
with open(out_js_file, 'w', encoding='utf-8') as f:
    f.write("window.PRODUCTION_PLANNING_DATA = " + json.dumps(output_data, ensure_ascii=False) + ";\n")

def safe_copy(src, dst):
    try:
        shutil.copy2(src, dst)
    except Exception:
        try:
            with open(src, 'rb') as f_in, open(dst, 'wb') as f_out:
                f_out.write(f_in.read())
        except Exception as e:
            print(f"Aviso copiando {src} a {dst}: {e}")

dir_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'web_dashboard_direccion')
if os.path.exists(dir_dir):
    safe_copy(out_file, get_dest(os.path.join('web_dashboard_direccion', 'production_planning_data.json')))
    safe_copy(out_js_file, get_dest(os.path.join('web_dashboard_direccion', 'production_planning_data.js')))

print(f"\n[OK] Successfully generated {out_file} and synchronized to web_dashboard_direccion.")
print(f"Revisado -> Base SKUs: {len(data_revisado['base_skus'])}, Sugerencias: {len(data_revisado['stock_suggestions']['items'])} SKUs, Exceso: {data_revisado['stock_suggestions']['kpis']['exceso_total_cost_eur']:,.2f} €")
print(f"Oficial  -> Base SKUs: {len(data_oficial['base_skus'])}, Sugerencias: {len(data_oficial['stock_suggestions']['items'])} SKUs, Exceso: {data_oficial['stock_suggestions']['kpis']['exceso_total_cost_eur']:,.2f} €")
