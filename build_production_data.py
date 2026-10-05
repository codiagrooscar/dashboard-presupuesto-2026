import openpyxl
import json
import re
from datetime import datetime

print("=== Generating Production Planning Data ===")

import os
import shutil
import subprocess

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

def get_latest_october_file():
    candidates = [
        f for f in os.listdir('.')
        if f.startswith('Pedidos') and f.endswith('.xlsx') and not f.startswith(('temp_', '~$'))
    ]
    oct_candidates = []
    for f in candidates:
        m = re.search(r'(\d{1,2})\.(\d{1,2})', f)
        if m:
            day, month = int(m.group(1)), int(m.group(2))
            if month == 10 and day >= 2:
                oct_candidates.append((day, month, f))
    if oct_candidates:
        oct_candidates.sort(key=lambda x: (x[1], x[0]), reverse=True)
        return oct_candidates[0][2]
    return 'Pedidos 02.10 comerciales.xlsx' if os.path.exists('Pedidos 02.10 comerciales.xlsx') else None

# 1. Load Stock from latest stock file
stock_path = get_latest_stock_file() or r'C:\Users\oscar.ocampo\OneDrive - Agroquímica Codiagro S.L\Escritorio\Estado de pedidos\Aplicacion web\Stock 02.10.xlsx'
print(f"Loading stock from: {stock_path}")
wb_stock = safe_load_workbook(stock_path, data_only=True)
ws_stock = wb_stock.active

stock_map = {}
for row in ws_stock.iter_rows(min_row=4, values_only=True):
    code = row[3]
    desc = row[4]
    qty = row[5]
    price = row[6]
    if code:
        c_clean = str(code).strip().upper()
        stock_map[c_clean] = {
            'codigo': c_clean,
            'desc': str(desc or '').strip(),
            'stock': float(qty or 0),
            'precio': float(price or 0)
        }

print(f"Loaded {len(stock_map)} stock items from warehouse. Total units in stock: {sum(x['stock'] for x in stock_map.values()):,.0f}")

# 2. Load Pending Orders from dashboard_data.json (period 2026-09) and latest October orders
with open('web_dashboard/dashboard_data.json', 'r', encoding='utf-8') as f:
    dash_data = json.load(f)

pending_orders = {}

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

# Additional pending orders from latest October orders file
oct_file = get_latest_october_file()
if oct_file:
    try:
        print(f"Loading October pending orders from: {oct_file}")
        wb_ped = safe_load_workbook(oct_file, data_only=True)
        ws_ped = wb_ped.active
        for row in ws_ped.iter_rows(min_row=2, values_only=True):
            p_uds = row[10]
            if p_uds and float(p_uds) > 0:
                c_art = str(row[6] or '').strip().upper()
                d_art = str(row[7] or '').strip()
                cli = str(row[5] or '').strip()
                if c_art not in pending_orders:
                    pending_orders[c_art] = {
                        'sku': c_art,
                        'desc': d_art,
                        'pendientes_uds': 0.0,
                        'is_nacional': True,
                        'clientes': set()
                    }
                pending_orders[c_art]['pendientes_uds'] += float(p_uds)
                pending_orders[c_art]['clientes'].add(cli)
    except Exception as e:
        print(f"Notice checking {oct_file}:", e)

# Convert client sets to sorted lists
for s in pending_orders:
    pending_orders[s]['clientes'] = sorted(list(pending_orders[s]['clientes']))

print(f"Loaded {len(pending_orders)} pending order SKUs. Total pending units: {sum(x['pendientes_uds'] for x in pending_orders.values()):,.0f}")

# 3. Load 15-month Forecast from Presupuesto_Ventas_2027_Definitivo.xlsx
wb_prev = safe_load_workbook('Presupuesto_Ventas_2027_Definitivo.xlsx', read_only=False, data_only=True)
ws_prev = wb_prev['Previsión Matriz Horizontal']

month_keys = [
    '2026-10', '2026-11', '2026-12',
    '2027-01', '2027-02', '2027-03', '2027-04', '2027-05', '2027-06',
    '2027-07', '2027-08', '2027-09', '2027-10', '2027-11', '2027-12'
]
month_col_idx = [27, 29, 31, 35, 37, 39, 41, 43, 45, 47, 49, 51, 53, 55, 57]

detailed_skus = {} # sku -> details

country_code_map = {
    'TURQUIA': 'TR', 'COLOMBIA': 'CO', 'PERU': 'PE', 'EGIPTO': 'EG',
    'MARRUECOS': 'MA', 'GRECIA': 'GR', 'ECUADOR': 'EC', 'PORTUGAL': 'PT',
    'CHILE': 'CL', 'ARGELIA': 'DZ', 'MEXICO': 'MX', 'ITALIA': 'IT',
    'JORDANIA': 'JO', 'LIBANO': 'LB', 'TUNEZ': 'TN', 'GUATEMALA': 'GT'
}

for row in ws_prev.iter_rows(min_row=2, values_only=True):
    comercial = str(row[0] or '').strip()
    pais = str(row[1] or '').strip().upper()
    cliente = str(row[2] or '').strip()
    codigo = row[4]
    desc = str(row[5] or '').strip()
    envase = str(row[6] or '').strip()
    
    if not codigo:
        continue
    
    raw_sku = str(codigo).strip().upper()
    is_nacional = (pais in ['ESPAÑA', 'ESP', ''])
    
    # Export suffix logic
    has_suffix = bool(re.search(r'[A-Z]{2}$', raw_sku)) and len(raw_sku) >= 8
    if not is_nacional and not has_suffix:
        ctry_code = country_code_map.get(pais, pais[:2] if len(pais) >= 2 else 'EX')
        sku = f"{raw_sku}{ctry_code}"
    else:
        sku = raw_sku

    # Base Physical SKU (packaging level)
    # E.g. AN0005TR -> AN0005, BIO0020CO -> BIO0020, A-50(2)0020 -> A-50(2)0020
    m_base = re.match(r'^([A-Z0-9\-\(\)]+?\d{4})([A-Z]{2})$', sku)
    if m_base:
        base_sku = m_base.group(1)
    else:
        base_sku = sku

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
        
    for m_key, c_idx in zip(month_keys, month_col_idx):
        if c_idx < len(row):
            val = row[c_idx]
            try:
                if val:
                    detailed_skus[sku]['monthly_forecast'][m_key] += float(val)
            except:
                pass

print(f"Total Unique Forecast SKUs: {len(detailed_skus)}")

# 4. Integrate Pending Orders that might not be in forecast
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

# Attach pending units to detailed SKUs
for sku, data in detailed_skus.items():
    p_info = pending_orders.get(sku, {})
    data['pendientes_uds'] = p_info.get('pendientes_uds', 0.0)
    data['clientes'] = sorted(list(data['clientes']))
    data['comerciales'] = sorted(list(data['comerciales']))

# 5. Resolve Stock Mapping
# Priority:
# 1. Exact match in stock
# 2. Base SKU match (e.g. AN0020CO -> AN0020)
# 3. Prefix product family match (e.g. AGMG0005 -> AGMG)
def resolve_stock(base_sku, sku):
    if sku in stock_map:
        return stock_map[sku]['stock'], stock_map[sku]['desc'], 'exact'
    if base_sku in stock_map:
        return stock_map[base_sku]['stock'], stock_map[base_sku]['desc'], 'base'
    # try prefix before 4-digit package
    m = re.match(r'^([A-Z0-9\-\(\)]+?)(\d{4})$', base_sku)
    if m:
        pfx = m.group(1)
        if pfx in stock_map:
            return stock_map[pfx]['stock'], stock_map[pfx]['desc'], 'family'
    return 0.0, '', 'none'

# 6. Build Aggregated Physical Base SKUs (Plant Formulation / Packaging level)
base_skus_dict = {}

for sku, d in detailed_skus.items():
    b_sku = d['base_sku']
    if b_sku not in base_skus_dict:
        stock_qty, stock_desc, match_type = resolve_stock(b_sku, sku)
        base_skus_dict[b_sku] = {
            'base_sku': b_sku,
            'desc': d['desc'] or stock_desc,
            'envase': d['envase'],
            'stock_actual': stock_qty,
            'stock_match_type': match_type,
            'pedidos_pendientes': 0.0,
            'monthly_forecast': {m: 0.0 for m in month_keys},
            'destinos': [],
            'has_nacional': False,
            'has_export': False
        }
    
    b_item = base_skus_dict[b_sku]
    b_item['pedidos_pendientes'] += d['pendientes_uds']
    if d['is_nacional']:
        b_item['has_nacional'] = True
    else:
        b_item['has_export'] = True
        
    for m in month_keys:
        b_item['monthly_forecast'][m] += d['monthly_forecast'][m]
        
    b_item['destinos'].append({
        'sku': d['sku'],
        'desc': d['desc'],
        'pais': d['pais'],
        'ambito': d['ambito'],
        'pendientes_uds': d['pendientes_uds'],
        'monthly_forecast': d['monthly_forecast']
    })

# Also assign stock to detailed items
for sku, d in detailed_skus.items():
    b_sku = d['base_sku']
    b_item = base_skus_dict[b_sku]
    d['stock_actual'] = b_item['stock_actual']
    d['stock_match_type'] = b_item['stock_match_type']

print(f"Consolidated into {len(base_skus_dict)} Physical Base SKUs and {len(detailed_skus)} Detailed SKUs.")

# 7. Construct Final Output Structure
months_metadata = [
    {"key": "2026-10", "label_es": "Oct 26 (N)", "label_en": "Oct 26 (N)", "year": 2026, "month": 10},
    {"key": "2026-11", "label_es": "Nov 26", "label_en": "Nov 26", "year": 2026, "month": 11},
    {"key": "2026-12", "label_es": "Dic 26", "label_en": "Dec 26", "year": 2026, "month": 12},
    {"key": "2027-01", "label_es": "Ene 27", "label_en": "Jan 27", "year": 2027, "month": 1},
    {"key": "2027-02", "label_es": "Feb 27", "label_en": "Feb 27", "year": 2027, "month": 2},
    {"key": "2027-03", "label_es": "Mar 27", "label_en": "Mar 27", "year": 2027, "month": 3},
    {"key": "2027-04", "label_es": "Abr 27", "label_en": "Apr 27", "year": 2027, "month": 4},
    {"key": "2027-05", "label_es": "May 27", "label_en": "May 27", "year": 2027, "month": 5},
    {"key": "2027-06", "label_es": "Jun 27", "label_en": "Jun 27", "year": 2027, "month": 6},
    {"key": "2027-07", "label_es": "Jul 27", "label_en": "Jul 27", "year": 2027, "month": 7},
    {"key": "2027-08", "label_es": "Ago 27", "label_en": "Aug 27", "year": 2027, "month": 8},
    {"key": "2027-09", "label_es": "Sep 27", "label_en": "Sep 27", "year": 2027, "month": 9},
    {"key": "2027-10", "label_es": "Oct 27", "label_en": "Oct 27", "year": 2027, "month": 10},
    {"key": "2027-11", "label_es": "Nov 27", "label_en": "Nov 27", "year": 2027, "month": 11},
    {"key": "2027-12", "label_es": "Dic 27", "label_en": "Dec 27", "year": 2027, "month": 12}
]

presets = {
    "N": ["2026-10"],
    "N1": ["2026-10", "2026-11"],
    "N2": ["2026-10", "2026-11", "2026-12"],
    "N5": ["2026-10", "2026-11", "2026-12", "2027-01", "2027-02", "2027-03"],
    "REST_2026": ["2026-10", "2026-11", "2026-12"],
    "Q1_2027": ["2027-01", "2027-02", "2027-03"],
    "H1_2027": ["2027-01", "2027-02", "2027-03", "2027-04", "2027-05", "2027-06"],
    "ALL_2027": ["2027-01", "2027-02", "2027-03", "2027-04", "2027-05", "2027-06", "2027-07", "2027-08", "2027-09", "2027-10", "2027-11", "2027-12"],
    "ALL_15": month_keys
}

output_data = {
    "metadata": {
        "generated_at": datetime.now().isoformat(),
        "stock_source": os.path.basename(stock_path) if stock_path else "Stock.xlsx",
        "orders_source": oct_file or "Pedidos.xlsx",
        "forecast_source": "Presupuesto_Ventas_2027_Definitivo.xlsx",
        "current_month": "2026-10",
        "months": months_metadata,
        "presets": presets,
        "default_preset": "N2"
    },
    "base_skus": list(base_skus_dict.values()),
    "detailed_skus": list(detailed_skus.values())
}

out_file = 'web_dashboard/production_planning_data.json'
with open(out_file, 'w', encoding='utf-8') as f:
    json.dump(output_data, f, ensure_ascii=False, indent=2)

out_js_file = 'web_dashboard/production_planning_data.js'
with open(out_js_file, 'w', encoding='utf-8') as f:
    f.write("window.PRODUCTION_PLANNING_DATA = " + json.dumps(output_data, ensure_ascii=False) + ";\n")

# Keep root in sync
shutil.copy2(out_file, 'production_planning_data.json')
shutil.copy2(out_js_file, 'production_planning_data.js')

print(f"Successfully generated {out_file} and {out_js_file} ({len(output_data['base_skus'])} base SKUs, {len(output_data['detailed_skus'])} detailed SKUs)")
