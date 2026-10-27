import openpyxl
import pandas as pd
import unicodedata
import os

def norm(s):
    if not isinstance(s, str): return ''
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn').upper().strip()

targets = ["SUSTAINABLE", "ANA", "ALMENDRALIA", "CARCAIXENT"]

files = [
    "Consolidado_Clasificacion_Comercial_2026.xlsx",
    "Clasif clientes y productos 2026.xlsx",
    "Ventas reales de Eebro a Agosto de 2026.xlsx",
    "Ventas esperadas de sep a dic de 2026.xlsx",
    "Presupuesto_Ventas_2027.xlsx"
]

for f in files:
    if not os.path.exists(f): continue
    try:
        xl = pd.ExcelFile(f)
        for s in xl.sheet_names:
            df = xl.parse(s)
            for t in targets:
                # search in all string columns
                for col in df.columns:
                    match = df[df[col].astype(str).str.contains(t, case=False, na=False)]
                    if len(match) > 0:
                        print(f"MATCH '{t}' in {f} -> [{s}] -> col '{col}':")
                        # print interesting cols
                        cols_to_print = [c for c in ['Comercial', 'Cliente', 'CLIENTE', 'País', 'País ', col] if c in match.columns]
                        print(match[cols_to_print].drop_duplicates().head(5))
    except Exception as e:
        print(f"Error {f}: {e}")
