import pandas as pd
import glob, os

targets = ['ALMENDRALIA', 'CARCAIXENT', 'SUSTAINABLE']
files = [
    'Consolidado_Clasificacion_Comercial_2026.xlsx',
    'Consolidado_Ventas_y_Previsiones_2026.xlsx',
    'Ventas reales de Eebro a Agosto de 2026.xlsx',
    'Clasificacion_Clientes_y_Productos_2026.xlsx',
    'Clasif clientes y productos 2026.xlsx'
]

for f in files:
    if not os.path.exists(f): continue
    try:
        xl = pd.ExcelFile(f)
        for s in xl.sheet_names:
            df = xl.parse(s)
            for t in targets:
                for c in df.columns:
                    m = df[df[c].astype(str).str.contains(t, case=False, na=False)]
                    if len(m) > 0:
                        print(f"[{t}] found in {f} [{s}] col {c}: {len(m)}")
                        for _, row in m.head(2).iterrows():
                            print({k: row[k] for k in row.index if k in ['Comercial', 'Cliente', 'CLIENTE', 'País', c]})
    except Exception as e:
        pass
