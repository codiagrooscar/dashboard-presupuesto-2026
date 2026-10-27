import glob
import pandas as pd
import os

targets = ['SUSTAINABLE', 'DOÑA', 'DONA', 'ALMENDRALIA', 'CARCAIXENT', '430000427', '430000326', '430000430', '430000429']
excel_files = glob.glob('*.xlsx')

for f in excel_files:
    if f.startswith('~$'): continue
    try:
        xl = pd.ExcelFile(f)
        for s in xl.sheet_names:
            try:
                df = xl.parse(s, nrows=2000)
                for t in targets:
                    for col in df.columns:
                        m = df[df[col].astype(str).str.contains(t, case=False, na=False)]
                        if len(m) > 0:
                            print(f"File: {f} | Sheet: {s} | Target: {t} | Col: {col} | Rows: {len(m)}")
                            sample = m[[c for c in ['Comercial', 'Cliente', 'CLIENTE', 'País', col] if c in m.columns]].head(2)
                            print(sample)
            except Exception:
                pass
    except Exception:
        pass
