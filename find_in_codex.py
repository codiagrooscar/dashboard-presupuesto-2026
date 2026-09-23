import pandas as pd
import glob, os

targets = ['430000427', '430000326', '430000430', '430000429', 'ALMENDRALIA', 'CARCAIXENT']
codex_dir = r"C:\Users\oscar.ocampo\Documents\Codex\2026-05-20\files-mentioned-by-the-user-facturas"

for f in glob.glob(os.path.join(codex_dir, "*.xlsx")) + glob.glob(os.path.join(codex_dir, "*.csv")):
    try:
        if f.endswith('.xlsx'):
            xl = pd.ExcelFile(f)
            for s in xl.sheet_names:
                df = xl.parse(s, nrows=2000)
                for t in targets:
                    for c in df.columns:
                        m = df[df[c].astype(str).str.contains(t, case=False, na=False)]
                        if len(m) > 0:
                            print(f"[{t}] in {os.path.basename(f)} [{s}] col {c}: {len(m)} rows")
                            sample = m[[k for k in m.columns if any(w in k.lower() for w in ['comercial', 'cliente', 'nom', 'deleg', 'agente'])]].head(2)
                            print(sample)
    except Exception as e:
        pass
