from engine.rate_lookup import load_purchase_register, _find_col
import pandas as pd
pr = load_purchase_register('Template_Files/Purchase Register - 2021 - 2026 _Feb 26.xlsx')
tc = _find_col(pr, 'tracking')
mc = _find_col(pr, 'material')
month_col = _find_col(pr, 'month')
r = [c for c in pr.columns if str(c).strip().lower()=='rate'][0]
mask = pr[mc].astype(str).str.upper().str.strip().str.startswith('TPE')
for _, row in pr[mask].iterrows():
    m_val = row[month_col] if month_col else "?"
    print(f"MRR={row[tc]}, Mat={row[mc]}, Rate={row[r]}, Month={m_val}, Size={row.get('Size')}, Mic={row.get('Mic')}")
