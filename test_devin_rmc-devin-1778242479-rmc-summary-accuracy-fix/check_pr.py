import sys, io, pandas as pd
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
pr = pd.read_excel('Template_Files/PR_copy.xlsx', sheet_name=0, header=2)
col_map = {c: str(c).strip().replace('\n', ' ').replace('\r', '') for c in pr.columns}
pr.rename(columns=col_map, inplace=True)
trk = 'Tracking N o.'
for m in [85039, 85834, 85527, 85528, 85570, 85654, 85660, 84713]:
    count = (pd.to_numeric(pr[trk], errors='coerce') == m).sum()
    print(f'MRR {m}: {count} entries in PR')
