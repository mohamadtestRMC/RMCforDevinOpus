import shutil, os, pandas as pd
shutil.copy2('Template_Files/Stores Recordings.xlsx', 'Template_Files/_st4.xlsx')
df = pd.read_excel('Template_Files/_st4.xlsx', sheet_name=0, header=1)
os.remove('Template_Files/_st4.xlsx')

# Find column names
mic_c = [c for c in df.columns if 'mic' in str(c).lower()][0]
width_c = [c for c in df.columns if 'width' in str(c).lower()][0]
qty_c = [c for c in df.columns if 'issue' in str(c).lower() and 'qty' in str(c).lower()][0]
mat_c = [c for c in df.columns if 'sub' in str(c).lower()][0]
proc_c = [c for c in df.columns if 'issue' in str(c).lower() and 'process' in str(c).lower()][0]
wo_c = [c for c in df.columns if 'issue' in str(c).lower() and 'wo' in str(c).lower()][0]
mrr_c = [c for c in df.columns if 'm.r.r' in str(c).lower() and 'no' in str(c).lower()][0]
sup_c = [c for c in df.columns if 'supplier' in str(c).lower()][0]

for mrr in [85157, 85226]:
    mask = pd.to_numeric(df[mrr_c], errors='coerce') == mrr
    rows = df[mask]
    wo_mask = rows[wo_c].astype(str).str.upper() == 'L00335'
    wo_rows = rows[wo_mask]
    print(f"MRR {mrr} for WO L00335: {len(wo_rows)} rows")
    if len(wo_rows) > 0:
        for _, r in wo_rows.iterrows():
            print(f"  Size={r[width_c]}, Mic={r[mic_c]}, Qty={r[qty_c]}, "
                  f"Mat={r[mat_c]}, Process={r[proc_c]}, Supplier={r[sup_c]}")
    else:
        print("  No rows. All WOs for this MRR:")
        for _, r in rows.head(5).iterrows():
            print(f"  WO={r[wo_c]}, Size={r[width_c]}, Mic={r[mic_c]}, Qty={r[qty_c]}, Mat={r[mat_c]}")

# Also check Row 45: MRR 85330 for G00418
print()
mrr = 85330
mask = pd.to_numeric(df[mrr_c], errors='coerce') == mrr
rows = df[mask]
wo_mask = rows[wo_c].astype(str).str.upper() == 'G00418'
wo_rows = rows[wo_mask]
print(f"MRR {mrr} for WO G00418: {len(wo_rows)} rows")
for _, r in wo_rows.iterrows():
    print(f"  Size={r[width_c]}, Mic={r[mic_c]}, Qty={r[qty_c]}, "
          f"Mat={r[mat_c]}, Process={r[proc_c]}, Supplier={r[sup_c]}")
