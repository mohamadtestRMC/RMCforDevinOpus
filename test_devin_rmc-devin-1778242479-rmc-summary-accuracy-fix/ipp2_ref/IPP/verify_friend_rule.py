"""
Re-verify with friend's confirmed rule:
- ALL MRRs (no date filter, no threshold filter)
- Weighted average by quantity
"""
import io, os, shutil, openpyxl
from engine.fill_jobtrack import COLS
from engine.mrr_lookup import load_stores_recordings, lookup_mrr_with_qty
from engine.rate_lookup import load_purchase_register, lookup_film_rate_weighted, _find_col

BASE = "Template_Files"
stores_df = load_stores_recordings(f"{BASE}/Stores Recordings.xlsx")
pr_df = load_purchase_register(f"{BASE}/Purchase Register - 2021 - 2026 _Feb 26.xlsx")

tmp = f"{BASE}/_tmp_gt2.xlsx"
shutil.copy2(f"{BASE}/Jobtrack Feb With MRR.xlsx", tmp)
wb_gt = openpyxl.load_workbook(tmp, data_only=True)
ws_gt = wb_gt.active

rows = [
    (42, 'Fresh1', 78, 79, 71, 'LAMINATION'),
    (43, 'Fresh1', 78, 79, 71, 'LAMINATION'),
    (45, 'Film', 54, 55, 47, 'PRINTING'),
    (46, 'Film', 54, 55, 47, 'PRINTING'),
]

print("FRIEND'S RULE: ALL MRRs, weighted avg, no date filter, no threshold")
print("="*70)

for row, mat_type, mr_col, rate_col, name_col, process in rows:
    order = str(ws_gt.cell(row=row, column=COLS['Order_No']).value or '').strip()
    mat_name = str(ws_gt.cell(row=row, column=name_col).value or '').strip()
    mat_size = ws_gt.cell(row=row, column=name_col + 1).value
    mat_mic = ws_gt.cell(row=row, column=name_col + 2).value
    gt_mr = ws_gt.cell(row=row, column=mr_col).value
    gt_rate = float(ws_gt.cell(row=row, column=rate_col).value or 0)

    # Find ALL MRRs (no filtering)
    mrr_qty = lookup_mrr_with_qty(stores_df, mat_name, mat_mic, mat_size, order, process)
    if not mrr_qty:
        mrr_qty = lookup_mrr_with_qty(stores_df, mat_name, mat_mic, None, order, process)
    if not mrr_qty:
        mrr_qty = lookup_mrr_with_qty(stores_df, mat_name, mat_mic, None, order)

    # Get rate per MRR from PR
    rates = {}
    for m, q in mrr_qty.items():
        r = lookup_film_rate_weighted(pr_df, {m: q}, mat_name, mat_size, mat_mic)
        if r == 0:
            r = lookup_film_rate_weighted(pr_df, {m: q}, mat_name, None, mat_mic)
        rates[m] = r

    # Weighted average using ALL
    total_qty = sum(mrr_qty.values())
    weighted = sum(rates[m] * mrr_qty[m] for m in mrr_qty) / total_qty if total_qty else 0

    match = abs(weighted - gt_rate) < 0.01
    diff = abs(weighted - gt_rate)

    print(f"\nRow {row} | {mat_type}={mat_name} | Order={order}")
    for m in sorted(mrr_qty):
        print(f"  MRR {m}: Qty={mrr_qty[m]}, Rate={rates[m]}")
    print(f"  Weighted Avg (ALL): {weighted:.6f}")
    print(f"  Ground Truth:       {gt_rate:.6f}")
    print(f"  Diff:               {diff:.6f}")
    print(f"  Match: {'✅ YES' if match else '❌ NO'}")
    print(f"  GT MRRs listed:     {gt_mr}")

wb_gt.close()
os.remove(tmp)
