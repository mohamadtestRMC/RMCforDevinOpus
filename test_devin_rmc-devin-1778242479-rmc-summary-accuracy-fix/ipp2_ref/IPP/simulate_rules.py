"""
Simulate proposed rules on rows 42, 43, 45, 46 — NO code changes.
Rules:
1. Prefer current-month MRRs only (if exist)
2. Prefer single-supplier subset when possible
3. No 10% threshold filtering
4. Weighted avg by quantity
5. For Row 43, also try simple average
"""
import io, os, shutil, openpyxl, pandas as pd
from engine.fill_jobtrack import COLS
from engine.mrr_lookup import load_stores_recordings, lookup_mrr_with_qty
from engine.rate_lookup import (
    load_purchase_register, lookup_film_rate_weighted,
    filter_mrr_by_pr, _find_col
)
from engine.supplier_rates import build_mrr_supplier_map

BASE = "Template_Files"
stores_df = load_stores_recordings(f"{BASE}/Stores Recordings.xlsx")
pr_df = load_purchase_register(f"{BASE}/Purchase Register - 2021 - 2026 _Feb 26.xlsx")
mrr_sup = build_mrr_supplier_map(stores_df)

# Get PR month per MRR
tracking_col = _find_col(pr_df, 'tracking')
month_col = _find_col(pr_df, 'month')
rate_col_name = [c for c in pr_df.columns if str(c).strip().lower() == 'rate'][0]
material_col = _find_col(pr_df, 'material')

def get_mrr_month(mrr_num):
    mask = pd.to_numeric(pr_df[tracking_col], errors='coerce') == mrr_num
    rows = pr_df[mask]
    if rows.empty:
        return None
    return str(rows.iloc[0][month_col]).strip() if month_col else None

def get_mrr_rate(mrr_num, mat_name):
    return lookup_film_rate_weighted(pr_df, {mrr_num: 1}, mat_name, None, None)

# Load ground truth
tmp = f"{BASE}/_tmp_gt.xlsx"
shutil.copy2(f"{BASE}/Jobtrack Feb With MRR.xlsx", tmp)
wb_gt = openpyxl.load_workbook(tmp, data_only=True)
ws_gt = wb_gt.active

rows_to_check = [
    (42, 'Fresh1', 78, 79, 71, 'LAMINATION'),
    (43, 'Fresh1', 78, 79, 71, 'LAMINATION'),
    (45, 'Film', 54, 55, 47, 'PRINTING'),
    (46, 'Film', 54, 55, 47, 'PRINTING'),
]

REPORT_MONTH = "2-2026"

for row, mat_type, mr_col, rate_col, name_col, process in rows_to_check:
    uid = ws_gt.cell(row=row, column=1).value
    order = str(ws_gt.cell(row=row, column=COLS['Order_No']).value or '').strip()
    mat_name = str(ws_gt.cell(row=row, column=name_col).value or '').strip()
    mat_size = ws_gt.cell(row=row, column=name_col + 1).value
    mat_mic = ws_gt.cell(row=row, column=name_col + 2).value
    gt_mr = ws_gt.cell(row=row, column=mr_col).value
    gt_rate = ws_gt.cell(row=row, column=rate_col).value

    print(f"\n{'='*70}")
    print(f"ROW {row} | UID={uid} | Order={order} | {mat_type}={mat_name}")
    print(f"Ground Truth: MR={gt_mr}, Rate={gt_rate}")
    print(f"{'='*70}")

    # Discover all MRRs
    mrr_qty = lookup_mrr_with_qty(stores_df, mat_name, mat_mic, mat_size, order, process)
    if not mrr_qty:
        mrr_qty = lookup_mrr_with_qty(stores_df, mat_name, mat_mic, None, order, process)
    if not mrr_qty:
        mrr_qty = lookup_mrr_with_qty(stores_df, mat_name, mat_mic, None, order)

    # Enrich with month & supplier
    print(f"\nAll discovered MRRs:")
    mrr_info = {}
    for m, q in mrr_qty.items():
        month = get_mrr_month(m)
        sup = mrr_sup.get(int(m), 'UNKNOWN')
        rate = get_mrr_rate(m, mat_name)
        mrr_info[m] = {'qty': q, 'month': month, 'supplier': sup, 'rate': rate}
        print(f"  MRR {m}: Qty={q}, Rate={rate}, Month={month}, Supplier={sup}")

    # RULE 1: Prefer current-month MRRs
    current = {m: info for m, info in mrr_info.items() if info['month'] == REPORT_MONTH}
    old = {m: info for m, info in mrr_info.items() if info['month'] != REPORT_MONTH}

    if current:
        selected = current
        rule_applied = "Rule 1: Current-month MRRs only"
    else:
        selected = mrr_info
        rule_applied = "No current-month MRRs, using all"

    print(f"\nAfter Rule 1 (current month): {list(selected.keys())}")

    # RULE 2: Single-supplier subset (only if multiple suppliers)
    suppliers = set(info['supplier'] for info in selected.values())
    if len(suppliers) > 1:
        # Try each supplier, pick the one with most qty
        best_sup = None
        best_qty = 0
        for sup in suppliers:
            sup_qty = sum(info['qty'] for info in selected.values() if info['supplier'] == sup)
            if sup_qty > best_qty:
                best_qty = sup_qty
                best_sup = sup
        single_sup = {m: info for m, info in selected.items() if info['supplier'] == best_sup}
        print(f"  Multiple suppliers: {suppliers}")
        print(f"  Dominant supplier: {best_sup} (qty={best_qty})")
        # Don't force single supplier — just note it
    else:
        single_sup = selected

    # RULE 3: No 10% filtering — use all selected
    # RULE 4: Weighted average by quantity
    total_qty = sum(info['qty'] for info in selected.values())
    if total_qty > 0:
        weighted_rate = sum(info['rate'] * info['qty'] for info in selected.values()) / total_qty
    else:
        weighted_rate = 0

    chosen_mrrs = sorted(selected.keys())
    mr_str = '/'.join(str(m) for m in chosen_mrrs)

    match = abs(weighted_rate - float(gt_rate)) < 0.01 if gt_rate else False

    print(f"\n--- RESULT (Rules 1-4) ---")
    print(f"  Chosen MRRs:  {mr_str}")
    print(f"  Weighted Rate: {weighted_rate:.6f}")
    print(f"  Ground Truth:  {gt_rate}")
    print(f"  Match: {'YES' if match else 'NO'}")
    print(f"  Rule Applied:  {rule_applied}")

    # If no match, try with single-supplier subset
    if not match and len(suppliers) > 1:
        total_qty_s = sum(info['qty'] for info in single_sup.values())
        if total_qty_s > 0:
            weighted_rate_s = sum(info['rate'] * info['qty'] for info in single_sup.values()) / total_qty_s
        else:
            weighted_rate_s = 0
        match_s = abs(weighted_rate_s - float(gt_rate)) < 0.01
        chosen_s = sorted(single_sup.keys())
        print(f"\n--- ALT: Single-supplier ({best_sup}) ---")
        print(f"  Chosen MRRs:  {'/'.join(str(m) for m in chosen_s)}")
        print(f"  Weighted Rate: {weighted_rate_s:.6f}")
        print(f"  Match: {'YES' if match_s else 'NO'}")

    # RULE 5: For Row 43, try simple average
    if row == 43:
        rates = [info['rate'] for info in selected.values() if info['rate'] > 0]
        if rates:
            simple_avg = sum(rates) / len(rates)
            match_simple = abs(simple_avg - float(gt_rate)) < 0.01
            print(f"\n--- ALT: Simple Average (Row 43 only) ---")
            print(f"  Rates: {rates}")
            print(f"  Simple Avg: {simple_avg:.6f}")
            print(f"  Match: {'YES' if match_simple else 'NO'}")

        # Also try all MRRs weighted
        total_all = sum(info['qty'] for info in mrr_info.values())
        if total_all > 0:
            weighted_all = sum(info['rate'] * info['qty'] for info in mrr_info.values()) / total_all
            match_all = abs(weighted_all - float(gt_rate)) < 0.01
            print(f"\n--- ALT: All MRRs weighted avg ---")
            print(f"  Rate: {weighted_all:.6f}")
            print(f"  Match: {'YES' if match_all else 'NO'}")

        # Try simple avg of ALL MRRs
        all_rates = [info['rate'] for info in mrr_info.values() if info['rate'] > 0]
        if all_rates:
            simple_all = sum(all_rates) / len(all_rates)
            match_sa = abs(simple_all - float(gt_rate)) < 0.01
            print(f"\n--- ALT: Simple avg of ALL MRRs ---")
            print(f"  Rates: {all_rates}")
            print(f"  Simple Avg: {simple_all:.6f}")
            print(f"  Match: {'YES' if match_sa else 'NO'}")

wb_gt.close()
os.remove(tmp)
