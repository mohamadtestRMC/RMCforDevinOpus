"""
FULL SIMULATION: Friend's rule across ALL Jobtrack rows.
Rule: ALL MRRs, weighted avg, no date filter on MRR date, no 10% threshold.
Also test: Does filtering Stores by Issue Date (current/prev month) change results?

Compare every rate cell vs ground truth.
"""
import io, os, shutil, openpyxl, pandas as pd
from engine.fill_jobtrack import COLS, DATA_START_ROW, _safe_str, _safe_float, _compute_total
from engine.mrr_lookup import load_stores_recordings, lookup_mrr_with_qty, lookup_mrr
from engine.rate_lookup import (
    load_purchase_register, lookup_film_rate_weighted,
    filter_mrr_by_pr, lookup_material_rate_for_month, _find_col
)
from engine.supplier_rates import build_mrr_supplier_map, get_supplier_for_mrrs, load_granules_rates, load_megapack_rates, lookup_megapack_rate

BASE = "Template_Files"
TOLERANCE = 0.01

# Load data
stores_df = load_stores_recordings(f"{BASE}/Stores Recordings.xlsx")
pr_df = load_purchase_register(f"{BASE}/Purchase Register - 2021 - 2026 _Feb 26.xlsx")
mrr_sup_map = build_mrr_supplier_map(stores_df)

with open(f"{BASE}/Granules Recipe - February 2026.xlsx", "rb") as f:
    granules_rates = load_granules_rates(io.BytesIO(f.read()))
tmp_mp = f"{BASE}/_mp_tmp.xlsx"
shutil.copy2(f"{BASE}/MEGA PACK.xlsx", tmp_mp)
with open(tmp_mp, "rb") as f:
    megapack_rates = load_megapack_rates(io.BytesIO(f.read()))
os.remove(tmp_mp)

# Load ground truth
tmp_gt = f"{BASE}/_gt_tmp.xlsx"
shutil.copy2(f"{BASE}/Jobtrack Feb With MRR.xlsx", tmp_gt)
wb_gt = openpyxl.load_workbook(tmp_gt, data_only=True)
ws_gt = wb_gt.active

# Load source (without MRR) for reading material data
wb_src = openpyxl.load_workbook(f"{BASE}/Jobtrack Feb Without MRR.xlsx", data_only=True)
ws_src = wb_src.active

REPORT_MONTH = "2-2026"

# Check what Issue Date column looks like in Stores
print("=" * 80)
print("STORES ISSUE DATE ANALYSIS")
print("=" * 80)
issue_col = None
for c in stores_df.columns:
    if 'issue' in str(c).lower() and 'date' in str(c).lower():
        issue_col = c
        break
if not issue_col:
    for c in stores_df.columns:
        if 'issue' in str(c).lower():
            issue_col = c
            break

if issue_col:
    print(f"Found Issue Date column: '{issue_col}'")
    dates = pd.to_datetime(stores_df[issue_col], errors='coerce')
    valid = dates.dropna()
    print(f"  Total rows: {len(stores_df)}, with valid date: {len(valid)}")
    if len(valid) > 0:
        print(f"  Date range: {valid.min()} to {valid.max()}")
        # How many in Feb 2026?
        feb26 = valid[(valid.dt.month == 2) & (valid.dt.year == 2026)]
        jan26 = valid[(valid.dt.month == 1) & (valid.dt.year == 2026)]
        older = valid[~((valid.dt.month.isin([1,2])) & (valid.dt.year == 2026))]
        print(f"  Feb 2026: {len(feb26)} rows")
        print(f"  Jan 2026: {len(jan26)} rows")
        print(f"  Older:    {len(older)} rows")
else:
    print("No Issue Date column found!")
    print(f"Columns: {list(stores_df.columns)}")

# ══════════════════════════════════════════════════════════════
# FULL SIMULATION: Friend's rule vs Current engine vs GT
# ══════════════════════════════════════════════════════════════
print("\n" + "=" * 80)
print("FULL CELL-BY-CELL COMPARISON")
print("=" * 80)

# Track results
results = {
    'current_engine': {'match': 0, 'mismatch': 0, 'total': 0, 'mismatches': []},
    'friend_rule':    {'match': 0, 'mismatch': 0, 'total': 0, 'mismatches': []},
}

def get_friend_rate(stores_df, pr_df, mat_name, mat_mic, mat_size, order_no, process,
                    mrr_sup_map, granules_rates, megapack_rates):
    """Friend's rule: ALL MRRs, weighted avg, no threshold, no date filter."""
    # Supplier override for INH
    if mat_name.upper() in ('WPE', 'WLDPE', 'PTD WPE'):
        order_upper = str(order_no).strip().upper()
        if granules_rates and order_upper in granules_rates:
            return granules_rates[order_upper], 'Granules'
        # Standard INH
        mrr_list = lookup_mrr(stores_df, mat_name, mat_mic, mat_size, order_no, 'PRINTING')
        if not mrr_list:
            mrr_list = lookup_mrr(stores_df, 'WPE', mat_mic, None, order_no, 'PRINTING')
        from engine.rate_lookup import lookup_film_rate
        rate = lookup_film_rate(pr_df, mrr_list, 'WPE', None, mat_mic) if mrr_list else 0
        if rate == 0:
            rate = lookup_material_rate_for_month(pr_df, 'WPE', mat_mic, REPORT_MONTH)
        if rate == 0:
            rate = lookup_material_rate_for_month(pr_df, 'TPE', mat_mic, REPORT_MONTH)
        return rate, 'INH/PR'

    # Find ALL MRRs
    mrr_qty = lookup_mrr_with_qty(stores_df, mat_name, mat_mic, mat_size, order_no, process)
    if not mrr_qty:
        mrr_qty = lookup_mrr_with_qty(stores_df, mat_name, mat_mic, None, order_no, process)
    if not mrr_qty:
        mrr_qty = lookup_mrr_with_qty(stores_df, mat_name, mat_mic, None, order_no)
    if not mrr_qty:
        return 0, 'NO_MRR'

    # Supplier override check
    mrr_numbers = list(mrr_qty.keys())
    supplier = get_supplier_for_mrrs(mrr_sup_map, mrr_numbers)

    if supplier == 'MEGA PACK' and megapack_rates:
        rate = lookup_megapack_rate(megapack_rates, mat_name, 2026, 2)
        if rate > 0:
            return rate, 'MEGA PACK'
    elif supplier in ('BANDERA', 'CYM') and granules_rates:
        order_upper = str(order_no).strip().upper()
        if order_upper in granules_rates:
            return granules_rates[order_upper], 'Granules'

    # Friend's rule: use ALL MRRs, filter by PR, weighted avg — NO 10% threshold
    mrr_qty_filtered = filter_mrr_by_pr(pr_df, mrr_qty.copy(), mat_name, mat_size, mat_mic)
    rate = lookup_film_rate_weighted(pr_df, mrr_qty_filtered, mat_name, mat_size, mat_mic)
    if rate == 0:
        rate = lookup_film_rate_weighted(pr_df, mrr_qty.copy(), mat_name, None, mat_mic)
    if rate == 0:
        rate = lookup_material_rate_for_month(pr_df, mat_name, mat_mic, REPORT_MONTH)
    # NO outlier check — friend says just use the weighted avg
    return rate, 'PR_ALL'


def get_current_engine_rate(stores_df, pr_df, mat_name, mat_mic, mat_size, order_no, process,
                            mrr_sup_map, granules_rates, megapack_rates):
    """Current engine logic (with 10% threshold + outlier check)."""
    if mat_name.upper() in ('WPE', 'WLDPE', 'PTD WPE'):
        order_upper = str(order_no).strip().upper()
        if granules_rates and order_upper in granules_rates:
            return granules_rates[order_upper], 'Granules'
        mrr_list = lookup_mrr(stores_df, mat_name, mat_mic, mat_size, order_no, 'PRINTING')
        if not mrr_list:
            mrr_list = lookup_mrr(stores_df, 'WPE', mat_mic, None, order_no, 'PRINTING')
        from engine.rate_lookup import lookup_film_rate
        rate = lookup_film_rate(pr_df, mrr_list, 'WPE', None, mat_mic) if mrr_list else 0
        if rate == 0:
            rate = lookup_material_rate_for_month(pr_df, 'WPE', mat_mic, REPORT_MONTH)
        if rate == 0:
            rate = lookup_material_rate_for_month(pr_df, 'TPE', mat_mic, REPORT_MONTH)
        return rate, 'INH/PR'

    mrr_qty = lookup_mrr_with_qty(stores_df, mat_name, mat_mic, mat_size, order_no, process)
    if not mrr_qty:
        mrr_qty = lookup_mrr_with_qty(stores_df, mat_name, mat_mic, None, order_no, process)
    if not mrr_qty:
        mrr_qty = lookup_mrr_with_qty(stores_df, mat_name, mat_mic, None, order_no)
    if not mrr_qty:
        return 0, 'NO_MRR'

    mrr_numbers = list(mrr_qty.keys())
    supplier = get_supplier_for_mrrs(mrr_sup_map, mrr_numbers)

    if supplier == 'MEGA PACK' and megapack_rates:
        rate = lookup_megapack_rate(megapack_rates, mat_name, 2026, 2)
        if rate > 0:
            return rate, 'MEGA PACK'
    elif supplier in ('BANDERA', 'CYM') and granules_rates:
        order_upper = str(order_no).strip().upper()
        if order_upper in granules_rates:
            return granules_rates[order_upper], 'Granules'

    mrr_qty = filter_mrr_by_pr(pr_df, mrr_qty, mat_name, mat_size, mat_mic)
    rate = lookup_film_rate_weighted(pr_df, mrr_qty, mat_name, mat_size, mat_mic)
    if rate == 0:
        rate = lookup_film_rate_weighted(pr_df, mrr_qty, mat_name, None, mat_mic)
    if rate == 0:
        rate = lookup_material_rate_for_month(pr_df, mat_name, mat_mic, REPORT_MONTH)
    # Outlier check
    if rate > 0:
        month_rate = lookup_material_rate_for_month(pr_df, mat_name, mat_mic, REPORT_MONTH)
        if month_rate > 0 and abs(rate - month_rate) / month_rate > 0.50:
            rate = month_rate
    return rate, 'PR_ENGINE'


# Process each row
material_slots = [
    ('Film', COLS['Input_Name'], COLS['Input_Size'], COLS['Input_Mic'],
     COLS['Film_Rate']),
    ('Fresh1', COLS['Fresh1_Name'], COLS['Fresh1_Size'], COLS['Fresh1_Mic'],
     COLS['Fresh1_Rate']),
    ('Fresh2', COLS['Fresh2_Name'], COLS['Fresh2_Size'], COLS['Fresh2_Mic'],
     COLS['Fresh2_Rate']),
]

for row in range(DATA_START_ROW, ws_gt.max_row + 1):
    uid = ws_gt.cell(row=row, column=1).value
    process = ws_gt.cell(row=row, column=COLS['Process']).value
    if not uid or not process:
        continue
    p = str(process).strip().upper()
    order_no = str(ws_gt.cell(row=row, column=COLS['Order_No']).value or '').strip()

    for slot_name, name_col, size_col, mic_col, rate_col in material_slots:
        gt_rate = ws_gt.cell(row=row, column=rate_col).value
        if gt_rate is None:
            continue

        gt_rate = float(gt_rate)
        mat_name = _safe_str(ws_gt.cell(row=row, column=name_col).value)
        mat_size = ws_gt.cell(row=row, column=size_col).value
        mat_mic = ws_gt.cell(row=row, column=mic_col).value

        if not mat_name:
            continue

        proc = p if slot_name == 'Film' else 'LAMINATION'

        # Current engine
        eng_rate, eng_src = get_current_engine_rate(
            stores_df, pr_df, mat_name, mat_mic, mat_size, order_no, proc,
            mrr_sup_map, granules_rates, megapack_rates)
        eng_match = abs(eng_rate - gt_rate) < TOLERANCE if eng_rate > 0 else False
        results['current_engine']['total'] += 1
        if eng_match:
            results['current_engine']['match'] += 1
        else:
            results['current_engine']['mismatch'] += 1
            results['current_engine']['mismatches'].append(
                (row, uid, slot_name, mat_name, gt_rate, eng_rate, eng_src))

        # Friend's rule
        fr_rate, fr_src = get_friend_rate(
            stores_df, pr_df, mat_name, mat_mic, mat_size, order_no, proc,
            mrr_sup_map, granules_rates, megapack_rates)
        fr_match = abs(fr_rate - gt_rate) < TOLERANCE if fr_rate > 0 else False
        results['friend_rule']['total'] += 1
        if fr_match:
            results['friend_rule']['match'] += 1
        else:
            results['friend_rule']['mismatch'] += 1
            results['friend_rule']['mismatches'].append(
                (row, uid, slot_name, mat_name, gt_rate, fr_rate, fr_src))

wb_gt.close()
wb_src.close()
os.remove(tmp_gt)

# ══════════════════════════════════════════════════════════════
# REPORT
# ══════════════════════════════════════════════════════════════
print("\n" + "=" * 80)
print("COMPARISON: Current Engine vs Friend's Rule")
print("=" * 80)

for label, data in results.items():
    total = data['total']
    match = data['match']
    pct = (match / total * 100) if total > 0 else 0
    print(f"\n{label}:")
    print(f"  Rate cells compared: {total}")
    print(f"  Matches:   {match} ({pct:.1f}%)")
    print(f"  Mismatches: {data['mismatch']}")

# Show differences
print("\n" + "=" * 80)
print("ROWS THAT CHANGE BETWEEN CURRENT ENGINE AND FRIEND'S RULE")
print("=" * 80)

eng_miss = {(r[0], r[2]) for r in results['current_engine']['mismatches']}
fr_miss = {(r[0], r[2]) for r in results['friend_rule']['mismatches']}

fixed = eng_miss - fr_miss
broken = fr_miss - eng_miss
same = eng_miss & fr_miss

if fixed:
    print(f"\n✅ FIXED by friend's rule ({len(fixed)} cells):")
    for row, slot in sorted(fixed):
        eng = [r for r in results['current_engine']['mismatches'] if r[0]==row and r[2]==slot][0]
        print(f"  Row {row} {slot}: GT={eng[4]:.4f}, Engine={eng[5]:.4f} → NOW MATCHES")

if broken:
    print(f"\n❌ BROKEN by friend's rule ({len(broken)} cells):")
    for row, slot in sorted(broken):
        fr = [r for r in results['friend_rule']['mismatches'] if r[0]==row and r[2]==slot][0]
        print(f"  Row {row} {slot}: GT={fr[4]:.4f}, Friend={fr[5]:.4f}")

if same:
    print(f"\n⚠️  STILL MISMATCHED in both ({len(same)} cells):")
    for row, slot in sorted(same):
        eng = [r for r in results['current_engine']['mismatches'] if r[0]==row and r[2]==slot][0]
        fr = [r for r in results['friend_rule']['mismatches'] if r[0]==row and r[2]==slot][0]
        print(f"  Row {row} {slot} ({eng[3]}): GT={eng[4]:.4f}, Engine={eng[5]:.4f}, Friend={fr[5]:.4f}")
