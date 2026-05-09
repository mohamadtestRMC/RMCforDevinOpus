"""
COMPLETE GAP ANALYSIS: Check EVERY cell in engine output vs ground truth.
For each mismatch, trace the exact root cause.
"""
import io, os, shutil, openpyxl, pandas as pd
from engine.fill_jobtrack import (
    fill_jobtrack, COLS, DATA_START_ROW, _safe_str, _safe_float
)
from engine.mrr_lookup import load_stores_recordings, lookup_mrr_with_qty
from engine.rate_lookup import (
    load_purchase_register, lookup_film_rate_weighted,
    filter_mrr_by_pr, _find_col
)
from engine.supplier_rates import build_mrr_supplier_map

BASE = "Template_Files"
TOLERANCE_RATE = 0.01
TOLERANCE_VALUE = 1.0

# Run the actual engine with all files
with open(f"{BASE}/Jobtrack Feb Without MRR.xlsx", "rb") as f:
    jt = io.BytesIO(f.read())
with open(f"{BASE}/Stores Recordings.xlsx", "rb") as f:
    stores = io.BytesIO(f.read())
with open(f"{BASE}/Purchase Register - 2021 - 2026 _Feb 26.xlsx", "rb") as f:
    pr = io.BytesIO(f.read())
with open(f"{BASE}/Granules Recipe - February 2026.xlsx", "rb") as f:
    granules = io.BytesIO(f.read())
tmp_mp = f"{BASE}/_mp_tmp2.xlsx"
shutil.copy2(f"{BASE}/MEGA PACK.xlsx", tmp_mp)
with open(tmp_mp, "rb") as f:
    megapack = io.BytesIO(f.read())
os.remove(tmp_mp)

print("Running engine...")
output_bytes, results_log, fill_stats = fill_jobtrack(
    jt, stores, pr, granules_file=granules, megapack_file=megapack
)
output_bytes.seek(0)
wb_eng = openpyxl.load_workbook(output_bytes, data_only=False)
ws_eng = wb_eng.active

# Load ground truth
tmp_gt = f"{BASE}/_gt_tmp3.xlsx"
shutil.copy2(f"{BASE}/Jobtrack Feb With MRR.xlsx", tmp_gt)
wb_gt = openpyxl.load_workbook(tmp_gt, data_only=True)
ws_gt = wb_gt.active

# Also load stores for root cause analysis
stores_df = load_stores_recordings(f"{BASE}/Stores Recordings.xlsx")
pr_df = load_purchase_register(f"{BASE}/Purchase Register - 2021 - 2026 _Feb 26.xlsx")
mrr_sup = build_mrr_supplier_map(stores_df)

# All columns to compare
compare_cols = [
    ('Film_MR',    COLS['Film_MR'],    'MR'),
    ('Film_Rate',  COLS['Film_Rate'],  'Rate'),
    ('Film_Value', COLS['Film_Value'], 'Value'),
    ('Fresh1_MR',    COLS['Fresh1_MR'],    'MR'),
    ('Fresh1_Rate',  COLS['Fresh1_Rate'],  'Rate'),
    ('Fresh1_Value', COLS['Fresh1_Value'], 'Value'),
    ('Fresh2_MR',    COLS['Fresh2_MR'],    'MR'),
    ('Fresh2_Rate',  COLS['Fresh2_Rate'],  'Rate'),
    ('Fresh2_Value', COLS['Fresh2_Value'], 'Value'),
    ('Adh_Rate',  COLS['Adh_Rate'],  'Rate'),
    ('Adh_Value', COLS['Adh_Value'], 'Value'),
    ('Hard_Rate',  COLS['Hard_Rate'],  'Rate'),
    ('Hard_Value', COLS['Hard_Value'], 'Value'),
    ('Sol_Rate',  COLS['Sol_Rate'],  'Rate'),
    ('Sol_Value', COLS['Sol_Value'], 'Value'),
]

# Track everything
all_match = 0
all_mismatch = 0
all_total = 0
mismatches = []
category_stats = {}

for row in range(DATA_START_ROW, ws_gt.max_row + 1):
    uid = ws_gt.cell(row=row, column=1).value
    process = ws_gt.cell(row=row, column=COLS['Process']).value
    if not uid or not process:
        continue
    p = str(process).strip().upper()
    order = str(ws_gt.cell(row=row, column=COLS['Order_No']).value or '').strip()

    for col_name, col_idx, col_type in compare_cols:
        gt_val = ws_gt.cell(row=row, column=col_idx).value
        eng_val = ws_eng.cell(row=row, column=col_idx).value

        if gt_val is None:
            continue

        all_total += 1
        cat = f"{p}_{col_name}"
        if cat not in category_stats:
            category_stats[cat] = {'match': 0, 'mismatch': 0, 'total': 0}
        category_stats[cat]['total'] += 1

        # Compare
        match = False
        if col_type == 'MR':
            if eng_val is None:
                match = False
            else:
                gt_parts = set(str(gt_val).strip().split('/'))
                eng_parts = set(str(eng_val).strip().split('/'))
                match = gt_parts == eng_parts
        else:
            if eng_val is None:
                match = False
            else:
                tol = TOLERANCE_RATE if col_type == 'Rate' else TOLERANCE_VALUE
                try:
                    match = abs(float(eng_val) - float(gt_val)) <= tol
                except:
                    match = False

        if match:
            all_match += 1
            category_stats[cat]['match'] += 1
        else:
            all_mismatch += 1
            category_stats[cat]['mismatch'] += 1
            mismatches.append({
                'row': row, 'uid': uid, 'process': p, 'order': order,
                'col': col_name, 'type': col_type,
                'gt': gt_val, 'eng': eng_val
            })

wb_eng.close()
wb_gt.close()
os.remove(tmp_gt)

# ══════════════════════════════════════════════════════════════
# REPORT
# ══════════════════════════════════════════════════════════════
print("\n" + "=" * 80)
print("COMPLETE GAP ANALYSIS — ALL CELLS")
print("=" * 80)
pct = (all_match / all_total * 100) if all_total else 0
print(f"\nOverall: {all_match}/{all_total} match ({pct:.1f}%)")
print(f"Mismatches: {all_mismatch}")

print("\n" + "-" * 80)
print("CATEGORY BREAKDOWN")
print("-" * 80)
for cat in sorted(category_stats):
    s = category_stats[cat]
    p = (s['match'] / s['total'] * 100) if s['total'] else 0
    status = "✅" if s['mismatch'] == 0 else "❌"
    print(f"  {status} {cat:35s}: {s['match']:3d}/{s['total']:3d} ({p:5.1f}%)  "
          f"{'— ' + str(s['mismatch']) + ' mismatch' if s['mismatch'] else ''}")

print("\n" + "-" * 80)
print(f"ALL {all_mismatch} MISMATCHES — DETAILED")
print("-" * 80)

for m in mismatches:
    # Get material info for root cause
    if 'Film' in m['col']:
        name_col = COLS['Input_Name']
        size_col = COLS['Input_Size']
        mic_col = COLS['Input_Mic']
        proc = m['process']
    elif 'Fresh1' in m['col']:
        name_col = COLS['Fresh1_Name']
        size_col = COLS['Fresh1_Size']
        mic_col = COLS['Fresh1_Mic']
        proc = 'LAMINATION'
    elif 'Fresh2' in m['col']:
        name_col = COLS['Fresh2_Name']
        size_col = COLS['Fresh2_Size']
        mic_col = COLS['Fresh2_Mic']
        proc = 'LAMINATION'
    else:
        name_col = size_col = mic_col = None
        proc = None

    diff_str = ""
    if m['type'] in ('Rate', 'Value') and m['eng'] is not None:
        try:
            diff = abs(float(m['eng']) - float(m['gt']))
            pct_d = (diff / abs(float(m['gt'])) * 100) if float(m['gt']) != 0 else 0
            diff_str = f"  Diff={diff:.4f} ({pct_d:.2f}%)"
        except:
            pass

    print(f"\n  Row {m['row']} | {m['uid']} | {m['process']} | {m['order']}")
    print(f"    {m['col']}: GT={m['gt']}, Engine={m['eng']}{diff_str}")

    # Root cause for rate mismatches
    if m['type'] == 'Rate' and name_col:
        # Re-read from GT workbook (need to reopen)
        tmp_gt2 = f"{BASE}/_gt_tmp4.xlsx"
        shutil.copy2(f"{BASE}/Jobtrack Feb With MRR.xlsx", tmp_gt2)
        wb2 = openpyxl.load_workbook(tmp_gt2, data_only=True)
        ws2 = wb2.active
        
        mat_name = _safe_str(ws2.cell(row=m['row'], column=name_col).value)
        mat_size = ws2.cell(row=m['row'], column=size_col).value
        mat_mic = ws2.cell(row=m['row'], column=mic_col).value
        gt_mr = ws2.cell(row=m['row'], column=COLS.get(m['col'].replace('_Rate','_MR'), 0)).value
        
        wb2.close()
        os.remove(tmp_gt2)

        # Find MRRs
        mrr_qty = lookup_mrr_with_qty(stores_df, mat_name, mat_mic, mat_size, m['order'], proc)
        if not mrr_qty:
            mrr_qty = lookup_mrr_with_qty(stores_df, mat_name, mat_mic, None, m['order'], proc)
        if not mrr_qty:
            mrr_qty = lookup_mrr_with_qty(stores_df, mat_name, mat_mic, None, m['order'])

        if mrr_qty:
            # Show each MRR with supplier and rate
            print(f"    Material: {mat_name}, Size={mat_size}, Mic={mat_mic}")
            print(f"    GT MR#: {gt_mr}")
            print(f"    All MRRs found ({len(mrr_qty)}):")
            for mrr_num, qty in sorted(mrr_qty.items()):
                sup = mrr_sup.get(int(mrr_num), '?')
                rate = lookup_film_rate_weighted(pr_df, {mrr_num: qty}, mat_name, mat_size, mat_mic)
                if rate == 0:
                    rate = lookup_film_rate_weighted(pr_df, {mrr_num: qty}, mat_name, None, mat_mic)
                print(f"      MRR {mrr_num}: Qty={qty}, Rate={rate}, Supplier={sup}")

            # Weighted avg of ALL
            filtered = filter_mrr_by_pr(pr_df, mrr_qty.copy(), mat_name, mat_size, mat_mic)
            w_all = lookup_film_rate_weighted(pr_df, filtered, mat_name, mat_size, mat_mic)
            if w_all == 0:
                w_all = lookup_film_rate_weighted(pr_df, mrr_qty.copy(), mat_name, None, mat_mic)
            print(f"    Weighted avg ALL: {w_all:.6f}")
            
            # If GT used subset, compute its weighted avg
            if gt_mr:
                gt_mrrs = [int(float(x.strip())) for x in str(gt_mr).split('/') if x.strip()]
                gt_total_qty = sum(mrr_qty.get(m2, 0) for m2 in gt_mrrs)
                if gt_total_qty > 0:
                    gt_w = sum(
                        lookup_film_rate_weighted(pr_df, {m2: mrr_qty.get(m2, 1)}, mat_name, mat_size, mat_mic) * mrr_qty.get(m2, 0)
                        for m2 in gt_mrrs
                    ) / gt_total_qty
                    print(f"    Weighted avg GT subset ({gt_mr}): {gt_w:.6f}")
                    if abs(gt_w - float(m['gt'])) < 0.01:
                        print(f"    ➡️  ROOT CAUSE: GT used SUBSET of MRRs, engine used ALL")
                    else:
                        print(f"    ➡️  ROOT CAUSE: GT rate doesn't match even its own MRR subset")

    elif m['type'] == 'Value' and m['eng'] is not None:
        # Value mismatch — usually caused by rate mismatch or qty difference
        print(f"    ➡️  Value mismatch follows from rate mismatch in same row")

    elif m['type'] == 'MR' :
        print(f"    ➡️  MR# display difference (subset vs full list)")

# Summary
print("\n" + "=" * 80)
print("SUMMARY")
print("=" * 80)
print(f"Total cells compared: {all_total}")
print(f"Matches:   {all_match} ({all_match/all_total*100:.1f}%)")
print(f"Mismatches: {all_mismatch} ({all_mismatch/all_total*100:.1f}%)")
print(f"\nBy type:")
rate_m = sum(1 for m in mismatches if m['type'] == 'Rate')
val_m = sum(1 for m in mismatches if m['type'] == 'Value')
mr_m = sum(1 for m in mismatches if m['type'] == 'MR')
print(f"  Rate mismatches:  {rate_m}")
print(f"  Value mismatches: {val_m}")
print(f"  MR# mismatches:   {mr_m}")
