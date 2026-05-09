"""
CHECK: Does the Jobtrack Input Qty match the sum of friend's selected MRR quantities?
This could be the missing selection rule!

For each mismatched row, check:
- 1st Input Qty / Total 1st Input from Jobtrack
- Sum of friend's selected MRR quantities from Stores
- Sum of ALL MRR quantities vs selected subset
"""
import io, os, shutil, openpyxl, pandas as pd
from engine.fill_jobtrack import COLS, DATA_START_ROW, _safe_str
from engine.mrr_lookup import load_stores_recordings, lookup_mrr_with_qty
from engine.rate_lookup import load_purchase_register, lookup_film_rate_weighted, _find_col
from engine.supplier_rates import build_mrr_supplier_map

BASE = "Template_Files"

import shutil as _sh
_tmp_st = f"{BASE}/_stores_tmp.xlsx"
_sh.copy2(f"{BASE}/Stores Recordings.xlsx", _tmp_st)
stores_df = load_stores_recordings(_tmp_st)
_tmp_pr = f"{BASE}/_pr_tmp.xlsx"
_sh.copy2(f"{BASE}/Purchase Register - 2021 - 2026 _Feb 26.xlsx", _tmp_pr)
pr_df = load_purchase_register(_tmp_pr)
mrr_sup = build_mrr_supplier_map(stores_df)

# Load GT
tmp = f"{BASE}/_gt_qty.xlsx"
shutil.copy2(f"{BASE}/Jobtrack Feb With MRR.xlsx", tmp)
wb_gt = openpyxl.load_workbook(tmp, data_only=True)
ws_gt = wb_gt.active

print("=" * 90)
print("QTY MATCHING CHECK: Does Jobtrack qty = Sum of selected MRR quantities?")
print("=" * 90)

# Check ALL rows, not just mismatched ones
all_rows = []
for row in range(DATA_START_ROW, ws_gt.max_row + 1):
    uid = ws_gt.cell(row=row, column=1).value
    process = ws_gt.cell(row=row, column=COLS['Process']).value
    if not uid or not process:
        continue
    p = str(process).strip().upper()
    order = str(ws_gt.cell(row=row, column=COLS['Order_No']).value or '').strip()
    
    # Check Film (Printing) and Fresh1 (LAM)
    slots = []
    if 'PRINT' in p:
        slots.append(('Film', COLS['Input_Name'], COLS['Input_Size'], COLS['Input_Mic'],
                       COLS['Input_Qty'], 52, COLS['Total_1st_Input'],  # AY, AZ, BA
                       COLS['Film_MR'], COLS['Film_Rate'], 'PRINTING'))
    if 'LAM' in p:
        slots.append(('Fresh1', COLS['Fresh1_Name'], COLS['Fresh1_Size'], COLS['Fresh1_Mic'],
                       COLS['Fresh1_Qty'], COLS['Fresh1_Balance'], COLS['Total_Fresh1'],
                       COLS['Fresh1_MR'], COLS['Fresh1_Rate'], 'LAMINATION'))
        slots.append(('Fresh2', COLS['Fresh2_Name'], COLS['Fresh2_Size'], COLS['Fresh2_Mic'],
                       COLS.get('Fresh2_Qty', 85), COLS.get('Fresh2_Balance', 86), COLS.get('Total_Fresh2', 87),
                       COLS['Fresh2_MR'], COLS['Fresh2_Rate'], 'LAMINATION'))
    
    for slot, name_col, size_col, mic_col, qty_col, bal_col, total_col, mr_col, rate_col, proc in slots:
        gt_mr = ws_gt.cell(row=row, column=mr_col).value
        gt_rate = ws_gt.cell(row=row, column=rate_col).value
        if not gt_mr or not gt_rate:
            continue
        
        mat_name = _safe_str(ws_gt.cell(row=row, column=name_col).value)
        mat_size = ws_gt.cell(row=row, column=size_col).value
        mat_mic = ws_gt.cell(row=row, column=mic_col).value
        
        # Jobtrack quantities
        jt_qty = ws_gt.cell(row=row, column=qty_col).value  # Input Qty
        jt_bal = ws_gt.cell(row=row, column=bal_col).value   # Balance Qty  
        jt_total = ws_gt.cell(row=row, column=total_col).value  # Total (=Qty+Balance)
        
        # Parse GT MRR numbers
        gt_mrr_list = []
        for x in str(gt_mr).split('/'):
            try:
                gt_mrr_list.append(int(float(x.strip())))
            except:
                pass
        
        # Get ALL MRRs from Stores
        mrr_qty = lookup_mrr_with_qty(stores_df, mat_name, mat_mic, mat_size, order, proc)
        if not mrr_qty:
            mrr_qty = lookup_mrr_with_qty(stores_df, mat_name, mat_mic, None, order, proc)
        if not mrr_qty:
            mrr_qty = lookup_mrr_with_qty(stores_df, mat_name, mat_mic, None, order)
        
        # Sum of GT-selected MRR quantities
        gt_sum = sum(mrr_qty.get(m, 0) for m in gt_mrr_list)
        all_sum = sum(mrr_qty.values())
        
        all_rows.append({
            'row': row, 'uid': uid, 'order': order, 'slot': slot,
            'mat': mat_name, 'size': mat_size, 'mic': mat_mic,
            'jt_qty': jt_qty, 'jt_bal': jt_bal, 'jt_total': jt_total,
            'gt_mr': gt_mr, 'gt_rate': gt_rate,
            'gt_mrr_list': gt_mrr_list,
            'gt_sum_qty': gt_sum, 'all_sum_qty': all_sum,
            'all_mrrs': mrr_qty,
        })

# Print detailed comparison
print(f"\n{'Row':>4} {'Slot':>6} {'Order':>7} | {'JT Qty':>10} {'JT Bal':>10} {'JT Total':>10} | "
      f"{'GT MRR Sum':>10} {'All MRR Sum':>12} | {'JT=GT?':>8} {'JT=All?':>8} | MRRs")
print("-" * 130)

for d in all_rows:
    jt_q = float(d['jt_qty'] or 0)
    jt_b = float(d['jt_bal'] or 0)
    jt_t = float(d['jt_total'] or 0)
    
    gt_match = abs(jt_t - d['gt_sum_qty']) < 1.0 if jt_t and d['gt_sum_qty'] else False
    all_match = abs(jt_t - d['all_sum_qty']) < 1.0 if jt_t and d['all_sum_qty'] else False
    
    gt_sym = "✅" if gt_match else "❌"
    all_sym = "✅" if all_match else "❌"
    
    # Highlight mismatched rows
    flag = " *** MISMATCH ROW ***" if d['row'] in (42, 43, 45, 46) else ""
    
    print(f"{d['row']:>4} {d['slot']:>6} {d['order']:>7} | "
          f"{jt_q:>10.1f} {jt_b:>10.1f} {jt_t:>10.1f} | "
          f"{d['gt_sum_qty']:>10.1f} {d['all_sum_qty']:>12.1f} | "
          f"{gt_sym:>8} {all_sym:>8} | {d['gt_mr']}{flag}")

# Summary
print("\n" + "=" * 90)
print("SUMMARY")
print("=" * 90)
gt_matches = sum(1 for d in all_rows if abs(float(d['jt_total'] or 0) - d['gt_sum_qty']) < 1.0)
all_matches = sum(1 for d in all_rows if abs(float(d['jt_total'] or 0) - d['all_sum_qty']) < 1.0)
total = len(all_rows)
print(f"JT Total = Sum of GT-selected MRRs:  {gt_matches}/{total}")
print(f"JT Total = Sum of ALL MRRs:          {all_matches}/{total}")

# For mismatched rows, deep dive
print("\n" + "=" * 90)
print("DEEP DIVE: MISMATCHED ROWS — QTY ANALYSIS")
print("=" * 90)
for d in all_rows:
    if d['row'] not in (42, 43, 45, 46):
        continue
    
    print(f"\nRow {d['row']} | {d['uid']} | {d['slot']}={d['mat']} | Order={d['order']}")
    print(f"  Jobtrack: Qty={d['jt_qty']}, Balance={d['jt_bal']}, Total={d['jt_total']}")
    print(f"  GT MRR: {d['gt_mr']}")
    print(f"  GT MRR qty sum: {d['gt_sum_qty']}")
    print(f"  ALL MRR qty sum: {d['all_sum_qty']}")
    
    # Which MRRs qty subset sums to JT total?
    jt_t = float(d['jt_total'] or 0)
    if jt_t > 0:
        import itertools
        mrr_keys = list(d['all_mrrs'].keys())
        print(f"\n  Which MRR subsets sum to JT Total ({jt_t})?")
        found = False
        for size in range(1, len(mrr_keys) + 1):
            for combo in itertools.combinations(mrr_keys, size):
                combo_sum = sum(d['all_mrrs'][m] for m in combo)
                if abs(combo_sum - jt_t) < 2.0:
                    # Calculate rate for this subset
                    w_rate = sum(
                        lookup_film_rate_weighted(pr_df, {m: d['all_mrrs'][m]}, d['mat'], d['size'], d['mic']) * d['all_mrrs'][m]
                        for m in combo
                    ) / combo_sum
                    rate_match = abs(w_rate - float(d['gt_rate'])) < 0.01
                    print(f"    ✓ MRRs {combo}: sum={combo_sum:.1f} "
                          f"rate={w_rate:.6f} {'= GT ✅' if rate_match else '≠ GT'}")
                    found = True
        if not found:
            print(f"    No subset sums to {jt_t}")
            # Show closest
            for size in range(1, len(mrr_keys) + 1):
                for combo in itertools.combinations(mrr_keys, size):
                    combo_sum = sum(d['all_mrrs'][m] for m in combo)
                    diff = abs(combo_sum - jt_t)
                    if diff < 100:
                        print(f"    ~ MRRs {combo}: sum={combo_sum:.1f} (diff={diff:.1f} from JT)")

wb_gt.close()
os.remove(tmp)
