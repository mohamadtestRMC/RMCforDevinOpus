"""
VERIFY: If we strictly match by Size from Stores, do we get the correct MRRs?
Check all 4 problem rows AND all matching rows.
"""
import shutil, os, openpyxl, pandas as pd
from engine.fill_jobtrack import COLS, DATA_START_ROW, _safe_str
from engine.rate_lookup import load_purchase_register, lookup_film_rate_weighted, _find_col

BASE = "Template_Files"
shutil.copy2(f"{BASE}/Stores Recordings.xlsx", f"{BASE}/_st5.xlsx")
shutil.copy2(f"{BASE}/Purchase Register - 2021 - 2026 _Feb 26.xlsx", f"{BASE}/_pr5.xlsx")
shutil.copy2(f"{BASE}/Jobtrack Feb With MRR.xlsx", f"{BASE}/_gt5.xlsx")

stores_df = pd.read_excel(f"{BASE}/_st5.xlsx", sheet_name=0, header=1)
pr_df = load_purchase_register(f"{BASE}/_pr5.xlsx")

# Column mapping
mic_c = [c for c in stores_df.columns if 'mic' in str(c).lower()][0]
width_c = [c for c in stores_df.columns if 'width' in str(c).lower()][0]
qty_c = [c for c in stores_df.columns if 'issue' in str(c).lower() and 'qty' in str(c).lower()][0]
mat_c = [c for c in stores_df.columns if 'sub' in str(c).lower()][0]
proc_c = [c for c in stores_df.columns if 'issue' in str(c).lower() and 'process' in str(c).lower()][0]
wo_c = [c for c in stores_df.columns if 'issue' in str(c).lower() and 'wo' in str(c).lower()][0]
mrr_c = [c for c in stores_df.columns if 'm.r.r' in str(c).lower() and 'no' in str(c).lower()][0]
sup_c = [c for c in stores_df.columns if 'supplier' in str(c).lower()][0]

wb_gt = openpyxl.load_workbook(f"{BASE}/_gt5.xlsx", data_only=True)
ws_gt = wb_gt.active

print("=" * 80)
print("SIZE-STRICT MRR LOOKUP: Do MRRs match when we filter by EXACT size?")
print("=" * 80)

all_results = []

for row in range(DATA_START_ROW, ws_gt.max_row + 1):
    uid = ws_gt.cell(row=row, column=1).value
    process = ws_gt.cell(row=row, column=COLS['Process']).value
    if not uid or not process:
        continue
    p = str(process).strip().upper()
    order = str(ws_gt.cell(row=row, column=COLS['Order_No']).value or '').strip()
    
    slots = []
    if 'PRINT' in p:
        slots.append(('Film', COLS['Input_Name'], COLS['Input_Size'], COLS['Input_Mic'],
                       COLS['Film_MR'], COLS['Film_Rate'], 'PRINTING'))
    if 'LAM' in p:
        slots.append(('Fresh1', COLS['Fresh1_Name'], COLS['Fresh1_Size'], COLS['Fresh1_Mic'],
                       COLS['Fresh1_MR'], COLS['Fresh1_Rate'], 'LAMINATION'))
        slots.append(('Fresh2', COLS['Fresh2_Name'], COLS['Fresh2_Size'], COLS['Fresh2_Mic'],
                       COLS['Fresh2_MR'], COLS['Fresh2_Rate'], 'LAMINATION'))
    
    for slot, name_col, size_col, mic_col, mr_col, rate_col, proc in slots:
        gt_mr = ws_gt.cell(row=row, column=mr_col).value
        gt_rate = ws_gt.cell(row=row, column=rate_col).value
        if not gt_mr or not gt_rate:
            continue
        
        mat_name = _safe_str(ws_gt.cell(row=row, column=name_col).value)
        mat_size = ws_gt.cell(row=row, column=size_col).value
        mat_mic = ws_gt.cell(row=row, column=mic_col).value
        gt_rate = float(gt_rate)
        
        # Parse GT MRRs
        gt_mrrs = set()
        for x in str(gt_mr).split('/'):
            try: gt_mrrs.add(int(float(x.strip())))
            except: pass
        
        # SIZE-STRICT lookup from stores
        mask = pd.Series([True] * len(stores_df))
        # Exact size
        if mat_size:
            mask &= pd.to_numeric(stores_df[width_c], errors='coerce') == float(mat_size)
        # Exact mic
        if mat_mic:
            mask &= pd.to_numeric(stores_df[mic_c], errors='coerce') == float(mat_mic)
        # Material contains
        mat_upper = str(mat_name).strip().upper()
        mask &= stores_df[mat_c].astype(str).str.upper().str.contains(mat_upper, na=False)
        # Process
        proc_key = 'PRINT' if proc == 'PRINTING' else 'LAM'
        mask &= stores_df[proc_c].astype(str).str.upper().str.contains(proc_key, na=False)
        # Work order
        wo_mask = stores_df[wo_c].astype(str).str.strip().str.upper() == order.upper()
        mask &= wo_mask
        # Has issue qty
        mask &= pd.to_numeric(stores_df[qty_c], errors='coerce') > 0
        
        matched = stores_df[mask]
        
        # Get unique MRRs and sum qty per MRR
        mrr_qty = {}
        for _, sr in matched.iterrows():
            m = int(float(sr[mrr_c]))
            q = float(sr[qty_c])
            mrr_qty[m] = mrr_qty.get(m, 0) + q
        
        found_mrrs = set(mrr_qty.keys())
        mrr_match = found_mrrs == gt_mrrs
        
        # Calculate weighted rate with size-strict MRRs
        if mrr_qty:
            total_q = sum(mrr_qty.values())
            weighted = 0
            for m, q in mrr_qty.items():
                r = lookup_film_rate_weighted(pr_df, {m: q}, mat_name, mat_size, mat_mic)
                if r == 0:
                    r = lookup_film_rate_weighted(pr_df, {m: q}, mat_name, None, mat_mic)
                weighted += r * q
            weighted = weighted / total_q if total_q > 0 else 0
        else:
            weighted = 0
        
        rate_match = abs(weighted - gt_rate) < 0.01 if weighted > 0 else False
        
        flag = " ***" if row in (42, 43, 45, 46) else ""
        status = "✅" if mrr_match and rate_match else ("⚠️" if mrr_match else "❌")
        
        all_results.append({
            'row': row, 'slot': slot, 'order': order,
            'gt_mrrs': gt_mrrs, 'found_mrrs': found_mrrs,
            'mrr_match': mrr_match, 'rate_match': rate_match,
            'gt_rate': gt_rate, 'computed': weighted, 'mrr_qty': mrr_qty
        })
        
        print(f"{status} Row {row:>3} {slot:>6} {order:>7} | "
              f"GT_MRRs={sorted(gt_mrrs)} | Found={sorted(found_mrrs)} | "
              f"MR={'✅' if mrr_match else '❌'} | "
              f"Rate: GT={gt_rate:.4f} Calc={weighted:.4f} {'✅' if rate_match else '❌'}{flag}")

# Summary
print("\n" + "=" * 80)
print("SUMMARY WITH SIZE-STRICT LOOKUP")
print("=" * 80)
mrr_matches = sum(1 for r in all_results if r['mrr_match'])
rate_matches = sum(1 for r in all_results if r['rate_match'])
both_matches = sum(1 for r in all_results if r['mrr_match'] and r['rate_match'])
total = len(all_results)
print(f"MRR match:  {mrr_matches}/{total} ({mrr_matches/total*100:.1f}%)")
print(f"Rate match: {rate_matches}/{total} ({rate_matches/total*100:.1f}%)")
print(f"Both match: {both_matches}/{total} ({both_matches/total*100:.1f}%)")

wb_gt.close()
for f in [f"{BASE}/_st5.xlsx", f"{BASE}/_pr5.xlsx", f"{BASE}/_gt5.xlsx"]:
    try: os.remove(f)
    except: pass
