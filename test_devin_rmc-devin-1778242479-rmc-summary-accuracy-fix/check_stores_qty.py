"""
Check if Jobtrack "1st Input Qty" matches individual Store rows in column S.
The idea: match Stores rows where qty = JT Input Qty → those rows give us the MRR.
"""
import shutil, os, openpyxl, pandas as pd
from engine.fill_jobtrack import COLS, DATA_START_ROW, _safe_str
from engine.mrr_lookup import load_stores_recordings
from engine.rate_lookup import load_purchase_register, lookup_film_rate_weighted, _find_col

BASE = "Template_Files"

# Copy locked files
shutil.copy2(f"{BASE}/Stores Recordings.xlsx", f"{BASE}/_st2.xlsx")
shutil.copy2(f"{BASE}/Purchase Register - 2021 - 2026 _Feb 26.xlsx", f"{BASE}/_pr2.xlsx")
shutil.copy2(f"{BASE}/Jobtrack Feb With MRR.xlsx", f"{BASE}/_gt2.xlsx")

stores_df = load_stores_recordings(f"{BASE}/_st2.xlsx")
pr_df = load_purchase_register(f"{BASE}/_pr2.xlsx")

# First: show ALL Stores columns to find column S
print("=" * 80)
print("STORES FILE COLUMNS")
print("=" * 80)
for i, c in enumerate(stores_df.columns):
    print(f"  Col {i}: '{c}'")

# Also load raw to check column S (letter-based)
wb_st = openpyxl.load_workbook(f"{BASE}/_st2.xlsx", data_only=True)
ws_st = wb_st.active
print(f"\nRaw column S header (row 2): {ws_st.cell(row=2, column=19).value}")
print(f"Raw column S sample values:")
for r in range(3, 10):
    print(f"  Row {r}: {ws_st.cell(row=r, column=19).value}")

# Column S = column 19 (letter S = 19th)
# Let's also check nearby columns
print(f"\nColumns around S:")
for col in range(17, 22):
    letter = openpyxl.utils.get_column_letter(col)
    val = ws_st.cell(row=2, column=col).value
    print(f"  Col {letter} (idx {col}): header='{val}'")
    for r in range(3, 6):
        print(f"    Row {r}: {ws_st.cell(row=r, column=col).value}")

wb_st.close()

# Now identify the Issue Qty column in the dataframe
qty_col = None
for c in stores_df.columns:
    cl = str(c).lower()
    if 'issue' in cl and ('qty' in cl or 'q' in cl) and 'date' not in cl:
        qty_col = c
        break
if not qty_col:
    for c in stores_df.columns:
        if 'qty' in str(c).lower():
            qty_col = c
            break

# Also get MRR col, WO col, process col, material col, size, mic
mrr_col = None
for c in stores_df.columns:
    if 'm.r.r' in str(c).lower() or 'mrr' in str(c).lower():
        if 'date' not in str(c).lower():
            mrr_col = c
            break
if not mrr_col:
    for c in stores_df.columns:
        if 'm.r.r' in str(c).lower():
            mrr_col = c
            break

wo_col = None
for c in stores_df.columns:
    cl = str(c).lower()
    if 'w/o' in cl or ('work' in cl and 'order' in cl) or 'w.o' in cl:
        wo_col = c
        break

process_col = None
for c in stores_df.columns:
    if 'process' in str(c).lower():
        process_col = c
        break

sub_col = None
for c in stores_df.columns:
    if 'sub' in str(c).lower() and 'cat' in str(c).lower():
        sub_col = c
        break

size_col = None
for c in stores_df.columns:
    if 'size' in str(c).lower() or 'width' in str(c).lower():
        size_col = c
        break

mic_col = None
for c in stores_df.columns:
    if 'mic' in str(c).lower():
        mic_col = c
        break

print(f"\nIdentified columns:")
print(f"  Qty: {qty_col}")
print(f"  MRR: {mrr_col}")
print(f"  WO:  {wo_col}")
print(f"  Process: {process_col}")
print(f"  Sub Cat: {sub_col}")
print(f"  Size: {size_col}")
print(f"  Mic: {mic_col}")

# Load GT
wb_gt = openpyxl.load_workbook(f"{BASE}/_gt2.xlsx", data_only=True)
ws_gt = wb_gt.active

# For each mismatched row, find Stores rows that match on material/size/mic/order/process
# AND check if the qty matches the JT Input Qty
print("\n" + "=" * 80)
print("MATCHING STORES ROWS FOR MISMATCHED JOBTRACK ROWS")
print("=" * 80)

for row in [42, 43, 45, 46]:
    uid = ws_gt.cell(row=row, column=1).value
    process = str(ws_gt.cell(row=row, column=COLS['Process']).value or '').strip().upper()
    order = str(ws_gt.cell(row=row, column=COLS['Order_No']).value or '').strip()
    
    if 'PRINT' in process:
        slot = 'Film'
        name_col = COLS['Input_Name']
        size_jt = COLS['Input_Size']
        mic_jt = COLS['Input_Mic']
        qty_jt = COLS['Input_Qty']      # 1st Input Qty
        total_jt = COLS['Total_1st_Input']
        mr_col = COLS['Film_MR']
        rate_col = COLS['Film_Rate']
    else:
        slot = 'Fresh1'
        name_col = COLS['Fresh1_Name']
        size_jt = COLS['Fresh1_Size']
        mic_jt = COLS['Fresh1_Mic']
        qty_jt = COLS['Fresh1_Qty']
        total_jt = COLS['Total_Fresh1']
        mr_col = COLS['Fresh1_MR']
        rate_col = COLS['Fresh1_Rate']
    
    mat_name = _safe_str(ws_gt.cell(row=row, column=name_col).value)
    mat_size = ws_gt.cell(row=row, column=size_jt).value
    mat_mic = ws_gt.cell(row=row, column=mic_jt).value
    jt_qty = float(ws_gt.cell(row=row, column=qty_jt).value or 0)
    jt_total = float(ws_gt.cell(row=row, column=total_jt).value or 0)
    gt_mr = ws_gt.cell(row=row, column=mr_col).value
    gt_rate = float(ws_gt.cell(row=row, column=rate_col).value or 0)
    
    print(f"\n{'='*70}")
    print(f"Row {row} | {uid} | {slot}={mat_name} Size={mat_size} Mic={mat_mic}")
    print(f"  Order={order} | Process={process}")
    print(f"  JT Input Qty={jt_qty} | JT Total={jt_total}")
    print(f"  GT MR#={gt_mr} | GT Rate={gt_rate}")
    print(f"{'='*70}")
    
    # Find ALL matching Stores rows
    mask = pd.Series([True] * len(stores_df))
    
    # Filter by work order
    if wo_col:
        mask &= stores_df[wo_col].astype(str).str.strip().str.upper() == order.upper()
    
    # Filter by process
    if process_col:
        proc_match = 'PRINT' if 'PRINT' in process else 'LAM'
        mask &= stores_df[process_col].astype(str).str.strip().str.upper().str.contains(proc_match, na=False)
    
    # Filter by material
    if sub_col:
        mat_upper = mat_name.upper()
        mask &= stores_df[sub_col].astype(str).str.strip().str.upper().str.contains(mat_upper, na=False)
    
    matched = stores_df[mask].copy()
    
    print(f"\n  Matching Stores rows ({len(matched)}):")
    print(f"  {'MRR':>8} {'Qty':>10} {'Process':>12} {'Material':>15} {'Size':>8} {'Mic':>6}")
    print(f"  {'-'*65}")
    
    for _, sr in matched.iterrows():
        mrr = sr.get(mrr_col, '?')
        qty = sr.get(qty_col, '?')
        proc = sr.get(process_col, '?')
        mat = sr.get(sub_col, '?')
        sz = sr.get(size_col, '?')
        mc = sr.get(mic_col, '?')
        
        # Highlight if qty matches JT Input Qty
        qty_float = float(qty) if pd.notna(qty) else 0
        match_flag = " ← MATCHES JT Input Qty!" if abs(qty_float - jt_qty) < 1.0 else ""
        
        print(f"  {str(mrr):>8} {str(qty):>10} {str(proc):>12} {str(mat):>15} {str(sz):>8} {str(mc):>6}{match_flag}")
    
    # Check: which Stores rows have qty that sum to JT Input Qty?
    print(f"\n  Checking which MRR rows SUM to JT Input Qty ({jt_qty}):")
    import itertools
    store_rows = []
    for _, sr in matched.iterrows():
        mrr = sr.get(mrr_col, None)
        qty = float(sr.get(qty_col, 0)) if pd.notna(sr.get(qty_col)) else 0
        if mrr and qty > 0:
            store_rows.append((mrr, qty))
    
    # Try all subsets
    found_any = False
    for size in range(1, min(len(store_rows) + 1, 8)):
        for combo in itertools.combinations(range(len(store_rows)), size):
            combo_sum = sum(store_rows[i][1] for i in combo)
            if abs(combo_sum - jt_qty) < 2.0:
                mrrs_used = [store_rows[i][0] for i in combo]
                qtys_used = [store_rows[i][1] for i in combo]
                
                # Calculate weighted rate for this subset
                mrr_qty_dict = {}
                for i in combo:
                    m = int(float(store_rows[i][0]))
                    mrr_qty_dict[m] = mrr_qty_dict.get(m, 0) + store_rows[i][1]
                
                total_q = sum(mrr_qty_dict.values())
                w_rate = 0
                for m, q in mrr_qty_dict.items():
                    r = lookup_film_rate_weighted(pr_df, {m: q}, mat_name, mat_size, mat_mic)
                    if r == 0:
                        r = lookup_film_rate_weighted(pr_df, {m: q}, mat_name, None, mat_mic)
                    w_rate += r * q
                w_rate = w_rate / total_q if total_q > 0 else 0
                
                rate_match = abs(w_rate - gt_rate) < 0.01
                print(f"    ✓ Rows: {list(zip(mrrs_used, qtys_used))}")
                print(f"      Sum={combo_sum:.1f}, Rate={w_rate:.6f} {'= GT ✅' if rate_match else f'≠ GT ({gt_rate})'}")
                found_any = True
    
    if not found_any:
        print(f"    No exact match found. Closest subsets:")
        closest = []
        for size in range(1, min(len(store_rows) + 1, 6)):
            for combo in itertools.combinations(range(len(store_rows)), size):
                combo_sum = sum(store_rows[i][1] for i in combo)
                diff = abs(combo_sum - jt_qty)
                if diff < 200:
                    mrrs_used = [(store_rows[i][0], store_rows[i][1]) for i in combo]
                    closest.append((diff, combo_sum, mrrs_used))
        closest.sort()
        for diff, s, mrrs in closest[:5]:
            print(f"    ~ Sum={s:.1f} (diff={diff:.1f}): {mrrs}")

wb_gt.close()
# Cleanup
for f in [f"{BASE}/_st2.xlsx", f"{BASE}/_pr2.xlsx", f"{BASE}/_gt2.xlsx"]:
    try:
        os.remove(f)
    except:
        pass
