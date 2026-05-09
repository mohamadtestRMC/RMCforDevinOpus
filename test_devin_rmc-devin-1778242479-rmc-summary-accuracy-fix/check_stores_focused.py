"""
FOCUSED: Check Stores rows with EXACT Size+Mic+Material match,
then see if individual Issue Qty values match JT Input Qty.
"""
import shutil, os, openpyxl, pandas as pd
from engine.fill_jobtrack import COLS, DATA_START_ROW, _safe_str

BASE = "Template_Files"
shutil.copy2(f"{BASE}/Stores Recordings.xlsx", f"{BASE}/_st3.xlsx")
shutil.copy2(f"{BASE}/Jobtrack Feb With MRR.xlsx", f"{BASE}/_gt3.xlsx")

stores_df = pd.read_excel(f"{BASE}/_st3.xlsx", sheet_name=0, header=1)

# Identify columns
cols_map = {}
for c in stores_df.columns:
    cl = str(c).lower().strip()
    if 'sub' in cl and 'cat' in cl: cols_map['mat'] = c
    if 'mic' == cl or 'mic' in cl: cols_map['mic'] = c
    if 'width' in cl: cols_map['size'] = c
    if 'issue qty' in cl or ('issue' in cl and 'qty' in cl and 'date' not in cl): cols_map['qty'] = c
    if 'm.r.r' in cl and 'no' in cl: cols_map['mrr'] = c
    if 'issue' in cl and 'process' in cl: cols_map['proc'] = c
    if 'issue' in cl and 'wo' in cl: cols_map['wo'] = c
    if 'order' in cl and 'no' in cl and 'wo' not in cl: cols_map['order'] = c
    if 'supplier' in cl: cols_map['supplier'] = c

print(f"Columns: {cols_map}")

# Check if there's an Issue WO# or Order No# column  
wo_col = cols_map.get('wo')
order_col = cols_map.get('order')
print(f"WO col: {wo_col}, Order col: {order_col}")

# Show sample of WO column
if wo_col:
    samples = stores_df[wo_col].dropna().head(20)
    print(f"\nSample Issue WO# values: {list(samples)}")
if order_col:
    samples = stores_df[order_col].dropna().head(20)
    print(f"\nSample Order No# values: {list(samples)}")

wb_gt = openpyxl.load_workbook(f"{BASE}/_gt3.xlsx", data_only=True)
ws_gt = wb_gt.active

print("\n" + "=" * 90)
print("STORES ROWS MATCHING EXACT Size+Mic FOR EACH MISMATCHED ROW")
print("=" * 90)

for row in [42, 43, 45, 46]:
    uid = ws_gt.cell(row=row, column=1).value
    process = str(ws_gt.cell(row=row, column=COLS['Process']).value or '').strip().upper()
    order = str(ws_gt.cell(row=row, column=COLS['Order_No']).value or '').strip()
    
    if 'PRINT' in process:
        mat_name = _safe_str(ws_gt.cell(row=row, column=COLS['Input_Name']).value)
        mat_size = ws_gt.cell(row=row, column=COLS['Input_Size']).value
        mat_mic = ws_gt.cell(row=row, column=COLS['Input_Mic']).value
        jt_qty = float(ws_gt.cell(row=row, column=COLS['Input_Qty']).value or 0)
        jt_total = float(ws_gt.cell(row=row, column=COLS['Total_1st_Input']).value or 0)
        gt_mr = ws_gt.cell(row=row, column=COLS['Film_MR']).value
        gt_rate = float(ws_gt.cell(row=row, column=COLS['Film_Rate']).value or 0)
        slot = 'Film'
    else:
        mat_name = _safe_str(ws_gt.cell(row=row, column=COLS['Fresh1_Name']).value)
        mat_size = ws_gt.cell(row=row, column=COLS['Fresh1_Size']).value
        mat_mic = ws_gt.cell(row=row, column=COLS['Fresh1_Mic']).value
        jt_qty = float(ws_gt.cell(row=row, column=COLS['Fresh1_Qty']).value or 0)
        jt_total = float(ws_gt.cell(row=row, column=COLS['Total_Fresh1']).value or 0)
        gt_mr = ws_gt.cell(row=row, column=COLS['Fresh1_MR']).value
        gt_rate = float(ws_gt.cell(row=row, column=COLS['Fresh1_Rate']).value or 0)
        slot = 'Fresh1'
    
    print(f"\n{'='*80}")
    print(f"Row {row} | {uid} | {slot}={mat_name} | Size={mat_size} | Mic={mat_mic}")
    print(f"Order={order} | JT Input Qty={jt_qty} | JT Total={jt_total}")
    print(f"GT MR#={gt_mr} | GT Rate={gt_rate}")
    print(f"{'='*80}")
    
    # Filter stores by EXACT size + mic + material contains + process
    mask = pd.Series([True] * len(stores_df))
    
    # Size match
    if mat_size and cols_map.get('size'):
        mask &= pd.to_numeric(stores_df[cols_map['size']], errors='coerce') == float(mat_size)
    
    # Mic match
    if mat_mic and cols_map.get('mic'):
        mask &= pd.to_numeric(stores_df[cols_map['mic']], errors='coerce') == float(mat_mic)
    
    # Material match (contains PET but not MET PET)
    if cols_map.get('mat'):
        mat_upper = str(mat_name).strip().upper()
        if mat_upper == 'PET':
            # Exact PET — exclude MET PET
            def pet_match(x):
                xu = str(x).strip().upper()
                return xu == 'PET' or xu.startswith('PET ') or xu == 'PET CHEM' or xu == 'PET UPF'
            mask &= stores_df[cols_map['mat']].apply(pet_match)
    
    # Process
    if cols_map.get('proc'):
        proc_match = 'PRINT' if 'PRINT' in process else 'LAM'
        mask &= stores_df[cols_map['proc']].astype(str).str.upper().str.contains(proc_match, na=False)
    
    # Work order — try Issue WO# first, then Order No#
    wo_mask = pd.Series([False] * len(stores_df))
    if wo_col:
        wo_mask |= stores_df[wo_col].astype(str).str.strip().str.upper() == order.upper()
    if order_col:
        wo_mask |= stores_df[order_col].astype(str).str.strip().str.upper() == order.upper()
    mask &= wo_mask
    
    matched = stores_df[mask].copy()
    
    print(f"\n  Matching Stores rows (Size={mat_size}, Mic={mat_mic}, WO={order}): {len(matched)}")
    
    if len(matched) > 0:
        print(f"  {'MRR':>8} {'IssueQty':>10} {'Supplier':>15} {'Material':>15} {'Size':>6} {'Mic':>5}")
        print(f"  {'-'*65}")
        for _, sr in matched.iterrows():
            mrr = sr.get(cols_map['mrr'], '?')
            qty = sr.get(cols_map['qty'], '?')
            sup = sr.get(cols_map.get('supplier', ''), '?')
            mat = sr.get(cols_map['mat'], '?')
            sz = sr.get(cols_map['size'], '?')
            mc = sr.get(cols_map['mic'], '?')
            
            qty_f = float(qty) if pd.notna(qty) else 0
            flag = " ◄◄ = JT Qty!" if abs(qty_f - jt_qty) < 1 else ""
            print(f"  {str(mrr):>8} {str(qty):>10} {str(sup):>15} {str(mat):>15} {str(sz):>6} {str(mc):>5}{flag}")
        
        # Sum by MRR
        print(f"\n  Grouped by MRR:")
        if cols_map.get('mrr') and cols_map.get('qty'):
            grouped = matched.groupby(cols_map['mrr'])[cols_map['qty']].agg(['sum','count']).reset_index()
            for _, g in grouped.iterrows():
                mrr = g[cols_map['mrr']]
                total = g['sum']
                count = g['count']
                print(f"    MRR {mrr}: {count} rows, total qty = {total:.1f}")
            
            # Does the total of all matched rows = JT Qty?
            grand_total = matched[cols_map['qty']].sum()
            print(f"\n  Grand total all matched: {grand_total:.1f}")
            print(f"  JT Input Qty:           {jt_qty}")
            print(f"  JT Total:               {jt_total}")
            print(f"  Match Input Qty? {'✅ YES' if abs(grand_total - jt_qty) < 2 else '❌ NO'}")
            print(f"  Match Total?     {'✅ YES' if abs(grand_total - jt_total) < 2 else '❌ NO'}")
    else:
        # Try without size filter
        print("  No exact match with size. Trying without size filter...")
        mask2 = pd.Series([True] * len(stores_df))
        if mat_mic and cols_map.get('mic'):
            mask2 &= pd.to_numeric(stores_df[cols_map['mic']], errors='coerce') == float(mat_mic)
        if cols_map.get('mat'):
            mat_upper = str(mat_name).strip().upper()
            mask2 &= stores_df[cols_map['mat']].astype(str).str.upper().str.contains(mat_upper, na=False)
        if cols_map.get('proc'):
            proc_match = 'PRINT' if 'PRINT' in process else 'LAM'
            mask2 &= stores_df[cols_map['proc']].astype(str).str.upper().str.contains(proc_match, na=False)
        wo_mask2 = pd.Series([False] * len(stores_df))
        if wo_col:
            wo_mask2 |= stores_df[wo_col].astype(str).str.strip().str.upper() == order.upper()
        if order_col:
            wo_mask2 |= stores_df[order_col].astype(str).str.strip().str.upper() == order.upper()
        mask2 &= wo_mask2
        
        matched2 = stores_df[mask2]
        print(f"  Without size: {len(matched2)} rows")
        if len(matched2) > 0:
            for _, sr in matched2.head(20).iterrows():
                print(f"    MRR={sr.get(cols_map['mrr'])}, Qty={sr.get(cols_map['qty'])}, "
                      f"Size={sr.get(cols_map['size'])}, Mat={sr.get(cols_map['mat'])}")

wb_gt.close()
for f in [f"{BASE}/_st3.xlsx", f"{BASE}/_gt3.xlsx"]:
    try: os.remove(f)
    except: pass
