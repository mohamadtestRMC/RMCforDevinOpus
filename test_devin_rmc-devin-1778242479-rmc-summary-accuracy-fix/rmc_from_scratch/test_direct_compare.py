"""
Direct comparison: for each RMC column, compute from Jobtrack vs read from filled process sheet.
This reveals exactly which column mappings are wrong.
"""
import sys, logging
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
logging.basicConfig(level=logging.WARNING)

from rmc_engine.data_reader import (
    OrderIndex, open_workbook, read_sheet_fast, safe_float, safe_str,
)
from rmc_engine.rmc_compute import compute_rmc_summary, validate_rmc
from rmc_engine.process_builder import build_indexes_from_filled_reference

FILLED = Path(r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP\Files_need_to_study\Filled_Output\1 Base RMC _ 2026 February.xlsx")
UNFILLED = Path(r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP\Files_need_to_study\Unfilled\1 Base RMC _ 2026 February.xlsx")

# Step 1: Read filled process sheets (correct data)
print("Reading filled reference...")
idx_ref, rmc_ref_rows, offsets_tuple = build_indexes_from_filled_reference(FILLED)
offsets, transfers, other_film, combined = offsets_tuple

# Step 2: Read Jobtrack from filled file
wb = open_workbook(FILLED, data_only=True, read_only=True)
jt_h, jt_rows = read_sheet_fast(wb, "Jobtrack", 4, 3000)
wb.close()

jt_cm = {h: i for i, h in enumerate(jt_h)}
proc_ci = jt_cm.get("Process", -1)
order_ci = jt_cm.get("Order No", -1)

# Group JT by process
jt_by_proc = {}
for r in jt_rows:
    proc = safe_str(r[proc_ci]).strip()
    order = safe_str(r[order_ci]).strip()
    if proc and order:
        jt_by_proc.setdefault(proc, {}).setdefault(order, []).append(r)

# Step 3: Check which orders in the filled Lam sheet are NOT in JT LAM
lam_ref = idx_ref.get("Lam")
if lam_ref:
    lam_ref_orders = set(lam_ref.orders())
    jt_lam_orders = set(jt_by_proc.get("LAM", {}).keys())
    missing_lam = lam_ref_orders - jt_lam_orders
    extra_lam = jt_lam_orders - lam_ref_orders
    print(f"\nLam: {len(lam_ref_orders)} orders in filled sheet, {len(jt_lam_orders)} in JT LAM")
    print(f"  Missing from JT: {len(missing_lam)} orders")
    if missing_lam:
        print(f"  Examples: {sorted(missing_lam)[:10]}")
        # Check their values
        for o in sorted(missing_lam)[:5]:
            fmq = lam_ref.sumif(o, "Fresh Mat Qty")
            fmv = lam_ref.sumif(o, "Fresh Mat Value")
            print(f"    {o}: Fresh Mat Qty={fmq:.1f}, Value={fmv:.1f}")

# Same for Print
prn_ref = idx_ref.get("Print")
if prn_ref:
    prn_ref_orders = set(prn_ref.orders())
    jt_prn_orders = set(jt_by_proc.get("Printing", {}).keys())
    missing_prn = prn_ref_orders - jt_prn_orders
    print(f"\nPrint: {len(prn_ref_orders)} orders in filled sheet, {len(jt_prn_orders)} in JT Printing")
    print(f"  Missing from JT: {len(missing_prn)} orders")

# Step 4: For orders IN the JT, compare SUMIF values
print("\n" + "="*80)
print("COLUMN-BY-COLUMN SUMIF COMPARISON (filled process sheet vs JT computation)")
print("="*80)

test_orders = ["B01065", "L00327", "L00335", "N00765", "H01372", "J00877"]

for order in test_orders:
    print(f"\n--- {order} ---")
    
    # Print Film Input from filled sheet
    if prn_ref:
        ref_film = prn_ref.sumif(order, "Film Input (Kgs)")
        ref_dry_ink = prn_ref.sumif(order, "Dry Ink (Kgs)")
        ref_film_val = prn_ref.sumif(order, "Film Value")
        ref_ink_val = prn_ref.sumif(order, "Ink Value")
        ref_prn_waste_qty = prn_ref.sumif(order, "Wastage Qty (Calc)")
        ref_prn_waste_val = prn_ref.sumif(order, "Wastage Value (AED)")
    else:
        ref_film = ref_dry_ink = ref_film_val = ref_ink_val = ref_prn_waste_qty = ref_prn_waste_val = 0
    
    # Compute from JT Printing rows
    jt_film = jt_dry_ink = jt_film_val = jt_prn_waste = 0
    for r in jt_by_proc.get("Printing", {}).get(order, []):
        jt_film += safe_float(r[jt_cm.get("Total 1st Input", -1)] if jt_cm.get("Total 1st Input", -1) >= 0 else None)
        jt_dry_ink += safe_float(r[jt_cm.get("DRY INK QTY", -1)] if jt_cm.get("DRY INK QTY", -1) >= 0 else None)
        jt_film_val += safe_float(r[jt_cm.get("Film Value", -1)] if jt_cm.get("Film Value", -1) >= 0 else None)
        jt_prn_waste += safe_float(r[jt_cm.get("Wastage (Calc)", -1)] if jt_cm.get("Wastage (Calc)", -1) >= 0 else None)
    
    film_match = abs(ref_film - jt_film) < 1
    print(f"  Print Film Input: REF={ref_film:.2f}, JT={jt_film:.2f} {'OK' if film_match else 'MISMATCH'}")
    print(f"  Print Dry Ink:    REF={ref_dry_ink:.2f}, JT={jt_dry_ink:.2f}")
    print(f"  Print Film Value: REF={ref_film_val:.2f}, JT={jt_film_val:.2f}")
    print(f"  Print Ink Value:  REF={ref_ink_val:.2f} (JT: need Ink Rate)")
    print(f"  Print Waste Qty:  REF={ref_prn_waste_qty:.2f}, JT={jt_prn_waste:.2f}")
    
    # Lam Fresh Mat
    if lam_ref:
        ref_fresh_qty = lam_ref.sumif(order, "Fresh Mat Qty")
        ref_fresh_val = lam_ref.sumif(order, "Fresh Mat Value")
        ref_adh_hard_qty = lam_ref.sumif(order, "Adh+Hard Solids Qty")
        ref_adh_hard_val = lam_ref.sumif(order, "Adh+Hard +Solv Val")
        ref_lam_waste_qty = lam_ref.sumif(order, "Wastage (Calc)")
        ref_lam_waste_val = lam_ref.sumif(order, "Wastage (AED)")
    else:
        ref_fresh_qty = ref_fresh_val = ref_adh_hard_qty = ref_adh_hard_val = ref_lam_waste_qty = ref_lam_waste_val = 0
    
    # From JT LAM rows
    jt_fresh1 = jt_fresh2 = jt_fresh1_val = jt_fresh2_val = 0
    jt_adh_sol = jt_hard_sol = jt_adh_val = jt_hard_val = jt_sol_val = 0
    jt_ethyl_waste = jt_adh_hard_waste = jt_lam_waste = 0
    for r in jt_by_proc.get("LAM", {}).get(order, []):
        jt_fresh1 += safe_float(r[jt_cm["Total 1st Fresh Material Qty"]] if "Total 1st Fresh Material Qty" in jt_cm else None)
        jt_fresh2 += safe_float(r[jt_cm["Total 2nd Fresh Material Qty"]] if "Total 2nd Fresh Material Qty" in jt_cm else None)
        jt_fresh1_val += safe_float(r[jt_cm["1st Fresh Value"]] if "1st Fresh Value" in jt_cm else None)
        jt_fresh2_val += safe_float(r[jt_cm.get("2nd Fresh Value", -1)] if jt_cm.get("2nd Fresh Value", -1) >= 0 else None)
        jt_adh_sol += safe_float(r[jt_cm["Adh Solids"]] if "Adh Solids" in jt_cm else None)
        jt_hard_sol += safe_float(r[jt_cm["Hard Solids"]] if "Hard Solids" in jt_cm else None)
        jt_adh_val += safe_float(r[jt_cm["Adh Value"]] if "Adh Value" in jt_cm else None)
        jt_hard_val += safe_float(r[jt_cm["Hard Value"]] if "Hard Value" in jt_cm else None)
        jt_sol_val += safe_float(r[jt_cm["Sol Value"]] if "Sol Value" in jt_cm else None)
        jt_ethyl_waste += safe_float(r[jt_cm.get("Ethyl Wastage", -1)] if jt_cm.get("Ethyl Wastage", -1) >= 0 else None)
        jt_adh_hard_waste += safe_float(r[jt_cm.get("ADH + HARD Wastage", -1)] if jt_cm.get("ADH + HARD Wastage", -1) >= 0 else None)
        jt_lam_waste += safe_float(r[jt_cm["Wastage (Calc)"]] if "Wastage (Calc)" in jt_cm else None)
    
    jt_fresh_qty = jt_fresh1 + jt_fresh2
    jt_fresh_val = jt_fresh1_val + jt_fresh2_val
    jt_adh_hard_qty = jt_adh_sol + jt_hard_sol
    jt_adh_hard_simple = jt_adh_val + jt_hard_val + jt_sol_val
    
    # Check formula: Adh+Hard +Solv Val includes wastage?
    # From analysis: = Adh Value + Hard Value + Solv Value + ADH+HARD Wastage Value + Solv Wastage Value
    # But wastage values need rates. Let's check the component breakdown
    
    print(f"  Lam Fresh Qty:    REF={ref_fresh_qty:.2f}, JT={jt_fresh_qty:.2f}")
    print(f"  Lam Fresh Val:    REF={ref_fresh_val:.2f}, JT={jt_fresh_val:.2f}")
    print(f"  Lam Adh+H Qty:   REF={ref_adh_hard_qty:.2f}, JT={jt_adh_hard_qty:.2f}")
    print(f"  Lam Adh+H+S Val: REF={ref_adh_hard_val:.2f}, JT(simple)={jt_adh_hard_simple:.2f}")
    if abs(ref_adh_hard_val - jt_adh_hard_simple) > 1:
        # Try adding wastage values
        print(f"    Ethyl Waste: {jt_ethyl_waste:.2f}, ADH+H Waste: {jt_adh_hard_waste:.2f}")
    print(f"  Lam Waste Qty:    REF={ref_lam_waste_qty:.2f}, JT={jt_lam_waste:.2f}")

# Step 5: Check WIP comparison
print("\n" + "="*80)
print("WIP / FG COMPARISON: Template vs Filled")
print("="*80)

# Read from unfilled template
wb_u = open_workbook(UNFILLED, data_only=True, read_only=True)
opn_u = OrderIndex.from_sheet(wb_u, "OPN_WIP", 5, 300, "W/O")
cls_u = OrderIndex.from_sheet(wb_u, "CLS_WIP", 5, 300, "W/O")
wb_u.close()

opn_ref = idx_ref.get("OPN_WIP")
cls_ref = idx_ref.get("CLS_WIP")

# Compare for sample orders
for order in ["B01065", "N00765", "B00684"]:
    print(f"\n  OPN_WIP {order}:")
    if opn_ref and opn_u:
        rq = opn_ref.sumif(order, "Qty")
        rv = opn_ref.sumif(order, "Value")
        uq = opn_u.sumif(order, "Qty")
        uv = opn_u.sumif(order, "Value")
        print(f"    REF: Qty={rq:.2f}, Value={rv:.2f}")
        print(f"    TPL: Qty={uq:.2f}, Value={uv:.2f}")
    
    print(f"  CLS_WIP {order}:")
    if cls_ref and cls_u:
        rq = cls_ref.sumif(order, "Qty")
        rv = cls_ref.sumif(order, "Value")
        uq = cls_u.sumif(order, "Qty")
        uv = cls_u.sumif(order, "Value")
        print(f"    REF: Qty={rq:.2f}, Value={rv:.2f}")
        print(f"    TPL: Qty={uq:.2f}, Value={uv:.2f}")

# Count CLS_WIP #N/A values
cls_na_count = 0
if cls_u:
    for row in cls_u.all_rows:
        rate_ci = cls_u._ci("Rate")
        if rate_ci >= 0 and rate_ci < len(row):
            if str(row[rate_ci]) == "#N/A":
                cls_na_count += 1
print(f"\n  CLS_WIP #N/A rates in template: {cls_na_count} / {len(cls_u.all_rows) if cls_u else 0}")
