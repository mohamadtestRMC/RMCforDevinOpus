"""
Test: Use process sheets from UNFILLED template + derive orders.
Compare against filled reference to measure accuracy.
"""
import sys, logging
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
logging.basicConfig(level=logging.INFO)

from rmc_engine.data_reader import (
    OrderIndex, open_workbook, read_sheet_fast, safe_float, safe_str,
)
from rmc_engine.rmc_compute import compute_rmc_summary, validate_rmc
from rmc_engine.process_builder import build_indexes_from_filled_reference

UNFILLED = Path(r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP\Files_need_to_study\Unfilled\1 Base RMC _ 2026 February.xlsx")
FILLED = Path(r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP\Files_need_to_study\Filled_Output\1 Base RMC _ 2026 February.xlsx")

# ---- STEP 1: Read process sheet indexes from UNFILLED template ----
print("=" * 60)
print("STEP 1: Read process sheets from UNFILLED template")
print("=" * 60)
wb = open_workbook(UNFILLED, data_only=True, read_only=True)

idx = {}
sheet_configs = [
    ("BFL",         6, 200, "Order No"),
    ("Print",       6, 400, "Order  No"),
    ("Lam",         6, 700, "Order No"),
    ("Slit",        6, 550, "Order No"),
    ("Bag&Pouch",   6, 100, "Order No"),
    ("Spout&Valve", 6, 20,  "Order No"),
    ("PTR Rew",     6, 200, "Order No"),
    ("HCI Rew",     6, 150, "Order No"),
    ("OPN_WIP",     5, 300, "W/O"),
    ("CLS_WIP",     5, 300, "W/O"),
]
for sn, hr, mr, oc in sheet_configs:
    oi = OrderIndex.from_sheet(wb, sn, hr, mr, oc)
    if oi:
        idx[sn] = oi
        print(f"  {sn}: {len(oi.all_rows)} rows, {len(oi.headers)} cols, {len(oi.orders())} orders")

if "FG" in wb.sheetnames:
    ws = wb["FG"]
    fg_rows_raw = list(ws.iter_rows(min_row=4, max_row=650, values_only=True))
    fg_headers = ["Row Labels", "Raw Output", "_c3", "_c4", "_c5", "HCI Wastage", "Final FG"]
    fg_data = [tuple(r[i] if i < len(r) else None for i in range(7)) for r in fg_rows_raw if r and r[0] is not None]
    idx["FG"] = OrderIndex(fg_headers, fg_data, 0)
    print(f"  FG: {len(fg_data)} rows")

wb.close()

# ---- STEP 2: Derive orders from process sheets ----
print("\n" + "=" * 60)
print("STEP 2: Derive order list from process sheets")
print("=" * 60)

all_orders = set()
for sn, oi in idx.items():
    all_orders.update(oi.orders())
print(f"  Unique orders from all sheets: {len(all_orders)}")

# Read order metadata from Jobtrack in template
wb = open_workbook(UNFILLED, data_only=True, read_only=True)
jt_headers, jt_rows = read_sheet_fast(wb, "Jobtrack", 4, 3000)
wb.close()
print(f"  Jobtrack rows: {len(jt_rows)}, headers: {len(jt_headers)}")

# Find key columns in Jobtrack
jt_cm = {h: i for i, h in enumerate(jt_headers)}
for cn in ["Order No", "Design Name", "Customer Name", "Sales Code", "Material", "Remarks", "Structure"]:
    if cn in jt_cm:
        print(f"  JT col '{cn}' at index {jt_cm[cn]}")
    else:
        matches = [h for h in jt_headers if cn.lower().replace(" ", "") in h.lower().replace(" ", "")]
        print(f"  JT col '{cn}' NOT found, similar: {matches[:3]}")

# ---- STEP 3: Get reference data for comparison ----
print("\n" + "=" * 60)
print("STEP 3: Read FILLED reference for validation")
print("=" * 60)

_, rmc_ref_rows, offsets_tuple = build_indexes_from_filled_reference(FILLED)
offsets, transfer_orders, other_film_orders, combined_orders = offsets_tuple
print(f"  Reference orders: {len(rmc_ref_rows)}")
print(f"  Orders with offsets: {len(offsets)}")
print(f"  Transfer orders: {transfer_orders}")
print(f"  Other film orders: {other_film_orders}")
print(f"  Combined orders: {combined_orders}")

# Show offset details
print("\n  Offset details:")
for order, ofs in sorted(offsets.items()):
    for key, val in ofs.items():
        print(f"    {order} | {key}: {val:+.4f}")

# ---- STEP 4a: Compute WITHOUT offsets using UNFILLED template indexes ----
print("\n" + "=" * 60)
print("STEP 4a: Compute WITHOUT offsets (zero offsets)")
print("=" * 60)

rmc_rows_no_ofs = compute_rmc_summary(
    rmc_ref_rows, idx,
    offsets={}, transfer_orders=transfer_orders,
    other_film_orders=other_film_orders, combined_orders=combined_orders,
)
val_no = validate_rmc(rmc_ref_rows, rmc_rows_no_ofs)
print(f"  Accuracy (no offsets): {val_no['accuracy_pct']}%")
print(f"  Exact: {val_no['exact_matches']} / {val_no['total_checks']}")
print(f"  Close (<1): {val_no['close_lt1']}")
print(f"  Mismatches (>1): {val_no['mismatches_gt1']}")
if val_no['top_mismatches']:
    print(f"\n  Top mismatches:")
    for m in val_no['top_mismatches'][:10]:
        print(f"    {m['order']:10s} | {m['col']:35s} | computed={m['computed']:12,.2f} | ref={m['reference']:12,.2f} | diff={m['diff']:+12,.2f}")

# ---- STEP 4b: Compute WITH offsets from filled reference ----
print("\n" + "=" * 60)
print("STEP 4b: Compute WITH offsets (from filled reference)")
print("=" * 60)

rmc_rows_with_ofs = compute_rmc_summary(
    rmc_ref_rows, idx,
    offsets=offsets, transfer_orders=transfer_orders,
    other_film_orders=other_film_orders, combined_orders=combined_orders,
)
val_with = validate_rmc(rmc_ref_rows, rmc_rows_with_ofs)
print(f"  Accuracy (with offsets): {val_with['accuracy_pct']}%")
print(f"  Exact: {val_with['exact_matches']} / {val_with['total_checks']}")
print(f"  Close (<1): {val_with['close_lt1']}")
print(f"  Mismatches (>1): {val_with['mismatches_gt1']}")
if val_with['top_mismatches']:
    print(f"\n  Top mismatches:")
    for m in val_with['top_mismatches'][:10]:
        print(f"    {m['order']:10s} | {m['col']:35s} | computed={m['computed']:12,.2f} | ref={m['reference']:12,.2f} | diff={m['diff']:+12,.2f}")

# ---- STEP 5: Check which orders from reference are NOT in process sheets ----
print("\n" + "=" * 60)
print("STEP 5: Order coverage check")
print("=" * 60)
ref_orders = {safe_str(r.get("order")) for r in rmc_ref_rows}
missing = ref_orders - all_orders
extra = all_orders - ref_orders
print(f"  Reference orders: {len(ref_orders)}")
print(f"  Template sheet orders: {len(all_orders)}")
print(f"  Missing from template: {len(missing)} → {sorted(missing)[:20]}")
print(f"  Extra in template: {len(extra)} → {sorted(extra)[:20]}")
