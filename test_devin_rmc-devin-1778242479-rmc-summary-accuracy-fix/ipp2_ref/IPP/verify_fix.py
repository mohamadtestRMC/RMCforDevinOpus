"""Verify the plain-number qty matching fix for row 2350."""
import sys, os
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.path.insert(0, r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP")

from engine.mrr_lookup import load_stores_recordings, match_formula_qtys_to_store
from engine.fill_jobtrack import _pick_mrrs_by_total_qty

BASE = r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP\Template3"
stores_df = load_stores_recordings(os.path.join(BASE, "Stores Recordings.xlsx"))

print("=" * 60)
print("TEST 1: Row 2350 - MET PET, BW=100, WO=N01067")
print("=" * 60)

# Simulate the fix: plain number 100 treated as [100.0]
formula_qtys = [100.0]
mrr_qty = match_formula_qtys_to_store(
    stores_df, formula_qtys, 'MET PET', 12, 'N01067', 'LAMINATION'
)
print(f"match_formula_qtys_to_store([100.0], MET PET, mic=12, WO=N01067, LAMINATION)")
print(f"  Result: {mrr_qty}")
if mrr_qty:
    mrr_list = _pick_mrrs_by_total_qty(mrr_qty, 100.0)
    print(f"  _pick_mrrs_by_total_qty -> {mrr_list}")
    expected = [81732]
    if mrr_list == expected:
        print(f"  PASS: Correctly returns only MRR 81732")
    else:
        print(f"  FAIL: Expected {expected}, got {mrr_list}")
else:
    print("  match_formula_qtys_to_store returned empty - trying fallback")
    from engine.mrr_lookup import lookup_mrr_with_qty
    mrr_qty_full = lookup_mrr_with_qty(stores_df, 'MET PET', 12, 773, 'N01067', 'LAMINATION')
    print(f"  Fallback lookup_mrr_with_qty: {mrr_qty_full}")
    mrr_list = _pick_mrrs_by_total_qty(mrr_qty_full, 100.0)
    print(f"  _pick_mrrs_by_total_qty -> {mrr_list}")

print()
print("=" * 60)
print("TEST 2: Row 625 - PET, AY=107+500.1, WO=B01077")
print("=" * 60)
formula_qtys_625 = [107.0, 500.1]
mrr_qty_625 = match_formula_qtys_to_store(
    stores_df, formula_qtys_625, 'PET', 12, 'B01077', 'PRINTING'
)
print(f"match_formula_qtys_to_store([107, 500.1], PET, mic=12, WO=B01077, PRINTING)")
print(f"  Result: {mrr_qty_625}")
if mrr_qty_625:
    mrr_list_625 = _pick_mrrs_by_total_qty(mrr_qty_625, 242.1)
    print(f"  _pick_mrrs_by_total_qty(target=242.1) -> {mrr_list_625}")
    print(f"  Expected: [84526, 85460] (both contribute)")

print()
print("=" * 60)
print("TEST 3: Row 747 - PET, AY=365, WO=B01077")
print("=" * 60)
# With fix: plain number 365 treated as [365.0]
formula_qtys_747 = [365.0]
mrr_qty_747 = match_formula_qtys_to_store(
    stores_df, formula_qtys_747, 'PET', 12, 'B01077', 'PRINTING'
)
print(f"match_formula_qtys_to_store([365.0], PET, mic=12, WO=B01077, PRINTING)")
print(f"  Result: {mrr_qty_747}")
if not mrr_qty_747:
    print("  No match (365 is a balance, not a direct Store qty)")
    print("  This is expected - balance tracing should handle it")

print("\nDONE!")
