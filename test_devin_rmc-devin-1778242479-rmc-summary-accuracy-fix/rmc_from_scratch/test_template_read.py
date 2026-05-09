"""Test: can we read process sheet data from the UNFILLED template?"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
from rmc_engine.data_reader import open_workbook, read_sheet_fast, safe_float, safe_str
from rmc_engine.process_builder import build_indexes_from_filled_reference
from rmc_engine.rmc_compute import compute_rmc_summary, validate_rmc

UNFILLED = Path(r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP\Files_need_to_study\Unfilled\1 Base RMC _ 2026 February.xlsx")
FILLED = Path(r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP\Files_need_to_study\Filled_Output\1 Base RMC _ 2026 February.xlsx")

print("Reading UNFILLED template as if it were a filled reference...")
try:
    idx_u, rmc_ref_u, offsets_u = build_indexes_from_filled_reference(UNFILLED)
    offsets, transfers, other_film, combined = offsets_u
    print(f"Indexes: {list(idx_u.keys())}")
    print(f"Orders: {len(rmc_ref_u)}")
    print(f"Offsets: {len(offsets)}, Transfers: {len(transfers)}, Other Film: {len(other_film)}, Combined: {len(combined)}")

    rmc_rows = compute_rmc_summary(rmc_ref_u, idx_u, offsets, transfers, other_film, combined)
    print(f"Computed {len(rmc_rows)} rows")

    # Validate against filled reference
    print("\nReading FILLED reference for validation...")
    _, rmc_ref_filled, _ = build_indexes_from_filled_reference(FILLED)
    val = validate_rmc(rmc_ref_filled, rmc_rows)

    print(f"\nAccuracy: {val['accuracy_pct']}%")
    print(f"Exact: {val['exact_matches']} / {val['total_checks']}")
    print(f"Close (<1): {val['close_lt1']}")
    print(f"Mismatches (>1): {val['mismatches_gt1']}")

    if val['top_mismatches']:
        print(f"\nTop 5 mismatches:")
        for m in val['top_mismatches'][:5]:
            print(f"  {m['order']:10s} | {m['col']:35s} | diff={m['diff']:12,.2f}")
    else:
        print("\nPERFECT MATCH!")

except Exception as e:
    import traceback
    traceback.print_exc()
