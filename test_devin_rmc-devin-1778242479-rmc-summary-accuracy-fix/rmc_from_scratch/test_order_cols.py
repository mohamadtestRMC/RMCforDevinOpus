"""Check what's in the Order No columns of unfilled vs filled process sheets."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
from rmc_engine.data_reader import open_workbook, read_sheet_fast, safe_str

UNFILLED = Path(r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP\Files_need_to_study\Unfilled\1 Base RMC _ 2026 February.xlsx")
FILLED = Path(r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP\Files_need_to_study\Filled_Output\1 Base RMC _ 2026 February.xlsx")

sheets = [
    ("Print", 6, "Order  No"),
    ("Lam", 6, "Order No"),
    ("BFL", 6, "Order No"),
    ("Slit", 6, "Order No"),
]

for sn, hr, oc in sheets:
    print(f"\n{'='*60}")
    print(f"SHEET: {sn} (looking for '{oc}' column)")
    print(f"{'='*60}")

    for label, path in [("UNFILLED", UNFILLED), ("FILLED", FILLED)]:
        wb = open_workbook(path, data_only=True, read_only=True)
        h, rows = read_sheet_fast(wb, sn, hr, hr + 10)
        wb.close()

        oci = h.index(oc) if oc in h else -1
        if oci < 0:
            for i, hh in enumerate(h):
                if "order" in hh.lower():
                    oci = i
                    break

        print(f"\n  {label}: headers[0:8]={h[:8]}, order_col_idx={oci}")
        for ri, row in enumerate(rows[:5]):
            order_val = row[oci] if oci >= 0 and oci < len(row) else "N/A"
            # Show first few non-None values
            non_none = [(i, str(v)[:25]) for i, v in enumerate(row) if v is not None][:6]
            print(f"    Row {ri}: order='{order_val}' | sample: {non_none}")

    # Also check with data_only=False for formulas
    print(f"\n  UNFILLED FORMULAS:")
    wb = open_workbook(UNFILLED, data_only=False, read_only=True)
    h, rows = read_sheet_fast(wb, sn, hr, hr + 5)
    wb.close()
    oci = h.index(oc) if oc in h else -1
    if oci < 0:
        for i, hh in enumerate(h):
            if "order" in hh.lower():
                oci = i
                break
    for ri, row in enumerate(rows[:3]):
        order_val = row[oci] if oci >= 0 and oci < len(row) else "N/A"
        # Show first few cells
        sample = [(i, str(v)[:50]) for i, v in enumerate(row) if v is not None][:6]
        print(f"    Row {ri}: order_formula='{str(order_val)[:60]}' | sample: {sample}")
