"""Check WIP and FG data in unfilled vs filled templates."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
from rmc_engine.data_reader import open_workbook, read_sheet_fast, safe_float, safe_str

UNFILLED = Path(r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP\Files_need_to_study\Unfilled\1 Base RMC _ 2026 February.xlsx")
FILLED = Path(r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP\Files_need_to_study\Filled_Output\1 Base RMC _ 2026 February.xlsx")

for sn, hr in [("OPN_WIP", 5), ("CLS_WIP", 5)]:
    print(f"\n{'='*60}")
    print(f"SHEET: {sn}")
    for label, path in [("UNFILLED", UNFILLED), ("FILLED", FILLED)]:
        wb = open_workbook(path, data_only=True, read_only=True)
        h, rows = read_sheet_fast(wb, sn, hr, 310)
        wb.close()
        print(f"\n  {label}: {len(h)} cols, {len(rows)} rows")
        print(f"  Headers: {h}")
        # Show first 3 data rows
        for ri, row in enumerate(rows[:3]):
            vals = {h[i]: row[i] for i in range(min(len(h), len(row))) if row[i] is not None}
            print(f"    Row {ri}: {vals}")

# Check if values match
print(f"\n{'='*60}")
print("VALUE COMPARISON: CLS_WIP B01065")
for label, path in [("UNFILLED", UNFILLED), ("FILLED", FILLED)]:
    wb = open_workbook(path, data_only=True, read_only=True)
    h, rows = read_sheet_fast(wb, "CLS_WIP", 5, 310)
    wb.close()
    wo_ci = h.index("W/O") if "W/O" in h else 0
    for row in rows:
        if safe_str(row[wo_ci]) == "B01065":
            print(f"\n  {label} CLS_WIP B01065:")
            for i, (hh, v) in enumerate(zip(h, row)):
                print(f"    [{i}] {hh}: {v}")
            break
    else:
        print(f"\n  {label}: B01065 not found in CLS_WIP")

# Also compare FG
print(f"\n{'='*60}")
print("FG sheet comparison")
for label, path in [("UNFILLED", UNFILLED), ("FILLED", FILLED)]:
    wb = open_workbook(path, data_only=True, read_only=True)
    ws = wb["FG"]
    rows = list(ws.iter_rows(min_row=3, max_row=8, values_only=True))
    wb.close()
    print(f"\n  {label} FG rows 3-8:")
    for ri, row in enumerate(rows):
        non_none = [(i, str(v)[:25]) for i, v in enumerate(row) if v is not None][:8]
        print(f"    Row {3+ri}: {non_none}")
