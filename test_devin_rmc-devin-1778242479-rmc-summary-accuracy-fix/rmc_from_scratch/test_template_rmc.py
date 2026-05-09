"""Check what's in the unfilled template's RMC summary."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
from rmc_engine.data_reader import open_workbook, safe_str, safe_float

UNFILLED = Path(r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP\Files_need_to_study\Unfilled\1 Base RMC _ 2026 February.xlsx")

wb = open_workbook(UNFILLED, data_only=True, read_only=True)
ws = wb["RMC summary"]

# Check rows 1-10
print("=== RMC summary rows 1-10 ===")
for row in ws.iter_rows(min_row=1, max_row=10, values_only=True):
    non_none = [(i, str(v)[:40]) for i, v in enumerate(row) if v is not None]
    if non_none:
        print(f"  {non_none[:8]}")

# Check rows around the header (row 6) and first data rows
print("\n=== Rows 5-12 (header + first data) ===")
for i, row in enumerate(ws.iter_rows(min_row=5, max_row=12, values_only=True)):
    non_none = [(j, str(v)[:35]) for j, v in enumerate(row) if v is not None]
    print(f"  Row {5+i}: {non_none[:10]}")

# Also check with formulas
wb.close()
wb2 = open_workbook(UNFILLED, data_only=False, read_only=True)
ws2 = wb2["RMC summary"]
print("\n=== RMC summary FORMULAS rows 7-12 ===")
for i, row in enumerate(ws2.iter_rows(min_row=7, max_row=12, values_only=True)):
    non_none = [(j, str(v)[:50]) for j, v in enumerate(row) if v is not None]
    if non_none:
        print(f"  Row {7+i}: {non_none[:8]}")

# Check how many non-empty rows total
count = 0
for row in ws2.iter_rows(min_row=7, max_row=700, values_only=True):
    if row and row[1] is not None:
        count += 1
print(f"\n=== Total RMC summary rows with Order No (col B): {count} ===")

# Check col A too
count_a = 0
for row in ws2.iter_rows(min_row=7, max_row=700, values_only=True):
    if row and row[0] is not None:
        count_a += 1
print(f"Total with col A: {count_a}")

wb2.close()

# Now check: what does the FILLED version's process sheets look like vs unfilled?
FILLED = Path(r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP\Files_need_to_study\Filled_Output\1 Base RMC _ 2026 February.xlsx")
print("\n=== Comparing Print sheet row counts ===")
for label, path in [("UNFILLED", UNFILLED), ("FILLED", FILLED)]:
    from rmc_engine.data_reader import read_sheet_fast
    wb = open_workbook(path, data_only=True, read_only=True)
    h, rows = read_sheet_fast(wb, "Print", 6, 400)
    wb.close()
    print(f"  {label} Print: {len(rows)} rows, {len(h)} cols")

for label, path in [("UNFILLED", UNFILLED), ("FILLED", FILLED)]:
    wb = open_workbook(path, data_only=True, read_only=True)
    h, rows = read_sheet_fast(wb, "Lam", 6, 700)
    wb.close()
    print(f"  {label} Lam: {len(rows)} rows, {len(h)} cols")
