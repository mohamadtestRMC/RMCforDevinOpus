"""Analyze the How_to_link guide file to extract all linking rules."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
from rmc_engine.data_reader import open_workbook, safe_str

GUIDE = Path(r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP\Files_need_to_study\How_to_link\Base RMC Documents (2).xlsx")

# Read with data_only=True first
wb = open_workbook(GUIDE, data_only=True, read_only=True)
print(f"=== GUIDE FILE (data_only=True) ===")
print(f"Sheets: {wb.sheetnames}")

for sn in wb.sheetnames:
    ws = wb[sn]
    print(f"\n{'='*60}")
    print(f"SHEET: {sn}")
    print(f"{'='*60}")
    row_count = 0
    for row in ws.iter_rows(min_row=1, max_row=100, values_only=True):
        if row and any(c is not None for c in row):
            row_count += 1
            cells = [(i, str(v).strip()) for i, v in enumerate(row) if v is not None]
            if cells:
                print(f"  Row data: {cells}")
    print(f"  Total non-empty rows: {row_count}")

wb.close()

# Also read formulas
print("\n\n" + "="*60)
print("=== FORMULAS VIEW ===")
print("="*60)
wb2 = open_workbook(GUIDE, data_only=False, read_only=True)
for sn in wb2.sheetnames:
    ws = wb2[sn]
    print(f"\n--- SHEET: {sn} (formulas) ---")
    for row in ws.iter_rows(min_row=1, max_row=100, values_only=True):
        if row and any(c is not None for c in row):
            cells = [(i, str(v).strip()) for i, v in enumerate(row) if v is not None]
            if cells:
                for ci, cv in cells:
                    if cv.startswith("=") or len(cv) > 20:
                        print(f"  [{ci}] {cv}")
                    else:
                        print(f"  [{ci}] {cv}")
                print()
wb2.close()
