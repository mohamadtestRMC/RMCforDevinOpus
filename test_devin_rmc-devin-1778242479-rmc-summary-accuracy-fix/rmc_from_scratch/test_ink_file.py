"""Quick check of Ink Consumption file structure."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
from rmc_engine.data_reader import open_workbook, read_sheet_fast, safe_str, safe_float

ink = Path(r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP\Files_need_to_study\Unfilled\5 Ink Consumption February 2026.xlsx")
wb = open_workbook(ink, data_only=True, read_only=True)
print(f"Sheets: {wb.sheetnames}")

for sn in wb.sheetnames[:5]:
    print(f"\n--- Sheet: {sn} ---")
    for hr in [1, 2, 3, 4, 5]:
        h, rows = read_sheet_fast(wb, sn, hr, hr + 10)
        if h and any(hh for hh in h if hh != "_blank"):
            print(f"  header_row={hr}, {len(h)} cols, {len(rows)} rows")
            print(f"  Headers: {[hh for hh in h if hh != '_blank'][:15]}")
            for ri, r in enumerate(rows[:3]):
                vals = {h[i]: r[i] for i in range(min(len(h), len(r))) if r[i] is not None and h[i] != "_blank"}
                print(f"    Row {ri}: {dict(list(vals.items())[:8])}")
            break

# Also check the filled Print sheet for Ink Rate column
filled = Path(r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP\Files_need_to_study\Filled_Output\1 Base RMC _ 2026 February.xlsx")
wb2 = open_workbook(filled, data_only=True, read_only=True)
h, rows = read_sheet_fast(wb2, "Print", 6, 20)
ink_rate_ci = h.index("Ink Rate Per/Kg") if "Ink Rate Per/Kg" in h else -1
order_ci = h.index("Order  No") if "Order  No" in h else 1
dry_ink_ci = h.index("Dry Ink (Kgs)") if "Dry Ink (Kgs)" in h else -1
ink_val_ci = h.index("Ink Value") if "Ink Value" in h else -1

print(f"\n=== Print sheet Ink Rate analysis ===")
rates_seen = {}
for r in rows[:30]:
    o = safe_str(r[order_ci])
    ir = safe_float(r[ink_rate_ci]) if ink_rate_ci >= 0 else 0
    di = safe_float(r[dry_ink_ci]) if dry_ink_ci >= 0 else 0
    iv = safe_float(r[ink_val_ci]) if ink_val_ci >= 0 else 0
    if di > 0 and iv > 0:
        actual_rate = iv / di
        print(f"  {o}: DryInk={di:.2f}, InkRate={ir:.4f}, InkValue={iv:.2f}, calc_rate={actual_rate:.4f}")
        rates_seen[o] = actual_rate

# Check formulas for Ink Value
wb2.close()
wb3 = open_workbook(filled, data_only=False, read_only=True)
h, rows = read_sheet_fast(wb3, "Print", 6, 10)
print(f"\n=== Print Ink Value FORMULAS ===")
if ink_val_ci >= 0:
    print(f"  [{ink_val_ci}] {h[ink_val_ci]}")
    for ri, r in enumerate(rows[:3]):
        if ink_val_ci < len(r):
            print(f"    Row {ri}: {str(r[ink_val_ci])[:80]}")
if ink_rate_ci >= 0:
    print(f"  [{ink_rate_ci}] {h[ink_rate_ci]}")
    for ri, r in enumerate(rows[:3]):
        if ink_rate_ci < len(r):
            print(f"    Row {ri}: {str(r[ink_rate_ci])[:80]}")
wb3.close()
wb.close()
