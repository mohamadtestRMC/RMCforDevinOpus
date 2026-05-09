"""Check key rows for unmatched component tracing."""
import sys, io, openpyxl, pandas as pd
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

wb = openpyxl.load_workbook('Template_Files/JT_With_copy.xlsx')
ws = wb.active

# Check formulas and balances for key rows
for r in [40, 41, 53, 54, 63, 64, 65]:
    wo = ws.cell(row=r, column=11).value
    proc = ws.cell(row=r, column=6).value
    # Film (PRINTING)
    ay = ws.cell(row=r, column=51).value  # Input Qty
    az = ws.cell(row=r, column=52).value  # Balance
    bb_mr = ws.cell(row=r, column=54).value  # Film MR#
    # Fresh1 (LAM)
    bw = ws.cell(row=r, column=75).value  # Fresh1 Qty
    bx = ws.cell(row=r, column=76).value  # Fresh1 Balance
    bz_mr = ws.cell(row=r, column=78).value  # Fresh1 MR#
    # Fresh2
    cg = ws.cell(row=r, column=85).value  # Fresh2 Qty
    ch = ws.cell(row=r, column=86).value  # Fresh2 Balance
    cj_mr = ws.cell(row=r, column=88).value  # Fresh2 MR#
    
    print(f"Row {r}: WO={wo}, Proc={proc}")
    if str(proc).upper() == 'PRINTING':
        print(f"  Film: Qty={ay}, Bal={az}, MR={bb_mr}")
    else:
        print(f"  Fresh1: Qty={bw}, Bal={bx}, MR={bz_mr}")
        print(f"  Fresh2: Qty={cg}, Bal={ch}, MR={cj_mr}")
    print()
wb.close()
