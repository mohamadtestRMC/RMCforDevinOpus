"""
Analyze the exact column mapping between the Jobtrack and each process sheet.
Goal: find which Jobtrack columns correspond to SUMIF-used process sheet columns.
"""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
from rmc_engine.data_reader import open_workbook, read_sheet_fast, safe_float, safe_str

FILLED = Path(r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP\Files_need_to_study\Filled_Output\1 Base RMC _ 2026 February.xlsx")

wb = open_workbook(FILLED, data_only=True, read_only=True)

# Read Jobtrack headers
jt_h, jt_rows = read_sheet_fast(wb, "Jobtrack", 4, 3000)
print(f"Jobtrack: {len(jt_rows)} rows, {len(jt_h)} headers")

jt_cm = {h: i for i, h in enumerate(jt_h)}
proc_ci = jt_cm.get("Process", -1)
order_ci = jt_cm.get("Order No", -1)
print(f"  Process col: {proc_ci}, Order No col: {order_ci}")

# Group Jobtrack rows by Process
jt_by_proc = {}
for row in jt_rows:
    proc = safe_str(row[proc_ci]).strip() if proc_ci >= 0 else ""
    jt_by_proc.setdefault(proc, []).append(row)
print(f"\n  Processes: { {k: len(v) for k, v in jt_by_proc.items()} }")

# For each process sheet, list headers and the SUMIF columns used
sumif_cols = {
    "Print": ["Film Input (Kgs)", "Dry Ink (Kgs)", "Film Value", "Ink Value",
              "Wastage Qty (Calc)", "Wastage Value (AED)"],
    "Lam": ["Fresh Mat Qty", "Adh+Hard Solids Qty", "Fresh Mat Value",
            "Adh+Hard +Solv Val", "Wastage (Calc)", "Wastage (AED)",
            "Fresh Mat (Kgs)", "Lam Fresh Mat (Kgs)"],
    "Slit": ["Input (Kgs)", "Slitting Input Val (AED)", "Wastage (Kgs)", "Wastage Val (AED)"],
    "BFL": ["Wastage  (Kgs)", "Wastage  value (AED)"],
    "Bag&Pouch": ["PE STRIP + ZIPPER Qty", "PE STRIP + ZIPPER Value",
                  "Wastage (Calc)", "Wastage (AED)"],
    "Spout&Valve": ["TIN TIE+Valve+Spout Qty", "TIN TIE+Valve+Spout Value",
                    "Wastage (Calc)", "Wastage (AED)"],
    "HCI Rew": ["Sum of Wastage (Calc)", "Wastage Value (AED)"],
    "PTR Rew": ["Total Wastage Qty", "Total Wastage Value"],
}

sheet_configs = [
    ("BFL",         6, 200, "Order No"),
    ("Print",       6, 400, "Order  No"),
    ("Lam",         6, 700, "Order No"),
    ("Slit",        6, 550, "Order No"),
    ("Bag&Pouch",   6, 100, "Order No"),
    ("Spout&Valve", 6, 20,  "Order No"),
    ("PTR Rew",     6, 200, "Order No"),
    ("HCI Rew",     6, 150, "Order No"),
]

# Also check: for "Lam", read with extra columns
print("\n" + "="*80)
print("FULL HEADER LISTING FOR EACH PROCESS SHEET")
print("="*80)

for sn, hr, mr, oc in sheet_configs:
    h, rows = read_sheet_fast(wb, sn, hr, mr)
    # Find order col
    oci = h.index(oc) if oc in h else -1
    if oci < 0:
        for i, hh in enumerate(h):
            if "order" in hh.lower():
                oci = i
                break
    print(f"\n--- {sn} ({len(rows)} rows, {len(h)} cols) ---")
    for i, hh in enumerate(h):
        print(f"  [{i:2d}] {hh}")

    # Check which SUMIF columns are present
    used = sumif_cols.get(sn, [])
    print(f"\n  SUMIF columns check:")
    for col in used:
        if col in h:
            ci = h.index(col)
            # Get sample values for first 2 orders
            samples = []
            for row in rows[:3]:
                v = safe_float(row[ci]) if ci < len(row) else 0
                o = safe_str(row[oci]) if oci >= 0 and oci < len(row) else "?"
                samples.append(f"{o}={v:.2f}")
            print(f"    '{col}' -> col[{ci}] OK  samples: {', '.join(samples)}")
        else:
            partial = [hh for hh in h if col.lower().replace(" ", "") in hh.lower().replace(" ", "")]
            print(f"    '{col}' -> NOT FOUND  partial: {partial}")

# Now for each process sheet, find a sample order and trace its values back to Jobtrack
print("\n" + "="*80)
print("TRACE: Print B01065 vs Jobtrack B01065")
print("="*80)

# Get Print B01065
h_prn, rows_prn = read_sheet_fast(wb, "Print", 6, 400)
prn_oci = h_prn.index("Order  No") if "Order  No" in h_prn else 1
for row in rows_prn:
    if safe_str(row[prn_oci]) == "B01065":
        print("\nPrint B01065 values:")
        for i, (hh, v) in enumerate(zip(h_prn, row)):
            if v is not None:
                print(f"  [{i:2d}] {hh:35s} = {v}")
        break

# Get Jobtrack B01065 (first row)
for row in jt_rows:
    if safe_str(row[order_ci]) == "B01065":
        proc = safe_str(row[proc_ci])
        if "print" in proc.lower():
            print(f"\nJobtrack B01065 (Process={proc}) key columns:")
            for i, (hh, v) in enumerate(zip(jt_h, row)):
                if v is not None and i < 120:
                    print(f"  [{i:3d}] {hh:45s} = {str(v)[:50]}")
            break

# Also trace Lam
print("\n" + "="*80)
print("TRACE: Lam N00765 vs Jobtrack N00765")
print("="*80)

h_lam, rows_lam = read_sheet_fast(wb, "Lam", 6, 700)
lam_oci = h_lam.index("Order No") if "Order No" in h_lam else 1
for row in rows_lam:
    if safe_str(row[lam_oci]) == "N00765":
        print("\nLam N00765 first row values:")
        for i, (hh, v) in enumerate(zip(h_lam, row)):
            if v is not None:
                print(f"  [{i:2d}] {hh:40s} = {str(v)[:50]}")
        break

for row in jt_rows:
    if safe_str(row[order_ci]) == "N00765":
        proc = safe_str(row[proc_ci])
        if "lam" in proc.lower():
            print(f"\nJobtrack N00765 (Process={proc}) key columns:")
            for i, (hh, v) in enumerate(zip(jt_h, row)):
                if v is not None and i < 120:
                    print(f"  [{i:3d}] {hh:45s} = {str(v)[:50]}")
            break

wb.close()
