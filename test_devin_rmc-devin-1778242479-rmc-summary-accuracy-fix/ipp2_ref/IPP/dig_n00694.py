"""Dig into N00694 TPE INH case — where does 4.1951 come from?"""
import pandas as pd
import sys
sys.path.insert(0, '.')
from engine.supplier_rates import load_granules_rates

# Check all granules files for N00694
for name, path in [
    ("Nov 2025", "Template2/Granules Recipe -Nov_2025.xlsx"),
    ("Feb 2026", "Template_Files/Granules Recipe - February 2026.xlsx"),
]:
    print(f"\n=== {name} ===")
    rates = load_granules_rates(path)
    print(f"  All entries: {dict(rates)}")
    for k, v in rates.items():
        if 'N00694' in k.upper() or 'N0069' in k.upper():
            print(f"  >>> FOUND: {k} = {v}")

# Maybe 4.1951 is a calculated rate?
# Let's check if N00694 is anywhere in the granules file raw
print("\n=== Raw Granules Recipe (Nov 2025) ===")
df = pd.read_excel("Template2/Granules Recipe -Nov_2025.xlsx", sheet_name=0, header=None)
print(f"Shape: {df.shape}")
print(f"First 5 rows:")
print(df.head(10).to_string())

# Search for N00694
for c in df.columns:
    mask = df[c].astype(str).str.upper().str.contains('N00694', na=False)
    if mask.any():
        print(f"\nN00694 found in column {c}:")
        print(df[mask].to_string())

# Maybe it's in another sheet?
xf = pd.ExcelFile("Template2/Granules Recipe -Nov_2025.xlsx")
print(f"\nSheets: {xf.sheet_names}")
for sheet in xf.sheet_names:
    df = pd.read_excel(xf, sheet_name=sheet, header=None)
    for c in df.columns:
        mask = df[c].astype(str).str.upper().str.contains('N00694', na=False)
        if mask.any():
            print(f"\nSheet '{sheet}', col {c}: N00694 found!")
            print(df[mask].to_string())
    # Also check for 4.1951
    for c in df.columns:
        mask = df[c].apply(lambda x: abs(float(x) - 4.1951) < 0.001 if isinstance(x, (int,float)) else False)
        if mask.any():
            print(f"\nSheet '{sheet}', col {c}: value 4.1951 found!")
            print(df[mask].to_string())
