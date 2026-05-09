"""Focused check: where does 4.1951 come from for N00694?"""
import pandas as pd
import sys
sys.path.insert(0, '.')

# Read granules raw
df = pd.read_excel("Template2/Granules Recipe -Nov_2025.xlsx", sheet_name=0, header=None)

# Find N00694 in any cell
found = False
for c in df.columns:
    for idx, val in df[c].items():
        if 'N00694' in str(val).upper():
            print(f"Row {idx}, Col {c}: '{val}'")
            found = True
if not found:
    print("N00694 NOT found in Granules Recipe")

# Check what load_granules_rates does  
from engine.supplier_rates import load_granules_rates
rates = load_granules_rates("Template2/Granules Recipe -Nov_2025.xlsx")
print(f"\nAll Granules rates: {dict(rates)}")

# The rate 4.1951 - is it in the granules file somewhere?
print("\nSearching for values close to 4.1951...")
for c in df.columns:
    for idx, val in df[c].items():
        try:
            if abs(float(val) - 4.1951) < 0.01:
                print(f"  Row {idx}, Col {c}: {val}")
                # Print context - whole row
                row_data = df.iloc[idx].dropna().tolist()
                wo_found = [x for x in row_data if 'N00' in str(x).upper() or 'B00' in str(x).upper() or 'G00' in str(x).upper()]
                print(f"    WO#s in this row: {wo_found}")
        except: pass

# Maybe N00694 IS in the granules but load_granules_rates strips it?
# Let me check the last few columns where order/rate are
print(f"\nFile shape: {df.shape}")
# Check last columns of each row for order pattern
for idx in range(min(20, len(df))):
    row = df.iloc[idx]
    for c in df.columns:
        val = str(row[c]).strip().upper()
        if val and len(val) >= 5 and (val[0] in 'NBGCL') and val[1:4].isdigit():
            print(f"  Row {idx}: col {c} = '{val}'")
            
# Check the header to understand structure
print("\n=== Row 0 (headers?) ===")
for c in df.columns[:30]:
    print(f"  Col {c}: '{df.iloc[0][c]}'")
    
print("\n=== Row 1 ===")
for c in df.columns[:30]:
    print(f"  Col {c}: '{df.iloc[1][c]}'")
    
# How many rows have WO# data?
print("\n=== WO# column (col 1?) ===")
for idx in range(15):
    row = df.iloc[idx]
    print(f"  Row {idx}: col0='{row.iloc[0]}' col1='{row.iloc[1]}' col2='{row.iloc[2]}'")
