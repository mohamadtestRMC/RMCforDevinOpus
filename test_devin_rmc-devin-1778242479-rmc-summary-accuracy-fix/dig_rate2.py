"""What is 4.1951? Check all possible sources."""
import pandas as pd
import sys
sys.path.insert(0, '.')
from engine.rate_lookup import load_purchase_register, _find_col

pr = load_purchase_register("Template2/Purchase Register - 2021 - 2025 _Nov.xlsx")
cat_col = _find_col(pr, 'categery', 'category')
mat_col = _find_col(pr, 'material')

# Check PR for TPE entries
tpe_mask = pr[mat_col].astype(str).str.upper().str.strip().isin(['TPE', 'WPE', 'WLDPE'])
if cat_col:
    cat_mask = pr[cat_col].astype(str).str.lower().str.contains('film|raw', na=False)
    tpe_pr = pr[tpe_mask]
else:
    tpe_pr = pr[tpe_mask]

print(f"TPE/WPE/WLDPE in PR: {len(tpe_pr)} rows")
if len(tpe_pr) > 0:
    month_col = [c for c in pr.columns if str(c).strip().lower() == 'month']
    month_col = month_col[0] if month_col else None
    amt_col = [c for c in pr.columns if str(c).strip().lower() == 'amount']
    amt_col = amt_col[0] if amt_col else None
    qty_col = [c for c in pr.columns if str(c).strip().lower() == 'actual quantity']
    qty_col = qty_col[0] if qty_col else None
    rate_col = [c for c in pr.columns if str(c).strip().lower() == 'rate']
    rate_col = rate_col[0] if rate_col else None
    
    if month_col:
        for m in tpe_pr[month_col].astype(str).str.strip().unique():
            subset = tpe_pr[tpe_pr[month_col].astype(str).str.strip() == m]
            total_amt = pd.to_numeric(subset[amt_col], errors='coerce').fillna(0).sum()
            total_qty = pd.to_numeric(subset[qty_col], errors='coerce').fillna(0).sum()
            avg_rate = total_amt / total_qty if total_qty > 0 else 0
            print(f"  Month {m}: {len(subset)} rows, Amt={total_amt:.2f}, Qty={total_qty:.2f}, Rate={avg_rate:.4f}")
    
    print(f"\nSample TPE rows:")
    cols_to_show = [mat_col, cat_col, month_col, rate_col, amt_col, qty_col] 
    cols_to_show = [c for c in cols_to_show if c]
    print(tpe_pr[cols_to_show].head(20).to_string())

# Check the Granules Recipe file more thoroughly for N00694-related info
print("\n\n=== Granules Recipe deeper analysis ===")
df = pd.read_excel("Template2/Granules Recipe -Nov_2025.xlsx", sheet_name=0, header=None)

# Look at the row data to understand the formula
# Row 3 has 4.189 — let's see what row 3 is
print(f"\nRow 3 (has value close to 4.1951):")
r3 = df.iloc[3].dropna()
print(f"  Non-NaN values count: {len(r3)}")
# Show the last few values which might be rate/total
for idx in range(max(0, len(r3)-10), len(r3)):
    col = r3.index[idx]
    val = r3.iloc[idx]
    print(f"  Col {col}: {val}")

# Row 6 is headers — check last columns
print(f"\nRow 6 (headers) - last columns:")
r6 = df.iloc[6].dropna()
for idx in range(max(0, len(r6)-10), len(r6)):
    col = r6.index[idx]
    val = r6.iloc[idx]
    print(f"  Col {col}: {val}")

# What is the second-to-last and last column value for each WO# row?
print(f"\nFinal values per WO# row:")
for ridx in [7, 8, 9]:
    r = df.iloc[ridx].dropna()
    wo = df.iloc[ridx][1]
    last_vals = list(r.items())[-5:]
    print(f"  WO#{wo}: last values = {[(c, f'{v:.4f}' if isinstance(v, float) else v) for c, v in last_vals]}")
