"""
Final debug: S110 hardener lookup returning wrong rate.
The issue is likely that 'S110' matches other materials via the fuzzy 'in' check.
"""
import pandas as pd
import sys
sys.path.insert(0, '.')
from engine.rate_lookup import load_purchase_register, _find_col

pr_df = load_purchase_register("Template2/Purchase Register - 2021 - 2025 _Nov.xlsx")
cat_col = _find_col(pr_df, 'categery', 'category')
mat_col = _find_col(pr_df, 'material')
month_col = None
for c in pr_df.columns:
    if str(c).strip().lower() == 'month':
        month_col = c
        break
amt_col = None
qty_col = None
for c in pr_df.columns:
    cl = str(c).strip().lower()
    if cl == 'amount': amt_col = c
    if cl == 'actual quantity': qty_col = c

# Check what materials 'S110' matches via fuzzy logic
print("Materials matching 'S110 in x or x in S110':")
adh_mask = pr_df[cat_col].astype(str).str.lower().str.contains('adhesive', na=False)
for _, r in pr_df[adh_mask].iterrows():
    mat = str(r[mat_col]).strip().upper()
    if 'S110' in mat or mat in 'S110':
        m = str(r[month_col]).strip()
        # Only show relevant months
        if '2025' in m or '2024' in m:
            print(f"  Material='{mat}', Month={m}, Rate={r.get('Rate', '?')}, "
                  f"Qty={r.get('Actual Quantity', '?')}, Amt={r.get('Amount', '?')}")

# What's the issue? Let's check if S110 also matches something else
print("\n\nAll ADHESIVE materials containing 'S110':")
all_mats = pr_df[adh_mask][mat_col].astype(str).str.strip().unique()
for mat in sorted(all_mats):
    if 'S110' in mat.upper() or mat.upper() in 'S110':
        count = len(pr_df[adh_mask & (pr_df[mat_col].astype(str).str.strip() == mat)])
        print(f"  '{mat}' ({count} rows)")

# Now compute with EXACT match only
print("\n\nS110 EXACT match in Oct 2025 (most recent month):")
s110_exact = pr_df[adh_mask & (pr_df[mat_col].astype(str).str.upper().str.strip() == 'S110')]
s110_oct = s110_exact[s110_exact[month_col].astype(str).str.strip() == '10-2025']
if len(s110_oct) > 0:
    total_a = pd.to_numeric(s110_oct[amt_col], errors='coerce').fillna(0).sum()
    total_q = pd.to_numeric(s110_oct[qty_col], errors='coerce').fillna(0).sum()
    r = total_a / total_q if total_q > 0 else 0
    print(f"  Rate = {r:.4f} (GT target: 10.4766)")
    for _, row in s110_oct.iterrows():
        print(f"    Rate={row.get('Rate', '?')}, Qty={row.get('Actual Quantity', '?')}, Amt={row.get('Amount', '?')}")
else:
    print("  NO DATA")

# Now check what the FUZZY match includes
print("\n\nAll rows matching fuzzy 'S110' in Oct 2025:")
for _, r in pr_df[adh_mask].iterrows():
    mat = str(r[mat_col]).strip().upper()
    m = str(r[month_col]).strip()
    if ('S110' in mat or mat in 'S110') and m == '10-2025':
        print(f"  Material='{mat}', Rate={r.get('Rate', '?')}, Qty={r.get('Actual Quantity', '?')}, Amt={r.get('Amount', '?')}")

# Check all months where S110 or related materials appear near the report month
print("\n\nChecking what 'S110 in x' catches vs 'x == S110':")
fuzzy_months = {}
exact_months = {}
for _, r in pr_df[adh_mask].iterrows():
    mat = str(r[mat_col]).strip().upper()
    m = str(r[month_col]).strip()
    
    if mat == 'S110':
        if m not in exact_months:
            exact_months[m] = {'amt': 0, 'qty': 0}
        exact_months[m]['amt'] += float(r.get(amt_col, 0) or 0)
        exact_months[m]['qty'] += float(r.get('Actual Quantity', 0) or 0)
    
    if 'S110' in mat or mat in 'S110':
        if m not in fuzzy_months:
            fuzzy_months[m] = {'amt': 0, 'qty': 0, 'materials': set()}
        fuzzy_months[m]['amt'] += float(r.get(amt_col, 0) or 0)
        fuzzy_months[m]['qty'] += float(r.get('Actual Quantity', 0) or 0)
        fuzzy_months[m]['materials'].add(mat)

# Compare for months near 10-2025
for m in sorted(fuzzy_months.keys()):
    if '2025' in m or '2024' in m:
        fz = fuzzy_months[m]
        ex = exact_months.get(m, {'amt': 0, 'qty': 0})
        fz_rate = fz['amt'] / fz['qty'] if fz['qty'] > 0 else 0
        ex_rate = ex['amt'] / ex['qty'] if ex['qty'] > 0 else 0
        diff = "SAME" if abs(fz_rate - ex_rate) < 0.001 else "DIFFERENT"
        print(f"  {m}: Fuzzy={fz_rate:.4f} (mats: {fz['materials']}), Exact={ex_rate:.4f} [{diff}]")

print("\n\nCONCLUSION: The fix is to use EXACT material matching, not fuzzy 'in' matching.")
print("The fuzzy match 'S110 in x or x in S110' incorrectly includes other materials.")
