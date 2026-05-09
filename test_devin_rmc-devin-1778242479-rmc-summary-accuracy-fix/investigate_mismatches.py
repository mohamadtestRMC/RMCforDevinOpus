"""
Deep investigation of the 3 remaining mismatches in Nov 2025:
1. S621 adhesive: GT=11.6529 vs ENG=12.9339
2. S110 hardener: GT=10.4766 vs ENG=12.4808
3. CR 88-300 hardener: GT=9.5208 vs ENG=10.4477

The engine is returning rates that don't match. Let's find out exactly
which PR rows are being used and which SHOULD be used.
"""
import pandas as pd
import sys
sys.path.insert(0, '.')
from engine.rate_lookup import load_purchase_register, _find_col, _filter_by_month, _qty_weighted_rate

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
    if cl == 'amount':
        amt_col = c
    if cl == 'actual quantity':
        qty_col = c

rate_col = [c for c in pr_df.columns if str(c).strip().lower() == 'rate']
rate_col = rate_col[0] if rate_col else 'Rate'

report_month = '11-2025'

print("=" * 100)
print("DEEP INVESTIGATION: S621, S110, CR 88-300 Rate Mismatches")
print("=" * 100)

# ── 1. S621 (Adhesive) ──
print("\n" + "-" * 80)
print("1. S621 ADHESIVE | GT=11.6529, Engine=12.9339")
print("-" * 80)

# Show ALL S621 rows in PR
adh_mask = pr_df[cat_col].astype(str).str.lower().str.contains('adhesive', na=False)
s621_mask = pr_df[mat_col].astype(str).str.upper().str.strip() == 'S621'
s621_rows = pr_df[adh_mask & s621_mask]
print(f"\nAll S621 ADHESIVE rows in PR: {len(s621_rows)}")
for _, r in s621_rows.iterrows():
    print(f"  Month={r[month_col]}, Rate={r[rate_col]}, Qty={r[qty_col]}, Amt={r[amt_col]}, Tracking={r.get('Tracking N o.', '?')}")

# Filter by Nov 2025
s621_nov = s621_rows[s621_rows[month_col].astype(str).str.strip() == report_month]
print(f"\nS621 in Nov 2025: {len(s621_nov)} rows")
for _, r in s621_nov.iterrows():
    print(f"  Rate={r[rate_col]}, Qty={r[qty_col]}, Amt={r[amt_col]}")

if len(s621_nov) > 0:
    total_amt = pd.to_numeric(s621_nov[amt_col], errors='coerce').fillna(0).sum()
    total_qty = pd.to_numeric(s621_nov[qty_col], errors='coerce').fillna(0).sum()
    calc_rate = total_amt / total_qty if total_qty > 0 else 0
    print(f"  Nov 2025 weighted rate: {total_amt}/{total_qty} = {calc_rate:.4f}")

# What about S620? Similar name
s620_mask = pr_df[mat_col].astype(str).str.upper().str.strip() == 'S620'
s620_rows = pr_df[adh_mask & s620_mask]
print(f"\nAll S620 ADHESIVE rows in PR: {len(s620_rows)}")
s620_nov = s620_rows[s620_rows[month_col].astype(str).str.strip() == report_month]
print(f"S620 in Nov 2025: {len(s620_nov)} rows")
for _, r in s620_nov.iterrows():
    print(f"  Rate={r[rate_col]}, Qty={r[qty_col]}, Amt={r[amt_col]}")

# Check: maybe engine is matching both S620 and S621?
print("\nEngine filter test: what does 'S621 in x or x in S621' match?")
for _, r in pr_df[adh_mask].iterrows():
    mat = str(r[mat_col]).strip().upper()
    if 'S621' in mat or mat in 'S621':
        if str(r[month_col]).strip() == report_month:
            print(f"  MATCHED: Material='{mat}', Month={r[month_col]}, Rate={r[rate_col]}, Qty={r[qty_col]}, Amt={r[amt_col]}")

# The GT rate is 11.6529 - can we reverse-engineer what rows produce it?
print(f"\nReverse-engineering GT rate 11.6529:")
print(f"  If it's Amt/Qty, we need rows where total_amt/total_qty = 11.6529")

# ── 2. S110 (Hardener for S621) ──
print("\n" + "-" * 80)
print("2. S110 HARDENER | GT=10.4766, Engine=12.4808")
print("-" * 80)

s110_mask = pr_df[mat_col].astype(str).str.upper().str.strip() == 'S110'
s110_rows = pr_df[adh_mask & s110_mask]
print(f"\nAll S110 ADHESIVE rows in PR: {len(s110_rows)}")
for _, r in s110_rows.iterrows():
    print(f"  Month={r[month_col]}, Rate={r[rate_col]}, Qty={r[qty_col]}, Amt={r[amt_col]}")

s110_nov = s110_rows[s110_rows[month_col].astype(str).str.strip() == report_month]
print(f"\nS110 in Nov 2025: {len(s110_nov)} rows")
for _, r in s110_nov.iterrows():
    print(f"  Rate={r[rate_col]}, Qty={r[qty_col]}, Amt={r[amt_col]}")

if len(s110_nov) > 0:
    total_amt = pd.to_numeric(s110_nov[amt_col], errors='coerce').fillna(0).sum()
    total_qty = pd.to_numeric(s110_nov[qty_col], errors='coerce').fillna(0).sum()
    calc_rate = total_amt / total_qty if total_qty > 0 else 0
    print(f"  Nov 2025 weighted rate: {total_amt}/{total_qty} = {calc_rate:.4f}")

# ── 3. CR 88-300 (Hardener for 75-300) ──
print("\n" + "-" * 80)
print("3. CR 88-300 HARDENER | GT=9.5208, Engine=10.4477")
print("-" * 80)

cr88_mask = pr_df[mat_col].astype(str).str.upper().str.strip() == 'CR 88-300'
cr88_rows = pr_df[adh_mask & cr88_mask]
print(f"\nAll CR 88-300 ADHESIVE rows in PR: {len(cr88_rows)}")
for _, r in cr88_rows.iterrows():
    print(f"  Month={r[month_col]}, Rate={r[rate_col]}, Qty={r[qty_col]}, Amt={r[amt_col]}")

cr88_nov = cr88_rows[cr88_rows[month_col].astype(str).str.strip() == report_month]
print(f"\nCR 88-300 in Nov 2025: {len(cr88_nov)} rows")
if len(cr88_nov) == 0:
    print("  >> NO ROWS for Nov 2025! That's why rate is from other months (all-time avg)")
    print("  >> GT rate 9.5208 must come from somewhere else...")

# Check CR 800-300
cr800_mask = pr_df[mat_col].astype(str).str.upper().str.strip() == 'CR 800-300'
cr800_rows = pr_df[adh_mask & cr800_mask]
print(f"\nAll CR 800-300 ADHESIVE rows in PR: {len(cr800_rows)}")
cr800_nov = cr800_rows[cr800_rows[month_col].astype(str).str.strip() == report_month]
print(f"CR 800-300 in Nov 2025: {len(cr800_nov)} rows")
for _, r in cr800_nov.iterrows():
    print(f"  Rate={r[rate_col]}, Qty={r[qty_col]}, Amt={r[amt_col]}")

# ── 4. Check what EXACT month the GT rates correspond to ──
print("\n" + "-" * 80)
print("4. MONTH-BY-MONTH RATE CHECK for problematic materials")
print("-" * 80)

for mat_name in ['S621', 'S110', 'CR 88-300', 'CR 800-300']:
    mat_mask2 = pr_df[mat_col].astype(str).str.upper().str.strip() == mat_name
    rows = pr_df[adh_mask & mat_mask2]
    if len(rows) == 0:
        continue
    
    print(f"\n  {mat_name} rates by month:")
    months = rows[month_col].astype(str).str.strip().unique()
    for m in sorted(months):
        m_rows = rows[rows[month_col].astype(str).str.strip() == m]
        total_a = pd.to_numeric(m_rows[amt_col], errors='coerce').fillna(0).sum()
        total_q = pd.to_numeric(m_rows[qty_col], errors='coerce').fillna(0).sum()
        r = total_a / total_q if total_q > 0 else 0
        mark = " <<< GT target" if (
            (mat_name == 'S621' and abs(r - 11.6529) < 0.001) or
            (mat_name == 'S110' and abs(r - 10.4766) < 0.001) or
            (mat_name == 'CR 88-300' and abs(r - 9.5208) < 0.001)
        ) else ""
        print(f"    {m:>10}: Rate={r:.4f} (Amt={total_a:.2f}, Qty={total_q:.2f}){mark}")

# ── 5. Check if S200 exists (maybe S621 is mislabeled?) ──
print("\n" + "-" * 80)
print("5. CHECKING ALL 'S' prefix materials in ADHESIVE category, Nov 2025")
print("-" * 80)

s_mask = pr_df[mat_col].astype(str).str.upper().str.strip().str.startswith('S')
s_adh = pr_df[adh_mask & s_mask]
s_adh_nov = s_adh[s_adh[month_col].astype(str).str.strip() == report_month]
print(f"All 'S' materials in ADHESIVE, Nov 2025:")
for mat in sorted(s_adh_nov[mat_col].astype(str).str.strip().unique()):
    mat_rows = s_adh_nov[s_adh_nov[mat_col].astype(str).str.strip() == mat]
    total_a = pd.to_numeric(mat_rows[amt_col], errors='coerce').fillna(0).sum()
    total_q = pd.to_numeric(mat_rows[qty_col], errors='coerce').fillna(0).sum()
    r = total_a / total_q if total_q > 0 else 0
    print(f"  {mat}: Rate={r:.4f} (Amt={total_a:.2f}, Qty={total_q:.2f})")

# ── 6. Check if GT uses "10-2025" (October) for S621/S110/CR88 ──
print("\n" + "-" * 80)
print("6. CHECKING ADJACENT MONTHS (Oct, Nov, Sep 2025)")
print("-" * 80)

for mat_name in ['S621', 'S110', 'CR 88-300']:
    mat_mask3 = pr_df[mat_col].astype(str).str.upper().str.strip() == mat_name
    rows = pr_df[adh_mask & mat_mask3]
    
    for m in ['9-2025', '10-2025', '11-2025', '12-2025']:
        m_rows = rows[rows[month_col].astype(str).str.strip() == m]
        if len(m_rows) > 0:
            total_a = pd.to_numeric(m_rows[amt_col], errors='coerce').fillna(0).sum()
            total_q = pd.to_numeric(m_rows[qty_col], errors='coerce').fillna(0).sum()
            r = total_a / total_q if total_q > 0 else 0
            gt_match = ""
            if mat_name == 'S621' and abs(r - 11.6529) < 0.001:
                gt_match = " <<<< MATCHES GT!"
            if mat_name == 'S110' and abs(r - 10.4766) < 0.001:
                gt_match = " <<<< MATCHES GT!"
            if mat_name == 'CR 88-300' and abs(r - 9.5208) < 0.001:
                gt_match = " <<<< MATCHES GT!"
            print(f"  {mat_name} {m}: Rate={r:.4f} (Amt={total_a:.2f}, Qty={total_q:.2f}){gt_match}")
        else:
            print(f"  {mat_name} {m}: NO DATA")

# ── 7. Check combined months ──
print("\n" + "-" * 80)
print("7. CHECKING IF GT USES COMBINED MONTHS")
print("-" * 80)

for mat_name, gt_rate in [('S621', 11.6529), ('S110', 10.4766), ('CR 88-300', 9.5208)]:
    mat_mask4 = pr_df[mat_col].astype(str).str.upper().str.strip() == mat_name
    rows = pr_df[adh_mask & mat_mask4]
    all_months = sorted(rows[month_col].astype(str).str.strip().unique())
    
    print(f"\n  {mat_name} (GT={gt_rate}):")
    print(f"    Available months: {all_months}")
    
    # Try combining last N months
    # Convert to sortable form
    def month_sort_key(m):
        parts = m.split('-')
        return int(parts[1]) * 100 + int(parts[0])
    
    all_months_sorted = sorted(all_months, key=month_sort_key)
    
    # Try each suffix combination
    for i in range(len(all_months_sorted)):
        subset = all_months_sorted[i:]
        subset_rows = rows[rows[month_col].astype(str).str.strip().isin(subset)]
        total_a = pd.to_numeric(subset_rows[amt_col], errors='coerce').fillna(0).sum()
        total_q = pd.to_numeric(subset_rows[qty_col], errors='coerce').fillna(0).sum()
        r = total_a / total_q if total_q > 0 else 0
        if abs(r - gt_rate) < 0.001:
            print(f"    FOUND MATCH! Months={subset}: Rate={r:.4f} <<< MATCHES GT!")
            for _, row in subset_rows.iterrows():
                print(f"      {row[month_col]}: Rate={row[rate_col]}, Qty={row[qty_col]}, Amt={row[amt_col]}")
            break

print("\n" + "=" * 100)
print("DEEP INVESTIGATION COMPLETE")
print("=" * 100)
