"""Deep investigation of Row 111 rate 5.075 source."""
import sys, os, pandas as pd
sys.stdout.reconfigure(encoding='utf-8', errors='replace')

BASE = r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP\Template3"

# Load PR
pr_df = pd.read_excel(os.path.join(BASE, "Purchase Register - 2021 - 2026 _Feb 26.xlsx"), sheet_name=0, header=2)
col_map = {c: str(c).strip().replace('\n', ' ').replace('\r', '') for c in pr_df.columns}
pr_df.rename(columns=col_map, inplace=True)

# Find columns
mat_col = [c for c in pr_df.columns if 'material' in str(c).lower()][0]
cat_col = [c for c in pr_df.columns if 'categery' in str(c).lower() or 'category' in str(c).lower()][0]
rate_col = [c for c in pr_df.columns if str(c).strip().lower() == 'rate'][0]
mic_col = [c for c in pr_df.columns if str(c).strip().lower() == 'mic'][0]
month_col = [c for c in pr_df.columns if str(c).strip().lower() == 'month'][0]
amt_col = [c for c in pr_df.columns if str(c).strip().lower() == 'amount'][0]
qty_col = None
for c in pr_df.columns:
    cl = str(c).lower().strip()
    if 'actual' in cl and 'quant' in cl:
        qty_col = c
        break
if not qty_col:
    qty_col = [c for c in pr_df.columns if 'quantity' in str(c).lower()][0] if any('quantity' in str(c).lower() for c in pr_df.columns) else 'Actual Quantity'
print(f"Qty col: {qty_col}")
tracking_col = [c for c in pr_df.columns if 'tracking' in str(c).lower()][0]

print("=" * 70)
print("ROW 111: WPE (PE WHITE) rate investigation")
print("=" * 70)
print(f"Row 111: LAM, WO=J00877, Fresh1=WPE, Mic=100, Rate=5.075672...")
print()

# The rate comes from lookup_material_rate_for_month for WPE/mic=100
# Let's trace it step by step

# Step 1: Find all PR entries for WPE/PE WHITE
wpe_aliases = ['WPE', 'PE WHITE', 'WLDPE']
def mat_match(x):
    xu = str(x).strip().upper()
    for alias in wpe_aliases:
        if xu == alias or xu.startswith(alias + ' '):
            return True
    return False

wpe_mask = pr_df[mat_col].apply(mat_match)
print(f"All WPE/PE WHITE entries in PR: {wpe_mask.sum()} rows")

# Filter by Film category
cat_mask = pr_df[cat_col].astype(str).str.lower().str.contains('film', na=False)
combined = wpe_mask & cat_mask
print(f"WPE + Film category: {combined.sum()} rows")

wpe_film = pr_df[combined]
if not wpe_film.empty:
    print(f"\nAll WPE Film entries:")
    print(wpe_film[[tracking_col, mat_col, mic_col, rate_col, amt_col, qty_col, month_col]].to_string())

# Step 2: Filter by month 2-2026
month_filtered = wpe_film[wpe_film[month_col].astype(str).str.strip() == '2-2026']
print(f"\nWPE Film entries for month 2-2026: {len(month_filtered)} rows")
if not month_filtered.empty:
    print(month_filtered[[tracking_col, mat_col, mic_col, rate_col, amt_col, qty_col, month_col]].to_string())

# Step 3: Filter by mic=100
if not month_filtered.empty:
    mic_filtered = month_filtered[pd.to_numeric(month_filtered[mic_col], errors='coerce') == 100]
    print(f"\nWPE Film, month=2-2026, mic=100: {len(mic_filtered)} rows")
    if not mic_filtered.empty:
        print(mic_filtered[[tracking_col, mat_col, mic_col, rate_col, amt_col, qty_col]].to_string())
        
        # Calculate qty-weighted rate
        amounts = pd.to_numeric(mic_filtered[amt_col], errors='coerce').fillna(0)
        qtys = pd.to_numeric(mic_filtered[qty_col], errors='coerce').fillna(0)
        total_amt = amounts.sum()
        total_qty = qtys.sum()
        if total_qty > 0:
            calc_rate = total_amt / total_qty
            print(f"\nCalculated rate: Total Amount / Total Qty = {total_amt} / {total_qty} = {calc_rate}")
            print(f"This matches the engine rate: 5.075672...")

# Also check: What MRRs does J00877 have in Stores for WPE/LAMINATION?
stores_df = pd.read_excel(os.path.join(BASE, "Stores Recordings.xlsx"), sheet_name=0, header=1)
scol_map = {c: str(c).strip().replace('\n', ' ').replace('\r', '') for c in stores_df.columns}
stores_df.rename(columns=scol_map, inplace=True)

stores_cols = {}
for c in stores_df.columns:
    cl = str(c).lower().strip()
    if 'sub' in cl and 'cat' in cl: stores_cols['sub_cat'] = c
    elif cl == 'mic': stores_cols['mic'] = c
    elif 'm.r.r' in cl and 'no' in cl: stores_cols['mrr'] = c
    elif 'issue wo' in cl: stores_cols['wo'] = c
    elif 'issue' in cl and 'process' in cl: stores_cols['process'] = c
    elif 'issue' in cl and 'qty' in cl: stores_cols['issue_qty'] = c

print("\n" + "=" * 70)
print("Stores entries for WO=J00877 + PE WHITE + LAMINATION:")
print("=" * 70)
wo_mask = stores_df[stores_cols['wo']].astype(str).str.strip() == 'J00877'
mat_mask2 = stores_df[stores_cols['sub_cat']].astype(str).str.upper().str.strip().str.contains('PE WHITE', na=False)
proc_mask = stores_df[stores_cols['process']].astype(str).str.upper().str.strip().str.contains('LAMINATION', na=False)
result = stores_df[wo_mask & mat_mask2 & proc_mask]
if not result.empty:
    print(result[[stores_cols['sub_cat'], stores_cols['mic'], stores_cols['wo'], 
                   stores_cols['process'], stores_cols['issue_qty'], stores_cols['mrr']]].to_string())
    
    # Check if these MRRs exist in PR
    mrrs = pd.to_numeric(result[stores_cols['mrr']], errors='coerce').dropna().unique()
    print(f"\nMRRs found: {sorted([int(m) for m in mrrs])}")
    for mrr in sorted(mrrs):
        mrr_in_pr = pr_df[pd.to_numeric(pr_df[tracking_col], errors='coerce') == int(mrr)]
        print(f"  MRR {int(mrr)} in PR: {len(mrr_in_pr)} rows", end="")
        if not mrr_in_pr.empty:
            rates = mrr_in_pr[rate_col].tolist()
            mats = mrr_in_pr[mat_col].tolist()
            print(f" -> Material={mats[0]}, Rate={rates[0]}")
        else:
            print(" -> NOT IN PR!")

print("\nDONE!")
