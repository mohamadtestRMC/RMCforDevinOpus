"""Check NYLON stores lookup issue and Feb PET per-MRR logic."""
import pandas as pd
import sys
sys.path.insert(0, '.')
from engine.mrr_lookup import load_stores_recordings, lookup_mrr_with_qty

# ===== NYLON =====
print("=" * 80)
print("NYLON STORES LOOKUP FAILURE")
print("=" * 80)
stores = load_stores_recordings("Template2/Stores Recordings.xlsx")

# Search for N00694 and NYLON
order_col = wo_col = mat_col = proc_col = mrr_col = None
for c in stores.columns:
    cl = str(c).strip().lower().replace('\n',' ')
    if 'wo no' in cl or 'work order' in cl or 'wo #' in cl:
        wo_col = c
    if 'main group' in cl or cl == 'material':
        mat_col = c
    if 'application' in cl or 'process' in cl:
        proc_col = c

print(f"WO col: {wo_col}")
print(f"Mat col: {mat_col}")
print(f"Process col: {proc_col}")

# Find NYLON in Main Group
if mat_col:
    nylon_mask = stores[mat_col].astype(str).str.upper().str.contains('NYLON', na=False)
    nylon = stores[nylon_mask]
    print(f"\nNYLON in stores: {len(nylon)} rows")
    
    if wo_col:
        # Find N00694 + NYLON
        order_mask = stores[wo_col].astype(str).str.upper().str.contains('N00694', na=False)
        n_nylon = stores[nylon_mask & order_mask]
        print(f"NYLON + N00694: {len(n_nylon)} rows")
        
        if len(n_nylon) > 0:
            for _, r in n_nylon.head(5).iterrows():
                print(f"  {dict(list(r.items())[:8])}")
        
        # Check just N00694
        n_rows = stores[order_mask]
        print(f"\nAll N00694 rows: {len(n_rows)}")
        if mat_col and len(n_rows) > 0:
            mats = n_rows[mat_col].astype(str).str.strip().unique()
            print(f"  Materials: {mats}")

# Check how lookup_mrr_with_qty searches
print("\n\nDirect lookup attempts:")
for mat_name in ['NYLON', 'Nylon', 'nylon', 'NYLON PA']:
    result = lookup_mrr_with_qty(stores, mat_name, 15, 1005, 'N00694', 'PRINTING')
    if result:
        print(f"  '{mat_name}' + PRINTING: {result}")
    result = lookup_mrr_with_qty(stores, mat_name, 15, None, 'N00694')
    if result:
        print(f"  '{mat_name}' + no process: {result}")

# ===== FEB PET RATES =====
print("\n" + "=" * 80)
print("FEB PET - per-MRR rate differences")
print("=" * 80)
print("""
Row 42: GT MR#=85547/85588, GT rate=4.2200. Both MRRs have rate 4.2200.
  Engine MRRs = {85588:1857.7, 85547:1809.5, 85572:991.2, 85157:497.8, 85226:305.0}
  Engine weighted rate = 4.2983 (includes 85572@4.588, 85157@4.404)
  -> GT selected only MRRs with size=1197 that match exact rate 4.2200

Row 43: GT MR#=85226/85157/85547/85572/85588, GT rate=4.3500
  -> GT uses ALL 5 MRRs but computes different weighted avg = 4.2326
  -> But GT shows 4.3500 which doesn't match ANY weighted average
  -> Likely a MANUAL override in the GT

Row 45: GT MR#=85157/85573/85330, GT rate=4.5467
  -> Computed from GT MRRs: 4.5066 (doesn't match)
  -> Another manual override

Row 46: GT MR#=85572, GT rate=4.5880
  -> Exact match from single MRR. Engine adds 85226 which dilutes rate.
  -> RULE: GT selects fewer, more specific MRRs

Row 54: GT MR#=84080, GT rate=4.7848
  -> MEGA PACK supplier → needs MEGA PACK file rate
  -> Engine finds rate=0.9 from PR (wrong!)
  -> MEGA PACK file has 4.7848 for Feb 2026 TPE ← PERFECT MATCH!
""")
print("CONCLUSION: For PET rates (rows 42/43/45/46), the GT uses different MRR")
print("selection than the engine. These are per-MRR selection differences (~1-3%).")
print("For Row 54, it's a MEGA PACK supplier override issue.")
