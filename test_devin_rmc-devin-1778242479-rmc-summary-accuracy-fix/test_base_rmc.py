"""Quick integration test using pre-stripped template to verify all fillers."""
import sys, os, time, logging
logging.basicConfig(level=logging.INFO, format='%(message)s')

sys.path.insert(0, r'c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP')

base = r'c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP'
files_dir = os.path.join(base, 'Files_need_to_study')

# Use pre-stripped template (6.5MB instead of 14MB)
template = os.path.join(base, 'output', 'template_fast.xlsx')
if not os.path.exists(template):
    template = os.path.join(files_dir, 'Unfilled', '1 Base RMC _ 2026 February.xlsx')

print(f"Template: {os.path.basename(template)} ({os.path.getsize(template):,} bytes)", flush=True)

from engine.base_rmc.orchestrator import run_pipeline

def progress(pct, msg):
    print(f"[{pct:3d}%] {msg}", flush=True)

t0 = time.time()
unfilled = os.path.join(files_dir, 'Unfilled')

try:
    result = run_pipeline(
        base_rmc_template=template,
        purchase_register_file=os.path.join(unfilled, '2 Purchase Register - 2021 - 2026 _Feb 26.xlsx'),
        stores_file=os.path.join(unfilled, '3 RM FILM STOCK MAIN FILE - WORKING - 2026.xlsx'),
        filled_jobtrack_file=os.path.join(base, 'Jobtrack_Filled_MRR_20260506_1600.xlsx'),
        granules_file=os.path.join(unfilled, '4 Granules Recipe - February 2026.xlsx'),
        prev_granules_file=os.path.join(unfilled, '4 Granules Recipe - January 2026.xlsx'),
        ink_consumption_file=os.path.join(unfilled, '5 Ink Consumption February 2026.xlsx'),
        megapack_file=os.path.join(unfilled, '6 MEGAPACK Rate.xlsx'),
        opn_wip_file=os.path.join(unfilled, '9 Opening WIP Stock.xlsx'),
        cls_wip_file=os.path.join(unfilled, '10 Closing WIP Stock.xlsx'),
        valve_spout_file=os.path.join(unfilled, '11 Price of Tin Tie, Valve & Spout 2026 updated.xlsx'),
        component_consumption_file=os.path.join(unfilled, '12 Components Consumptions Dispensed Details.xlsx'),
        ink_stock_opening_file=os.path.join(unfilled, '7 Dispense Ink Stock Opening.xlsx'),
        dispensed_movement_file=os.path.join(unfilled, '8 Dispensed Stock Movement.xlsx'),
        progress_cb=progress,
    )

    elapsed = time.time() - t0
    print(f"\n{'='*60}", flush=True)
    print(f"PIPELINE COMPLETE in {elapsed:.1f}s", flush=True)
    print(f"  Output size: {len(result['output_bytes']):,} bytes", flush=True)
    print(f"  Metrics: {result['metrics']}", flush=True)
    print(f"  Errors: {len(result['errors'])}", flush=True)
    for e in result['errors'][:10]:
        print(f"    ERROR: {e}", flush=True)

    # Save output
    out_path = os.path.join(base, 'output', 'Base_RMC_Feb2026_FILLED.xlsx')
    with open(out_path, 'wb') as f:
        f.write(result['output_bytes'])
    print(f"  Output saved: {out_path}", flush=True)

except Exception as ex:
    elapsed = time.time() - t0
    print(f"\nFAILED after {elapsed:.1f}s: {ex}", flush=True)
    import traceback
    traceback.print_exc()
