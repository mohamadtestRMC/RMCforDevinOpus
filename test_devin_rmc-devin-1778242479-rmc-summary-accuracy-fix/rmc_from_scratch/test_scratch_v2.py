"""Test the improved from-scratch pipeline with all fixes."""
import sys, logging, time
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent.parent))
logging.basicConfig(level=logging.INFO, format='%(message)s')

from rmc_engine.config import PipelineConfig, SourceFiles
from rmc_engine.pipeline import RMCPipeline

BASE = Path(r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP\Files_need_to_study")
UNFILLED = BASE / "Unfilled"
FILLED = BASE / "Filled_Output"

sf = SourceFiles(
    jobtrack=UNFILLED / "1 Base RMC _ 2026 February.xlsx",
    stores_recordings=UNFILLED / "3 RM FILM STOCK MAIN FILE - WORKING - 2026.xlsx",
    purchase_register=UNFILLED / "2 Purchase Register - 2021 - 2026 _Feb 26.xlsx",
    granules_current=UNFILLED / "4 Granules Recipe - February 2026.xlsx",
    granules_prev=UNFILLED / "4 Granules Recipe - January 2026.xlsx",
    megapack_rates=UNFILLED / "6 MEGAPACK Rate.xlsx",
    opening_wip=UNFILLED / "9 Opening WIP Stock.xlsx",
    closing_wip=UNFILLED / "10 Closing WIP Stock.xlsx",
    ink_consumption=UNFILLED / "5 Ink Consumption February 2026.xlsx",
    components_consumption=UNFILLED / "12 Components Consumptions Dispensed Details.xlsx",
    base_rmc_template=UNFILLED / "1 Base RMC _ 2026 February.xlsx",
    filled_rmc_reference=FILLED / "1 Base RMC _ 2026 February.xlsx",
)

config = PipelineConfig(
    source_files=sf,
    output_dir=Path("output"),
    report_month="February",
)

pipeline = RMCPipeline(config)

def progress(step, total, msg):
    print(f"  [{step}/{total}] {msg}")

result = pipeline.run(mode="scratch", progress_cb=progress)

print("\n" + "="*60)
print("RESULTS")
print("="*60)
m = result["metrics"]
print(f"Mode: {m.get('mode')}")
print(f"Orders: {m.get('rmc_summary_orders')}")
print(f"Time: {m.get('elapsed_seconds')}s")
print(f"Accuracy: {m.get('accuracy_pct', 'N/A')}%")
print(f"Exact matches: {m.get('exact_matches', 'N/A')} / {m.get('total_checks', 'N/A')}")
print(f"Close (<1): {m.get('close_lt1', 'N/A')}")
print(f"Mismatches (>1): {m.get('mismatches_gt1', 'N/A')}")

if m.get('top_mismatches'):
    print(f"\nTop 10 mismatches:")
    for mm in m['top_mismatches'][:10]:
        print(f"  {mm['order']:10s} | {mm['col']:35s} | comp={mm['computed']:12,.2f} | ref={mm['reference']:12,.2f} | diff={mm['diff']:+12,.2f}")

# Categorize mismatches by column
if m.get('top_mismatches'):
    by_col = {}
    for mm in m['top_mismatches']:
        c = mm['col']
        by_col[c] = by_col.get(c, 0) + 1
    print(f"\nMismatches by column:")
    for c, count in sorted(by_col.items(), key=lambda x: -x[1]):
        print(f"  {c:40s}: {count}")
