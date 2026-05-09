"""Quick test script for the fast pipeline."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from rmc_fill.config import RMCPaths
from rmc_fill.fast_pipeline import FastRMCPipeline

paths = RMCPaths(
    workspace=Path(r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP")
)
output = Path(__file__).parent / "output" / "rmc_fast_output.xlsx"
pipe = FastRMCPipeline(paths, output)
result = pipe.run()

print("\n========== RESULT ==========")
for k, v in result["metrics"].items():
    if k != "top_mismatches":
        print(f"  {k}: {v}")

mismatches = result["metrics"].get("top_mismatches", [])
if mismatches:
    print(f"\nTop {min(20, len(mismatches))} mismatches (sorted by |diff|):")
    for m in mismatches[:20]:
        print(
            f"  {m['order']:8s} | {m['col']:40s} | "
            f"comp={m['computed']:>14.4f} | ref={m['reference']:>14.4f} | "
            f"diff={m['diff']:>12.4f}"
        )
else:
    print("\nNo mismatches > 1.0!")
