from __future__ import annotations

import argparse
from pathlib import Path

from .config import RMCPaths
from .pipeline import RMCAutofillPipeline


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="RMC autofill pipeline (new isolated implementation)")
    parser.add_argument(
        "--workspace",
        type=Path,
        required=True,
        help="Workspace root path (contains Files_need_to_study and Template3)",
    )
    parser.add_argument(
        "--output",
        type=Path,
        required=True,
        help="Output xlsx path for generated workbook",
    )
    parser.add_argument(
        "--copy-jobtrack",
        action="store_true",
        help="Also overwrite Jobtrack sheet from Template3 Job Track file (slower)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    paths = RMCPaths(workspace=args.workspace)
    pipeline = RMCAutofillPipeline(
        paths=paths,
        output_file=args.output,
        copy_jobtrack=args.copy_jobtrack,
    )
    result = pipeline.run()
    print("Output workbook:", result.output_file)
    print("Validation report:", result.validation_report_file)
    print("Metrics:")
    for k, v in sorted(result.metrics.items()):
        print(f"  - {k}: {v}")


if __name__ == "__main__":
    main()

