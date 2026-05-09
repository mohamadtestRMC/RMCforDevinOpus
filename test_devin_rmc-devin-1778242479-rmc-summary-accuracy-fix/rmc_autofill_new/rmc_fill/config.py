from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class RMCPaths:
    workspace: Path

    @property
    def files_need_to_study(self) -> Path:
        return self.workspace / "Files_need_to_study"

    @property
    def unfilled_dir(self) -> Path:
        return self.files_need_to_study / "Unfilled"

    @property
    def filled_dir(self) -> Path:
        return self.files_need_to_study / "Filled_Output"

    @property
    def template3_dir(self) -> Path:
        return self.workspace / "Template3"

    @property
    def unfilled_base_rmc(self) -> Path:
        return self.unfilled_dir / "1 Base RMC _ 2026 February.xlsx"

    @property
    def filled_base_rmc(self) -> Path:
        return self.filled_dir / "1 Base RMC _ 2026 February.xlsx"

    @property
    def jobtrack_source(self) -> Path:
        return self.template3_dir / "Job Track Feb 26.xlsx"

    @property
    def opening_wip_source(self) -> Path:
        return self.unfilled_dir / "9 Opening WIP Stock.xlsx"

    @property
    def closing_wip_source(self) -> Path:
        return self.unfilled_dir / "10 Closing WIP Stock.xlsx"

