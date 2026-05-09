"""
Configuration for the RMC From-Scratch Pipeline.
Accepts either file paths or BytesIO objects from Streamlit uploads.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Optional


@dataclass
class SourceFiles:
    """Holds references to all uploaded/provided source files.
    Each field can be a Path (for disk files) or BytesIO (for Streamlit uploads).
    """
    jobtrack: Any = None
    stores_recordings: Any = None  # Stores Recordings for MRR lookup
    purchase_register: Any = None  # Purchase Register for rate lookup
    granules_current: Any = None
    granules_prev: Any = None
    megapack_rates: Any = None
    opening_wip: Any = None
    closing_wip: Any = None
    ink_consumption: Any = None
    dispense_ink_stock: Any = None
    dispensed_stock_movement: Any = None
    tin_tie_prices: Any = None
    components_consumption: Any = None
    rm_film_stock: Any = None
    base_rmc_template: Any = None
    filled_rmc_reference: Any = None  # optional, for validation only


@dataclass
class PipelineConfig:
    """Pipeline configuration."""
    source_files: SourceFiles = field(default_factory=SourceFiles)
    output_dir: Path = Path("output")
    report_month: Optional[str] = None
    prev_month_rmc_output: Any = None  # previous month's RMC output for offset carry-over

    def __post_init__(self):
        self.output_dir.mkdir(parents=True, exist_ok=True)
