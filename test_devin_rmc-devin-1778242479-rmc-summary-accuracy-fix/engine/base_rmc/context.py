"""
Shared Context object — carries all loaded data and caches through the pipeline.
Every filler module reads from and writes to this context.
"""
from __future__ import annotations
import logging
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd
import openpyxl

logger = logging.getLogger(__name__)


@dataclass
class RMCContext:
    """Central context object passed through the entire pipeline."""

    # ── The workbook being filled ──
    wb: openpyxl.Workbook = None

    # ── Loaded DataFrames from source files ──
    purchase_register: pd.DataFrame = None
    stores_recordings: pd.DataFrame = None
    granules_rates: dict = field(default_factory=dict)          # {WO#: rate}
    prev_granules_rates: dict = field(default_factory=dict)     # {WO#: rate}
    megapack_rates: dict = field(default_factory=dict)          # {(year,month): {TPE:r, WPE:r}}
    ink_summary: pd.DataFrame = None                            # Ink Consumption Summary sheet
    ink_calculation: pd.DataFrame = None                        # Ink Consumption Calculation sheet
    component_consumption: pd.DataFrame = None                  # Components Dispensed
    valve_spout_prices: dict = field(default_factory=dict)      # {name: {rate_per_pc, rate_per_kg}}
    zipper_prices: dict = field(default_factory=dict)           # {code: rate_per_kg}
    opn_wip_df: pd.DataFrame = None                             # Opening WIP raw data
    cls_wip_df: pd.DataFrame = None                             # Closing WIP raw data (qty only)

    # ── Enriched Jobtrack ──
    jobtrack_df: pd.DataFrame = None                            # Full enriched Jobtrack
    report_month: str = ""                                      # e.g. "2-2026"
    report_year: int = 0
    report_month_num: int = 0

    # ── Process-level caches (populated by fillers in order) ──
    # Each maps Order No -> aggregated values
    bfl_by_order: Dict[str, dict] = field(default_factory=dict)
    print_by_order: Dict[str, dict] = field(default_factory=dict)
    lam_by_order: Dict[str, dict] = field(default_factory=dict)
    slit_by_order: Dict[str, dict] = field(default_factory=dict)
    bag_pouch_by_order: Dict[str, dict] = field(default_factory=dict)
    spout_valve_by_order: Dict[str, dict] = field(default_factory=dict)
    hci_rew_by_order: Dict[str, dict] = field(default_factory=dict)
    ptr_rew_by_order: Dict[str, dict] = field(default_factory=dict)
    embossing_by_order: Dict[str, dict] = field(default_factory=dict)
    fg_by_order: Dict[str, dict] = field(default_factory=dict)

    # ── Rate caches ──
    print_rate_cache: Dict[str, float] = field(default_factory=dict)    # order -> Print RMC/kg
    lam_rate_cache: Dict[Tuple[str, str], float] = field(default_factory=dict)  # (order, lampass) -> rate
    pivot_lam_rates: Dict[str, float] = field(default_factory=dict)     # "orderLamPass" -> avg rate
    bp_rate_cache: Dict[str, float] = field(default_factory=dict)       # order -> B&P Output RMC/kg
    slit_rate_cache: Dict[str, float] = field(default_factory=dict)     # order -> Slit RMC/kg
    ink_rate_cache: Dict[str, float] = field(default_factory=dict)      # order -> ink rate/kg
    solvent_rate: float = 0.0                                           # Monthly Ethyl Acetate rate

    # ── WIP data ──
    opn_wip_by_key: Dict[str, dict] = field(default_factory=dict)       # composite_key -> {qty, rate, value}
    cls_wip_by_key: Dict[str, dict] = field(default_factory=dict)       # composite_key -> {qty, rate, value}

    # ── Order metadata ──
    order_list: List[str] = field(default_factory=list)                 # All unique orders
    order_meta: Dict[str, dict] = field(default_factory=dict)           # order -> {design, customer, material, structure, remarks, combined_key}

    # ── Previous month RMC offsets (for carry-forward orders) ──
    prev_rmc_offsets: Dict[str, dict] = field(default_factory=dict)     # order -> {col: offset_value}

    # ── MRR → Supplier map ──
    mrr_supplier_map: dict = field(default_factory=dict)

    # ── Pipeline log ──
    log: List[str] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)

    def _log(self, msg: str):
        self.log.append(msg)
        logger.info(msg)

    def _error(self, msg: str):
        self.errors.append(msg)
        logger.error(msg)
