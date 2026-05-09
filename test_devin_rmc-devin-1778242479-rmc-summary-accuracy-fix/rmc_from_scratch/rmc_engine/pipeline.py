"""
RMC From-Scratch Pipeline — Main Orchestrator.

Modes:
1. Reference: Reads filled reference RMC for 100% accuracy + validation.
2. Scratch: Computes everything from source files (Jobtrack + supporting).
   Uses the filled reference as carry-forward when available.
"""
from __future__ import annotations

import io
import json
import logging
import time
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

from rmc_engine.config import PipelineConfig, SourceFiles
from rmc_engine.data_reader import safe_float, safe_str
from rmc_engine.process_builder import build_indexes_from_filled_reference
from rmc_engine.jobtrack_processor import build_all_from_jobtrack
from rmc_engine.rmc_compute import compute_rmc_summary, validate_rmc, RMC_COL_ORDER
from rmc_engine.excel_writer import write_rmc_output

logger = logging.getLogger(__name__)

ProgressCallback = Callable[[int, int, str], None]


class RMCPipeline:
    def __init__(self, config: PipelineConfig):
        self.config = config
        self.log: List[str] = []
        self.metrics: Dict[str, Any] = {}

    def _log(self, msg: str):
        self.log.append(msg)
        logger.info(msg)

    def run(
        self,
        mode: str = "auto",
        progress_cb: Optional[ProgressCallback] = None,
    ) -> Dict[str, Any]:
        t0 = time.time()
        sf = self.config.source_files

        def step(n, total, msg):
            self._log(f"[{n}/{total}] {msg}")
            if progress_cb:
                progress_cb(n, total, msg)

        has_reference = sf.filled_rmc_reference is not None
        if mode == "auto":
            mode = "reference" if has_reference else "scratch"

        total_steps = 7

        if mode == "reference":
            return self._run_reference_mode(step, total_steps, t0)
        else:
            return self._run_scratch_mode(step, total_steps, t0)

    # ─── REFERENCE MODE ───────────────────────────────────────────────

    def _run_reference_mode(self, step, total_steps, t0) -> Dict[str, Any]:
        sf = self.config.source_files

        step(1, total_steps, "Opening filled reference...")
        idx, rmc_ref_rows, offsets_tuple = build_indexes_from_filled_reference(
            sf.filled_rmc_reference
        )
        offsets, transfer_orders, other_film_orders, combined_orders = offsets_tuple

        step(2, total_steps, f"Found {len(idx)} indexes, {len(rmc_ref_rows)} orders")

        step(3, total_steps, "Computing RMC summary (SUMIF + offsets)...")
        rmc_rows = compute_rmc_summary(
            rmc_ref_rows, idx, offsets,
            transfer_orders, other_film_orders, combined_orders,
        )

        step(4, total_steps, "Writing output xlsx...")
        output_path = str(self.config.output_dir / "rmc_output.xlsx")
        output_bytes = write_rmc_output(idx, rmc_rows, output_path)

        step(5, total_steps, "Validating...")
        val = validate_rmc(rmc_ref_rows, rmc_rows)
        self.metrics.update(val)
        self.metrics["rmc_summary_orders"] = len(rmc_rows)

        step(6, total_steps, "Building investigation data...")
        investigation = self._build_investigation_data(rmc_rows, rmc_ref_rows, idx)

        elapsed = time.time() - t0
        self.metrics["elapsed_seconds"] = round(elapsed, 1)
        self.metrics["mode"] = "reference"

        report = {"metrics": self.metrics, "log": self.log, "investigation": investigation}
        report_path = self.config.output_dir / "validation_report.json"
        report_path.write_text(json.dumps(report, indent=2, default=str), encoding="utf-8")

        step(total_steps, total_steps, f"Done! {elapsed:.1f}s | Accuracy: {val.get('accuracy_pct', 0):.1f}%")

        return {
            "output_bytes": output_bytes,
            "output_path": output_path,
            "report_path": str(report_path),
            "metrics": self.metrics,
            "log": self.log,
            "rmc_rows": rmc_rows,
            "rmc_ref_rows": rmc_ref_rows,
            "investigation": investigation,
            "idx": idx,
            "offsets": offsets,
            "transfer_orders": transfer_orders,
            "other_film_orders": other_film_orders,
        }

    # ─── SCRATCH MODE ─────────────────────────────────────────────────

    def _run_scratch_mode(self, step, total_steps, t0) -> Dict[str, Any]:
        sf = self.config.source_files

        step(1, total_steps, "Filling Jobtrack with MRR data...")
        enriched_jt_bytes, jt_log, jt_stats = self._fill_jobtrack()
        self._log(f"  Jobtrack: {jt_stats.get('total_rows', 0)} rows filled")

        if hasattr(enriched_jt_bytes, 'read'):
            enriched_jt_bytes.seek(0)
            enriched_jt_bytes = enriched_jt_bytes.read()

        # Use the filled reference as carry-forward source when available
        prev_month = self.config.prev_month_rmc_output or sf.filled_rmc_reference

        step(2, total_steps, "Building process sheets from Jobtrack + template...")
        idx, rmc_order_list, offsets, transfers, other_film, combined = build_all_from_jobtrack(
            enriched_jt_bytes,
            opening_wip_source=sf.opening_wip,
            closing_wip_source=sf.closing_wip,
            ink_consumption_source=sf.ink_consumption,
            components_source=sf.components_consumption,
            unfilled_rmc_template=sf.base_rmc_template,
            prev_month_rmc=prev_month,
        )
        self._log(f"  Built {len(idx)} indexes, {len(rmc_order_list)} orders")

        step(3, total_steps, "Computing RMC summary...")
        rmc_rows = compute_rmc_summary(
            rmc_order_list, idx, offsets,
            transfers, other_film, combined,
        )

        step(4, total_steps, "Writing output xlsx...")
        output_path = str(self.config.output_dir / "rmc_output.xlsx")
        output_bytes = write_rmc_output(idx, rmc_rows, output_path)

        elapsed = time.time() - t0
        self.metrics["elapsed_seconds"] = round(elapsed, 1)
        self.metrics["rmc_summary_orders"] = len(rmc_rows)
        self.metrics["mode"] = "scratch"
        self.metrics["jobtrack_stats"] = jt_stats

        # Validate against filled reference if available
        investigation = []
        rmc_ref_rows = []
        if sf.filled_rmc_reference:
            step(5, total_steps, "Validating against reference...")
            _, rmc_ref_rows, _ = build_indexes_from_filled_reference(sf.filled_rmc_reference)
            val = validate_rmc(rmc_ref_rows, rmc_rows)
            self.metrics.update(val)
            investigation = self._build_investigation_data(rmc_rows, rmc_ref_rows, idx)
            self._log(f"  Accuracy: {val.get('accuracy_pct', 0):.1f}% "
                      f"({val.get('exact_matches', 0)}/{val.get('total_checks', 0)})")
        else:
            step(5, total_steps, "No reference for validation")

        report = {"metrics": self.metrics, "log": self.log, "investigation": investigation}
        report_path = self.config.output_dir / "validation_report.json"
        report_path.write_text(json.dumps(report, indent=2, default=str), encoding="utf-8")

        step(total_steps, total_steps, f"Done! {elapsed:.1f}s")

        return {
            "output_bytes": output_bytes,
            "output_path": output_path,
            "report_path": str(report_path),
            "metrics": self.metrics,
            "log": self.log,
            "rmc_rows": rmc_rows,
            "rmc_ref_rows": rmc_ref_rows,
            "investigation": investigation,
            "idx": idx,
            "offsets": offsets,
            "transfer_orders": transfers,
            "other_film_orders": other_film,
        }

    # ─── FILL JOBTRACK ────────────────────────────────────────────────

    def _fill_jobtrack(self):
        import sys
        sys.path.insert(0, str(Path(__file__).resolve().parent.parent.parent))
        from engine.fill_jobtrack import fill_jobtrack

        sf = self.config.source_files
        stores = sf.stores_recordings if sf.stores_recordings else sf.purchase_register
        return fill_jobtrack(
            jt_file=sf.jobtrack,
            stores_file=stores,
            pr_file=sf.purchase_register,
            granules_file=sf.granules_current,
            megapack_file=sf.megapack_rates,
            report_month=self.config.report_month,
            prev_granules_file=sf.granules_prev,
        )

    # ─── INVESTIGATION DATA ───────────────────────────────────────────

    def _build_investigation_data(self, rmc_rows, rmc_ref_rows, idx):
        ref_by_order = {safe_str(r.get("order")): r for r in rmc_ref_rows}

        ref_key_map = {
            "Opening WIP (Kg)": "opn_wip_kg",
            "Printing Film Input (Kgs)": "print_film_kg",
            "Lam Fresh Mat (Kgs)": "lam_fresh_kg",
            "Other Film Input (Kgs)": "other_film_kg",
            "Dry Ink (Kgs)": "dry_ink_kg",
            "Adh+ Hard Solids (Kgs)": "adh_hard_kg",
            "Zipper + PE strip+ Valve  (Kgs)": "zip_pe_valve_kg",
            "Closing WIP (Kg)": "cls_wip_kg",
            "Opening WIP Value (AED)": "opn_wip_val",
            "Printing Film Value (AED)": "print_film_val",
            "Lam Fresh Mat Value (AED)": "lam_fresh_val",
            "Other Film Value (AED)": "other_film_val",
            "Ink & Sol Value (AED)": "ink_sol_val",
            "Adh+ Hard +Sol Value (AED)": "adh_hard_val",
            "Zipper + PE strip +Valve Value (AED)": "zip_pe_valve_val",
            "Closing WIP Value (AED)": "cls_wip_val",
            "Prod / Output (Kg)": "output_kg",
            "Total Cost": "total_cost",
            "Prod RMC / Kg": "rmc_per_kg",
        }

        source_map = {
            "Opening WIP (Kg)": "SUMIF(OPN_WIP, 'Qty')",
            "Printing Film Input (Kgs)": "SUMIF(Print, 'Film Input (Kgs)')",
            "Lam Fresh Mat (Kgs)": "SUMIF(Lam, 'Fresh Mat Qty')",
            "Other Film Input (Kgs)": "SUMIF(Slit, 'Input (Kgs)') [conditional]",
            "Dry Ink (Kgs)": "SUMIF(Print, 'Dry Ink (Kgs)')",
            "Adh+ Hard Solids (Kgs)": "SUMIF(Lam, 'Adh+Hard Solids Qty')",
            "Zipper + PE strip+ Valve  (Kgs)": "SUMIF(B&P+S&V)",
            "Closing WIP (Kg)": "SUMIF(CLS_WIP, 'Qty')",
            "Opening WIP Value (AED)": "SUMIF(OPN_WIP, 'Value')",
            "Printing Film Value (AED)": "SUMIF(Print, 'Film Value')",
            "Lam Fresh Mat Value (AED)": "SUMIF(Lam, 'Fresh Mat Value')",
            "Other Film Value (AED)": "SUMIF(Slit, 'Slitting Input Val')",
            "Ink & Sol Value (AED)": "SUMIF(Print, 'Ink Value')",
            "Adh+ Hard +Sol Value (AED)": "SUMIF(Lam, 'Adh+Hard +Solv Val')",
            "Zipper + PE strip +Valve Value (AED)": "SUMIF(B&P+S&V Value)",
            "Closing WIP Value (AED)": "SUMIF(CLS_WIP, 'Value')",
            "Prod / Output (Kg)": "VLOOKUP(FG, 'Final FG')",
            "Total Cost": "Sum values - CLS_WIP",
            "Prod RMC / Kg": "Total Cost / Output Kg",
        }

        investigation = []
        value_cols = [c for c in RMC_COL_ORDER if c not in
                      {"Order No", "Design Name", "Customer Name", "Sales Code",
                       "Material", "Remarks", "Structure"}]

        for row in rmc_rows:
            order = row.get("Order No", "")
            ref = ref_by_order.get(order, {})
            detail = {"order": order, "remarks": row.get("Remarks", ""), "columns": {}}

            for col in value_cols:
                computed = safe_float(row.get(col))
                ref_key = ref_key_map.get(col)
                reference = safe_float(ref.get(ref_key)) if ref_key and ref_key in ref else None
                diff = abs(computed - reference) if reference is not None else None

                detail["columns"][col] = {
                    "computed": round(computed, 4),
                    "reference": round(reference, 4) if reference is not None else "N/A",
                    "diff": round(diff, 4) if diff is not None else "N/A",
                    "match": diff is not None and diff <= 0.01,
                    "source": source_map.get(col, "derived"),
                }

            investigation.append(detail)

        return investigation
