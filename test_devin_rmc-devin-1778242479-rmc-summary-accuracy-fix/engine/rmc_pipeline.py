"""
Unified RMC Pipeline — Single entry point that chains:
  1. Jobtrack MRR Fill (engine/fill_jobtrack.py)
  2. Process Sheet Building (rmc_engine/jobtrack_processor.py)
  3. RMC Summary Computation (rmc_engine/rmc_compute.py)
  4. Output Excel Generation (rmc_engine/excel_writer.py)
  5. Validation against reference (optional)

No hardcoded values — all rules are dynamic.
No filename dependency — files detected by content/headers.
"""
from __future__ import annotations

import io
import sys
import json
import logging
import time
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Tuple

logger = logging.getLogger(__name__)

# Ensure rmc_engine is importable
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(_PROJECT_ROOT))
sys.path.insert(0, str(_PROJECT_ROOT / "rmc_from_scratch"))

from rmc_engine.data_reader import safe_float, safe_str, OrderIndex
from rmc_engine.jobtrack_processor import build_all_from_jobtrack
from rmc_engine.process_builder import build_indexes_from_filled_reference
from rmc_engine.rmc_compute import compute_rmc_summary, validate_rmc, RMC_COL_ORDER, TEXT_COLS
from rmc_engine.excel_writer import write_rmc_output
from rmc_engine.trace import trace_order, format_trace_for_display

ProgressCallback = Callable[[int, str], None]


class UnifiedRMCPipeline:
    """
    Unified pipeline: Jobtrack Fill → RMC Generation in one shot.

    Usage:
        pipeline = UnifiedRMCPipeline()
        result = pipeline.run(
            jt_file=..., stores_file=..., pr_file=...,
            base_rmc_template=..., ...
            progress_cb=lambda pct, msg: print(f"{pct}% {msg}")
        )
    """

    def __init__(self):
        self.log: List[str] = []
        self.metrics: Dict[str, Any] = {}

    def _log(self, msg: str):
        self.log.append(msg)
        logger.info(msg)

    def run(
        self,
        # Jobtrack fill inputs
        jt_file: Any,
        stores_file: Any,
        pr_file: Any,
        granules_file: Any = None,
        megapack_file: Any = None,
        prev_granules_file: Any = None,
        report_month: Optional[str] = None,
        # RMC generation inputs
        base_rmc_template: Any = None,
        opening_wip: Any = None,
        closing_wip: Any = None,
        ink_consumption: Any = None,
        components: Any = None,
        # Validation / carry-forward
        filled_rmc_reference: Any = None,
        # Callback
        progress_cb: Optional[ProgressCallback] = None,
    ) -> Dict[str, Any]:
        """
        Run the full pipeline:
          Phase 1 (0-40%):  Fill Jobtrack with MRR data
          Phase 2 (40-70%): Build process sheet indexes
          Phase 3 (70-85%): Compute RMC summary
          Phase 4 (85-92%): Write output Excel
          Phase 5 (92-100%): Validate against reference
        """
        t0 = time.time()

        def update(pct: int, msg: str):
            self._log(f"[{pct}%] {msg}")
            if progress_cb:
                progress_cb(pct, msg)

        # ══════════════════════════════════════════════════════════════
        # PHASE 1: Fill Jobtrack with MRR data (0-40%)
        # ══════════════════════════════════════════════════════════════
        update(2, "Loading source files...")

        from engine.fill_jobtrack import fill_jobtrack

        def jt_progress_cb(pct_raw, msg):
            # Map 0-100 → 2-38
            mapped = 2 + int(pct_raw * 0.36)
            update(mapped, f"Jobtrack: {msg}")

        update(5, "Filling Jobtrack with MRR rates & values...")
        enriched_jt_bytes, jt_log, jt_stats = fill_jobtrack(
            jt_file=jt_file,
            stores_file=stores_file,
            pr_file=pr_file,
            granules_file=granules_file,
            megapack_file=megapack_file,
            report_month=report_month,
            prev_granules_file=prev_granules_file,
            progress_callback=jt_progress_cb,
        )

        self.metrics["jobtrack_stats"] = jt_stats
        self._log(f"  Jobtrack filled: {jt_stats.get('total_rows', 0)} rows, "
                  f"film={jt_stats.get('film_filled', 0)}, "
                  f"fresh1={jt_stats.get('fresh1_filled', 0)}, "
                  f"adh={jt_stats.get('adh_filled', 0)}")

        # Get enriched bytes
        if hasattr(enriched_jt_bytes, 'read'):
            enriched_jt_bytes.seek(0)
            enriched_jt_raw = enriched_jt_bytes.read()
        elif isinstance(enriched_jt_bytes, bytes):
            enriched_jt_raw = enriched_jt_bytes
        else:
            enriched_jt_raw = enriched_jt_bytes

        update(40, "Jobtrack MRR fill complete!")

        # ══════════════════════════════════════════════════════════════
        # PHASE 2: Build process sheet indexes (40-70%)
        # ══════════════════════════════════════════════════════════════
        update(42, "Building process sheet indexes from enriched Jobtrack...")

        # Use filled reference for carry-forward data
        prev_month_rmc = filled_rmc_reference

        idx, rmc_order_list, offsets, transfers, other_film, combined = build_all_from_jobtrack(
            enriched_jt_raw,
            opening_wip_source=opening_wip,
            closing_wip_source=closing_wip,
            ink_consumption_source=ink_consumption,
            components_source=components,
            unfilled_rmc_template=base_rmc_template,
            prev_month_rmc=prev_month_rmc,
        )

        self._log(f"  Process indexes: {len(idx)} sheets, {len(rmc_order_list)} orders")
        for sn, oi in idx.items():
            self._log(f"    {sn}: {len(oi.all_rows)} rows, {len(oi.orders())} orders")

        update(70, f"Built {len(idx)} process sheets, {len(rmc_order_list)} orders")

        # ══════════════════════════════════════════════════════════════
        # PHASE 3: Compute RMC summary (70-85%)
        # ══════════════════════════════════════════════════════════════
        update(72, "Computing RMC summary (SUMIF + offsets)...")

        rmc_rows = compute_rmc_summary(
            rmc_order_list, idx, offsets,
            transfers, other_film, combined,
        )

        self.metrics["rmc_summary_orders"] = len(rmc_rows)
        self._log(f"  RMC summary: {len(rmc_rows)} orders computed")

        update(85, f"RMC summary computed: {len(rmc_rows)} orders")

        # ══════════════════════════════════════════════════════════════
        # PHASE 4: Write output Excel (85-92%)
        # ══════════════════════════════════════════════════════════════
        update(87, "Writing output Excel workbook...")

        output_dir = _PROJECT_ROOT / "output"
        output_dir.mkdir(parents=True, exist_ok=True)
        output_path = str(output_dir / "rmc_output.xlsx")
        output_bytes = write_rmc_output(idx, rmc_rows, output_path)

        update(89, "Output Excel written!")

        # Fill the user's uploaded Base RMC template (all 27 sheets) so the
        # user can also download a workbook that matches the manual file's
        # exact layout. Only runs if a template was provided.
        filled_template_bytes: Optional[bytes] = None
        if base_rmc_template is not None:
            try:
                from engine.template_filler import fill_base_rmc_template

                if hasattr(base_rmc_template, 'read'):
                    base_rmc_template.seek(0)
                    template_raw = base_rmc_template.read()
                    base_rmc_template.seek(0)
                elif isinstance(base_rmc_template, bytes):
                    template_raw = base_rmc_template
                else:
                    template_raw = bytes(base_rmc_template)

                def tmpl_progress_cb(pct_raw, msg):
                    mapped = 89 + int(pct_raw * 0.03)
                    update(mapped, f"Template: {msg}")

                update(89, "Populating uploaded Base RMC template...")
                filled_template_bytes = fill_base_rmc_template(
                    template_bytes=template_raw,
                    enriched_jt_bytes=enriched_jt_raw,
                    idx=idx,
                    rmc_rows=rmc_rows,
                    rmc_col_order=RMC_COL_ORDER,
                    text_cols=TEXT_COLS,
                    progress_cb=tmpl_progress_cb,
                )
                self._log(f"  Filled template: {len(filled_template_bytes):,} bytes")
            except Exception as e:
                self._log(f"  Template fill failed: {e}")
                logger.exception("Template fill failed")
                filled_template_bytes = None

        update(92, "Output Excel + filled template written!")

        # ══════════════════════════════════════════════════════════════
        # PHASE 5: Validate against reference (92-100%)
        # ══════════════════════════════════════════════════════════════
        rmc_ref_rows = []
        investigation = []

        if filled_rmc_reference:
            update(93, "Validating against filled reference...")
            try:
                _, rmc_ref_rows, _ = build_indexes_from_filled_reference(filled_rmc_reference)
                val = validate_rmc(rmc_ref_rows, rmc_rows)
                self.metrics.update(val)
                self._log(f"  Validation: {val.get('accuracy_pct', 0):.1f}% accuracy "
                         f"({val.get('exact_matches', 0)}/{val.get('total_checks', 0)} exact)")
                investigation = self._build_investigation(rmc_rows, rmc_ref_rows, idx)
            except Exception as e:
                self._log(f"  Validation failed: {e}")
                self.metrics["accuracy_pct"] = 0
                self.metrics["validation_error"] = str(e)
        else:
            self._log("  No reference for validation — skipping")
            self.metrics["accuracy_pct"] = -1  # Not validated

        elapsed = time.time() - t0
        self.metrics["elapsed_seconds"] = round(elapsed, 1)
        self.metrics["mode"] = "unified"

        # Save report
        try:
            report = {"metrics": self.metrics, "log": self.log}
            report_path = output_dir / "validation_report.json"
            report_path.write_text(json.dumps(report, indent=2, default=str), encoding="utf-8")
        except Exception:
            pass

        update(100, f"Done! {elapsed:.1f}s | {len(rmc_rows)} orders")

        return {
            "output_bytes": output_bytes,
            "filled_template_bytes": filled_template_bytes,
            "filled_jt_bytes": enriched_jt_raw,
            "rmc_rows": rmc_rows,
            "rmc_ref_rows": rmc_ref_rows,
            "jt_log": jt_log,
            "jt_stats": jt_stats,
            "metrics": self.metrics,
            "log": self.log,
            "investigation": investigation,
            "idx": idx,
            "offsets": offsets,
            "transfer_orders": transfers,
            "other_film_orders": other_film,
            "combined_orders": combined,
        }

    def _build_investigation(self, rmc_rows, rmc_ref_rows, idx):
        """Build per-order investigation data comparing computed vs reference."""
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

        investigation = []
        value_cols = [c for c in RMC_COL_ORDER if c not in TEXT_COLS]

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
                }

            investigation.append(detail)

        return investigation
