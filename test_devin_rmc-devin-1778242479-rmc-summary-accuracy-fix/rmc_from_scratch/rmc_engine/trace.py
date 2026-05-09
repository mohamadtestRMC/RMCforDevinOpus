"""
Order Trace Module — shows exactly how each RMC value was computed.
Allows investigation of any order to see which rows contributed to each SUMIF.
"""
from __future__ import annotations

import logging
from typing import Any, Dict, List, Optional

from rmc_engine.data_reader import OrderIndex, safe_float, safe_str

logger = logging.getLogger(__name__)


def trace_order(
    order: str,
    idx: Dict[str, OrderIndex],
    offsets: Dict[str, Dict[str, float]],
    transfer_orders: set,
    other_film_orders: set,
) -> Dict[str, Any]:
    """
    Build a complete trace for a single order, showing:
    - Which process sheets contribute
    - Individual rows from each sheet that match the order
    - The SUMIF totals
    - Any offset adjustments applied
    """
    trace = {
        "order": order,
        "is_transfer": order in transfer_orders,
        "has_other_film": order in other_film_orders,
        "offsets_applied": offsets.get(order, {}),
        "sheets": {},
    }

    if order in transfer_orders:
        trace["note"] = "Transfer order — all values zeroed out"
        return trace

    sheet_col_map = {
        "Print": {
            "order_col": "Order  No",
            "columns": [
                ("Film Input (Kgs)", "sum", "-> Printing Film Input (Kgs)"),
                ("Film Value", "sum", "-> Printing Film Value (AED)"),
                ("Dry Ink (Kgs)", "sum", "-> Dry Ink (Kgs)"),
                ("Ink Value", "sum", "-> Ink & Sol Value (AED)"),
                ("Wastage Qty (Calc)", "sum", "-> Print Wastage Qty"),
                ("Wastage Value (AED)", "sum", "-> Print Wastage Val"),
            ],
        },
        "Lam": {
            "order_col": "Order No",
            "columns": [
                ("Fresh Mat Qty", "sum", "-> Lam Fresh Mat (Kgs)"),
                ("Fresh Mat Value", "sum", "-> Lam Fresh Mat Value (AED)"),
                ("Adh+Hard Solids Qty", "sum", "-> Adh+ Hard Solids (Kgs)"),
                ("Adh+Hard +Solv Val", "sum", "-> Adh+ Hard +Sol Value (AED)"),
                ("Wastage (Calc)", "sum", "-> Lam Wastage Qty"),
                ("Wastage (AED)", "sum", "-> Lam Wastage Val"),
            ],
        },
        "Slit": {
            "order_col": "Order No",
            "columns": [
                ("Input (Kgs)", "sum", "-> Other Film Input (Kgs) [if Other Film flag]"),
                ("Slitting Input Val (AED)", "sum", "-> Other Film Value (AED) [if Other Film flag]"),
                ("Wastage (Kgs)", "sum", "-> Slit Wastage Qty"),
                ("Wastage Val (AED)", "sum", "-> Slit Wastage Val"),
            ],
        },
        "BFL": {
            "order_col": "Order No",
            "columns": [
                ("Wastage  (Kgs)", "sum", "-> BFL Wastage Qty"),
                ("Wastage  value (AED)", "sum", "-> BFL Wastage Val"),
            ],
        },
        "Bag&Pouch": {
            "order_col": "Order No",
            "columns": [
                ("PE STRIP + ZIPPER Qty", "sum", "-> part of Zipper+PE+Valve (Kgs)"),
                ("PE STRIP + ZIPPER Value", "sum", "-> part of Zipper+PE+Valve Value"),
                ("Wastage (Calc)", "sum", "-> B&P Wastage Qty"),
                ("Wastage (AED)", "sum", "-> B&P Wastage Val"),
            ],
        },
        "Spout&Valve": {
            "order_col": "Order No",
            "columns": [
                ("TIN TIE+Valve+Spout Qty", "sum", "-> part of Zipper+PE+Valve (Kgs)"),
                ("TIN TIE+Valve+Spout Value", "sum", "-> part of Zipper+PE+Valve Value"),
                ("Wastage (Calc)", "sum", "-> S&V Wastage Qty"),
                ("Wastage (AED)", "sum", "-> S&V Wastage Val"),
            ],
        },
        "HCI Rew": {
            "order_col": "Order No",
            "columns": [
                ("Sum of Wastage (Calc)", "sum", "-> HCI Wastage Qty"),
                ("Wastage Value (AED)", "sum", "-> HCI Wastage Val"),
            ],
        },
        "PTR Rew": {
            "order_col": "Order No",
            "columns": [
                ("Total Wastage Qty", "sum", "-> PTR Wastage Qty"),
                ("Total Wastage Value", "sum", "-> PTR Wastage Val"),
            ],
        },
        "OPN_WIP": {
            "order_col": "W/O",
            "columns": [
                ("Qty", "sum", "-> Opening WIP (Kg)"),
                ("Value", "sum", "-> Opening WIP Value (AED)"),
            ],
        },
        "CLS_WIP": {
            "order_col": "W/O",
            "columns": [
                ("Qty", "sum", "-> Closing WIP (Kg)"),
                ("Value", "sum", "-> Closing WIP Value (AED)"),
            ],
        },
        "FG": {
            "order_col": "Row Labels",
            "columns": [
                ("Final FG", "vlookup", "-> Prod / Output (Kg)"),
            ],
        },
    }

    for sheet_name, config in sheet_col_map.items():
        oi = idx.get(sheet_name)
        if oi is None:
            trace["sheets"][sheet_name] = {"status": "not found in workbook"}
            continue

        order_col = config["order_col"]
        oci = oi._col_map.get(order_col, -1)
        matching_rows = oi._by_order.get(order, [])

        sheet_trace = {
            "matching_rows": len(matching_rows),
            "columns": {},
        }

        if matching_rows:
            row_details = []
            for i, row in enumerate(matching_rows):
                row_info = {}
                for col_name, agg_type, target in config["columns"]:
                    ci = oi._col_map.get(col_name, -1)
                    if ci >= 0 and ci < len(row):
                        row_info[col_name] = safe_float(row[ci])
                row_details.append(row_info)
            sheet_trace["row_details"] = row_details

        for col_name, agg_type, target in config["columns"]:
            if agg_type == "sum":
                val = oi.sumif(order, col_name)
            else:
                val = oi.vlookup(order, col_name)

            ofs_key = _col_to_offset_key(target)
            offset_val = offsets.get(order, {}).get(ofs_key, 0.0) if ofs_key else 0.0
            final_val = val + offset_val

            sheet_trace["columns"][col_name] = {
                "raw_sumif": round(val, 4),
                "offset": round(offset_val, 4) if abs(offset_val) > 0.001 else 0,
                "final": round(final_val, 4),
                "target": target,
            }

        trace["sheets"][sheet_name] = sheet_trace

    return trace


def _col_to_offset_key(target_str: str) -> Optional[str]:
    """Map a target description to offset dictionary key."""
    mapping = {
        "Opening WIP (Kg)": "opn_wip_kg",
        "Printing Film Input (Kgs)": "print_film_kg",
        "Lam Fresh Mat (Kgs)": "lam_fresh_kg",
        "Dry Ink (Kgs)": "dry_ink_kg",
        "Adh+ Hard Solids (Kgs)": "adh_hard_kg",
        "Zipper+PE+Valve (Kgs)": "zip_pe_valve_kg",
        "Closing WIP (Kg)": "cls_wip_kg",
        "Opening WIP Value (AED)": "opn_wip_val",
        "Printing Film Value (AED)": "print_film_val",
        "Lam Fresh Mat Value (AED)": "lam_fresh_val",
        "Ink & Sol Value (AED)": "ink_sol_val",
        "Adh+ Hard +Sol Value (AED)": "adh_hard_val",
        "Zipper+PE+Valve Value": "zip_pe_valve_val",
        "Closing WIP Value (AED)": "cls_wip_val",
    }
    for k, v in mapping.items():
        if k in target_str:
            return v
    return None


def format_trace_for_display(trace: Dict) -> str:
    """Format a trace dict into human-readable text."""
    lines = []
    order = trace["order"]
    lines.append(f"=== TRACE: Order {order} ===")

    if trace.get("is_transfer"):
        lines.append("  [TRANSFER ORDER] All values zeroed")
        return "\n".join(lines)

    if trace.get("has_other_film"):
        lines.append("  [HAS OTHER FILM] SUMIFS on Slit for Other Film columns")

    ofs = trace.get("offsets_applied", {})
    if ofs:
        lines.append(f"  Offsets applied: {ofs}")

    for sheet_name, sheet_data in trace.get("sheets", {}).items():
        if isinstance(sheet_data, dict) and sheet_data.get("status") == "not found in workbook":
            lines.append(f"\n  [{sheet_name}] Not found in workbook")
            continue

        n_rows = sheet_data.get("matching_rows", 0)
        lines.append(f"\n  [{sheet_name}] {n_rows} matching rows")

        for col_name, col_data in sheet_data.get("columns", {}).items():
            raw = col_data.get("raw_sumif", 0)
            offset = col_data.get("offset", 0)
            final = col_data.get("final", 0)
            target = col_data.get("target", "")

            if abs(offset) > 0.001:
                lines.append(f"    {col_name}: {raw:,.4f} + offset({offset:,.4f}) = {final:,.4f}  {target}")
            elif abs(raw) > 0.001:
                lines.append(f"    {col_name}: {final:,.4f}  {target}")

    return "\n".join(lines)
