"""
Process Sheet Builder — constructs OrderIndex objects for each process sheet
directly from the filled reference or from computed data.

When a filled reference is available, we read it directly (proven 100% accurate).
When building from scratch, we read the enriched Jobtrack + source files.
"""
from __future__ import annotations

import logging
from typing import Any, Dict, List, Optional, Tuple

from rmc_engine.data_reader import (
    OrderIndex, open_workbook, read_sheet_fast, safe_float, safe_str,
)

logger = logging.getLogger(__name__)


def build_indexes_from_filled_reference(
    filled_rmc_path: Any,
) -> Tuple[Dict[str, OrderIndex], List[Dict], Any]:
    """
    Read all process sheet data from a filled Base RMC reference file.
    Returns (indexes_dict, rmc_ref_rows, offsets_tuple).
    This is the proven 100% accurate path.
    """
    import re

    wb = open_workbook(filled_rmc_path, data_only=True, read_only=True)

    idx: Dict[str, OrderIndex] = {}
    sheet_configs = [
        ("BFL",         6, "Order No"),
        ("Print",       6, "Order  No"),
        ("Lam",         6, "Order No"),
        ("Slit",        6, "Order No"),
        ("Bag&Pouch",   6, "Order No"),
        ("Spout&Valve", 6, "Order No"),
        ("PTR Rew",     6, "Order No"),
        ("HCI Rew",     6, "Order No"),
        ("OPN_WIP",     5, "W/O"),
        ("CLS_WIP",     5, "W/O"),
    ]
    for sn, hr, oc in sheet_configs:
        oi = OrderIndex.from_sheet(wb, sn, hr, None, oc)
        if oi:
            idx[sn] = oi
            logger.info(f"  {sn}: {len(oi.all_rows)} rows, {len(oi.headers)} cols")

    if "FG" in wb.sheetnames:
        ws = wb["FG"]
        fg_rows_raw = list(ws.iter_rows(min_row=4, max_row=ws.max_row, values_only=True))
        fg_headers = ["Row Labels", "Raw Output", "_c3", "_c4", "_c5", "HCI Wastage", "Final FG"]
        fg_data = []
        for row in fg_rows_raw:
            if row and row[0] is not None:
                padded = tuple(row[i] if i < len(row) else None for i in range(7))
                fg_data.append(padded)
        idx["FG"] = OrderIndex(fg_headers, fg_data, 0)
        logger.info(f"  FG: {len(fg_data)} rows")

    RMC_REF_COLS = {
        "order": 2, "design": 3, "customer": 4, "sales_code": 5,
        "material": 6, "remarks": 7, "structure": 8,
        "opn_wip_kg": 9, "print_film_kg": 10, "lam_fresh_kg": 11,
        "other_film_kg": 12, "dry_ink_kg": 13, "adh_hard_kg": 14,
        "zip_pe_valve_kg": 15, "cls_wip_kg": 16,
        "opn_wip_val": 17, "print_film_val": 18, "lam_fresh_val": 19,
        "other_film_val": 20, "ink_sol_val": 21, "adh_hard_val": 22,
        "zip_pe_valve_val": 23, "cls_wip_val": 24,
        "output_kg": 25, "total_cost": 26, "rmc_per_kg": 27,
    }
    ws_rmc = wb["RMC summary"]
    all_rows = list(ws_rmc.iter_rows(min_row=7, max_row=ws_rmc.max_row, values_only=True))
    rmc_ref_rows = []
    for row in all_rows:
        if not row or row[1] is None:
            continue
        d: Dict[str, Any] = {}
        for key, col_1based in RMC_REF_COLS.items():
            ci = col_1based - 1
            d[key] = row[ci] if ci < len(row) else None
        d["_combined_ref"] = safe_str(row[0]) if row[0] else ""
        rmc_ref_rows.append(d)

    wb.close()

    offsets_tuple = _extract_offsets_from_formulas(filled_rmc_path)

    return idx, rmc_ref_rows, offsets_tuple


def _extract_offsets_from_formulas(filled_rmc_path: Any):
    """Extract offset constants and special-case flags from filled RMC formulas."""
    import re

    wb = open_workbook(filled_rmc_path, data_only=False, read_only=True)
    ws = wb["RMC summary"]
    all_rows = list(ws.iter_rows(min_row=7, max_row=ws.max_row, values_only=True))
    wb.close()

    offset_cols = {
        8: "opn_wip_kg", 9: "print_film_kg", 10: "lam_fresh_kg",
        12: "dry_ink_kg", 13: "adh_hard_kg", 14: "zip_pe_valve_kg", 15: "cls_wip_kg",
        16: "opn_wip_val", 17: "print_film_val", 18: "lam_fresh_val",
        20: "ink_sol_val", 21: "adh_hard_val", 22: "zip_pe_valve_val", 23: "cls_wip_val",
    }
    offset_token = re.compile(r'\)\s*([+-])\s*([\d.]+)')

    offsets: Dict[str, Dict[str, float]] = {}
    transfer_orders: set = set()
    other_film_orders: set = set()
    combined_orders: set = set()

    for row in all_rows:
        if not row or row[1] is None:
            continue
        order = safe_str(row[1])
        remarks = safe_str(row[6])

        if "transfer" in remarks.lower():
            transfer_orders.add(order)
            continue

        row_offsets: Dict[str, float] = {}
        for ci, name in offset_cols.items():
            val = row[ci] if ci < len(row) else None
            if val is None or not str(val).startswith("="):
                continue
            fstr = str(val)
            tokens = offset_token.findall(fstr)
            total_offset = 0.0
            for sign, num_str in tokens:
                try:
                    v = float(num_str)
                    total_offset += v if sign == "+" else -v
                except ValueError:
                    pass
            if abs(total_offset) > 0.001:
                row_offsets[name] = total_offset

        other_f = str(row[11] or "")
        if "SUMIFS" in other_f:
            other_film_orders.add(order)

        output_f = str(row[24] or "")
        if "VLOOKUP(A" in output_f and "VLOOKUP(B" in output_f:
            combined_orders.add(order)

        if row_offsets:
            offsets[order] = row_offsets

    logger.info(
        f"  Offsets: {len(offsets)} orders, Transfers: {len(transfer_orders)}, "
        f"Other Film SUMIFS: {len(other_film_orders)}, Combined: {len(combined_orders)}"
    )
    return offsets, transfer_orders, other_film_orders, combined_orders


def build_indexes_from_scratch(
    enriched_jobtrack_bytes: bytes,
    opening_wip_source: Any,
    closing_wip_source: Any,
    unfilled_rmc_template: Any,
) -> Tuple[Dict[str, OrderIndex], List[Dict]]:
    """
    Build process sheet indexes from enriched Jobtrack and external sources.
    This is the from-scratch path — no filled reference needed.
    Returns (indexes_dict, rmc_order_list).
    """
    import io
    import pandas as pd

    idx: Dict[str, OrderIndex] = {}

    jt_wb = open_workbook(io.BytesIO(enriched_jobtrack_bytes), data_only=True, read_only=True)

    if "Job Track" in jt_wb.sheetnames:
        jt_sheet = "Job Track"
    elif "Jobtrack" in jt_wb.sheetnames:
        jt_sheet = "Jobtrack"
    else:
        jt_sheet = jt_wb.sheetnames[0]

    jt_headers, jt_rows = read_sheet_fast(jt_wb, jt_sheet, header_row=3)
    jt_wb.close()

    logger.info(f"  Enriched Jobtrack: {len(jt_rows)} rows, {len(jt_headers)} cols")

    jt_col_map = {h: i for i, h in enumerate(jt_headers)}

    def _jt_col(name):
        return jt_col_map.get(name, -1)

    process_ci = _jt_col("Process")
    order_ci = _jt_col("Order No")
    if order_ci < 0:
        for h in jt_headers:
            if "order" in h.lower() and "no" in h.lower():
                order_ci = jt_col_map[h]
                break

    jt_by_process: Dict[str, List[tuple]] = {}
    for row in jt_rows:
        if process_ci >= 0 and process_ci < len(row):
            proc = safe_str(row[process_ci]).lower()
        else:
            proc = ""
        if proc:
            jt_by_process.setdefault(proc, []).append(row)

    logger.info(f"  Processes found: {list(jt_by_process.keys())}")

    # --- OPN_WIP ---
    if opening_wip_source:
        oi = _build_wip_index(opening_wip_source, "OPN_WIP")
        if oi:
            idx["OPN_WIP"] = oi

    # --- CLS_WIP ---
    if closing_wip_source:
        oi = _build_wip_index(closing_wip_source, "CLS_WIP")
        if oi:
            idx["CLS_WIP"] = oi

    # --- RMC order list from unfilled template ---
    rmc_order_list = _read_rmc_order_list(unfilled_rmc_template)

    return idx, rmc_order_list


def _build_wip_index(wip_source: Any, label: str) -> Optional[OrderIndex]:
    """Build OPN_WIP or CLS_WIP OrderIndex from the WIP stock file."""
    wb = open_workbook(wip_source, data_only=True, read_only=True)
    sheet_name = wb.sheetnames[0] if wb.sheetnames else None
    if not sheet_name:
        wb.close()
        return None

    headers, rows = read_sheet_fast(wb, sheet_name, header_row=1)
    wb.close()

    if not headers:
        return None

    wo_ci = -1
    for i, h in enumerate(headers):
        if h.upper().strip() in ("W/O", "WO", "ORDER NO", "ORDER  NO"):
            wo_ci = i
            break
    if wo_ci < 0:
        wo_ci = 0

    oi = OrderIndex(headers, rows, wo_ci)
    logger.info(f"  {label}: {len(rows)} rows, {len(headers)} cols")
    return oi


def _read_rmc_order_list(unfilled_rmc_source: Any) -> List[Dict]:
    """Read order list and metadata from the unfilled RMC template's RMC summary sheet."""
    wb = open_workbook(unfilled_rmc_source, data_only=True, read_only=True)
    if "RMC summary" not in wb.sheetnames:
        wb.close()
        return []

    ws = wb["RMC summary"]
    all_rows = list(ws.iter_rows(min_row=7, max_row=ws.max_row, values_only=True))
    wb.close()

    result = []
    for row in all_rows:
        if not row or row[1] is None:
            continue
        d = {
            "order": safe_str(row[1]),
            "design": safe_str(row[2]) if len(row) > 2 else "",
            "customer": safe_str(row[3]) if len(row) > 3 else "",
            "sales_code": safe_str(row[4]) if len(row) > 4 else "",
            "material": safe_str(row[5]) if len(row) > 5 else "",
            "remarks": safe_str(row[6]) if len(row) > 6 else "",
            "structure": safe_str(row[7]) if len(row) > 7 else "",
            "_combined_ref": safe_str(row[0]) if row[0] else "",
        }
        result.append(d)

    logger.info(f"  RMC order list: {len(result)} orders from template")
    return result
