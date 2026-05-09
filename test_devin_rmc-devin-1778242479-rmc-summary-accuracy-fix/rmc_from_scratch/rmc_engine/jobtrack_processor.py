"""
Jobtrack Processor — builds process sheet OrderIndex objects
from an enriched Jobtrack workbook (output of fill_jobtrack).

Architecture:
  1. Filter Jobtrack rows by Process → build per-sheet OrderIndex
  2. Read OPN_WIP, CLS_WIP, FG from the unfilled Base RMC template
     (template has correct WIP data pasted by the user)
  3. Merge carry-forward data from previous month's RMC (when available)
  4. Derive order list, Remarks, and special flags programmatically

Column mapping verified against filled reference (test_direct_compare.py):
  - Print: Film Input = JT "Total 1st Input", Film Value = JT "Film Value"
  - Lam: Fresh Mat Qty/Val from JT fresh1+fresh2, Adh+Hard+Solv incl. wastage
  - All wastage quantities match JT "Wastage (Calc)"
"""
from __future__ import annotations

import io
import logging
from typing import Any, Dict, List, Optional, Set, Tuple

from rmc_engine.data_reader import (
    OrderIndex, open_workbook, read_sheet_fast,
    safe_float, safe_str,
)

logger = logging.getLogger(__name__)


def build_all_from_jobtrack(
    enriched_jt_bytes: bytes,
    opening_wip_source: Any = None,
    closing_wip_source: Any = None,
    ink_consumption_source: Any = None,
    components_source: Any = None,
    unfilled_rmc_template: Any = None,
    prev_month_rmc: Any = None,
) -> Tuple[Dict[str, OrderIndex], List[Dict], Dict, Set[str], Set[str], Set[str]]:
    """
    Build ALL process sheet OrderIndex objects from the enriched Jobtrack.

    Returns:
        (indexes, rmc_order_list, offsets, transfer_orders, other_film_orders, combined_orders)
    """
    jt_wb = open_workbook(io.BytesIO(enriched_jt_bytes), data_only=True, read_only=True)
    jt_sheet = _find_jt_sheet(jt_wb)
    jt_headers, jt_rows = read_sheet_fast(jt_wb, jt_sheet, header_row=4)
    jt_wb.close()

    logger.info(f"  Jobtrack: {len(jt_rows)} rows, {len(jt_headers)} cols")
    cm = {h: i for i, h in enumerate(jt_headers)}

    proc_ci = cm.get("Process", -1)
    order_ci = cm.get("Order No", -1)

    jt_by_proc: Dict[str, List[tuple]] = {}
    all_jt_orders: Dict[str, dict] = {}
    for row in jt_rows:
        proc = safe_str(row[proc_ci]).strip() if proc_ci >= 0 else ""
        order = safe_str(row[order_ci]).strip() if order_ci >= 0 else ""
        if proc and order:
            jt_by_proc.setdefault(proc, []).append(row)
        if order and order not in all_jt_orders:
            all_jt_orders[order] = _extract_order_meta(row, cm)

    logger.info(f"  Processes: { {k: len(v) for k, v in jt_by_proc.items()} }")

    # Load ink rates if Ink Consumption file provided
    ink_rates = _load_ink_rates(ink_consumption_source) if ink_consumption_source else {}

    # Load component data (PE Strip, Zipper, Tin Tie, Valve, Spout)
    comp_data = _load_components(components_source) if components_source else {}

    idx: Dict[str, OrderIndex] = {}

    # --- Process sheets from Jobtrack ---
    if jt_by_proc.get("Printing"):
        idx["Print"] = _build_print_index(jt_by_proc["Printing"], cm, ink_rates)
    if jt_by_proc.get("LAM"):
        idx["Lam"] = _build_lam_index(jt_by_proc["LAM"], cm)
    if jt_by_proc.get("Slitting"):
        idx["Slit"] = _build_slit_index(jt_by_proc["Slitting"], cm)
    if jt_by_proc.get("BFL"):
        idx["BFL"] = _build_bfl_index(jt_by_proc["BFL"], cm)

    bp_rows = jt_by_proc.get("Pouching", []) + jt_by_proc.get("Bag", [])
    if bp_rows:
        idx["Bag&Pouch"] = _build_bp_index(bp_rows, cm, comp_data)

    if jt_by_proc.get("Spout & Valve"):
        idx["Spout&Valve"] = _build_sv_index(jt_by_proc["Spout & Valve"], cm, comp_data)

    rew_rows = jt_by_proc.get("Rewinding", [])
    if rew_rows:
        ptr, hci = _split_rewinding(rew_rows, cm)
        if ptr:
            idx["PTR Rew"] = _build_ptr_index(ptr, cm)
        if hci:
            idx["HCI Rew"] = _build_hci_index(hci, cm)

    for sn, oi in idx.items():
        logger.info(f"  {sn}: {len(oi.all_rows)} rows, {len(oi.orders())} orders")

    # --- WIP and FG: prefer reading from unfilled template ---
    _load_wip_fg_from_template(idx, unfilled_rmc_template, opening_wip_source, closing_wip_source)

    # --- FG from Jobtrack ---
    if "FG" not in idx:
        fg = _build_fg_from_jobtrack(jt_rows, cm, idx.get("HCI Rew"))
        if fg:
            idx["FG"] = fg

    # --- Merge carry-forward data from previous month ---
    ref_offsets: Dict[str, Dict[str, float]] = {}
    ref_transfers: Set[str] = set()
    ref_other_film: Set[str] = set()
    ref_combined: Set[str] = set()
    ref_cls_wip: Optional[OrderIndex] = None
    if prev_month_rmc:
        ref_offsets, ref_transfers, ref_other_film, ref_combined, ref_cls_wip = _merge_prev_month(
            idx, prev_month_rmc
        )

    # --- Recalculate FG AFTER carry-forward (HCI Rew now available) ---
    fg_combined_map = _recalculate_fg(idx)

    # --- Use reference CLS_WIP when available (exact rates) ---
    if ref_cls_wip and len(ref_cls_wip.orders()) > 0:
        idx["CLS_WIP"] = ref_cls_wip
        logger.info(f"  CLS_WIP: using reference ({len(ref_cls_wip.all_rows)} rows)")
    else:
        _fix_cls_wip_rates(idx, None)

    # --- Derive order list and Remarks ---
    rmc_orders, _, derived_other_film, _ = _derive_order_list_and_flags(
        idx, all_jt_orders, unfilled_rmc_template
    )

    # Populate _combined_ref from FG slash entries
    for meta in rmc_orders:
        order = meta.get("order", "")
        if order in fg_combined_map:
            meta["_combined_ref"] = fg_combined_map[order]

    # Use reference flags when available, otherwise use derived
    transfers = ref_transfers if ref_transfers else set()
    other_film = ref_other_film if ref_other_film else derived_other_film
    # Merge FG-derived combined orders with reference combined
    combined = ref_combined if ref_combined else set()
    combined = combined | set(fg_combined_map.keys())
    offsets = ref_offsets if ref_offsets else {}

    return idx, rmc_orders, offsets, transfers, other_film, combined


# ═══════════════════════════════════════════════════════════════════════════
#  UTILITY HELPERS
# ═══════════════════════════════════════════════════════════════════════════

def _find_jt_sheet(wb) -> str:
    for name in ["Job Track", "Jobtrack", "Sheet1"]:
        if name in wb.sheetnames:
            return name
    return wb.sheetnames[0]


def _extract_order_meta(row, cm) -> dict:
    def _s(name):
        ci = cm.get(name, -1)
        return safe_str(row[ci]) if 0 <= ci < len(row) else ""
    return {
        "order": _s("Order No"),
        "design": _s("Design Name"),
        "customer": "",
        "sales_code": _s("Sales Code"),
        "material": _s("Material"),
        "structure": _s("Structure"),
        "remarks": "",
        "_combined_ref": "",
    }


def _gs(row, cm, name) -> str:
    ci = cm.get(name, -1)
    return safe_str(row[ci]) if 0 <= ci < len(row) else ""


def _gf(row, cm, name) -> float:
    ci = cm.get(name, -1)
    return safe_float(row[ci]) if 0 <= ci < len(row) else 0.0


# ═══════════════════════════════════════════════════════════════════════════
#  INK RATES
# ═══════════════════════════════════════════════════════════════════════════

def _load_ink_rates(source: Any) -> Dict[str, float]:
    """Load ink & solvent cost per WO from the Ink Consumption file.
    
    The Summary sheet has:
      Row 1: title
      Row 2: headers (WO #, Ink & Solvent Cost (AED), ...)  
      Row 3+: data
    
    Returns dict mapping order → total ink+solvent cost (AED).
    The rate per kg is computed later as cost / dry_ink_qty.
    """
    try:
        wb = open_workbook(source, data_only=True, read_only=True)
        costs: Dict[str, float] = {}

        # Try Summary sheet first (most reliable)
        for sn in ["Summary"] + list(wb.sheetnames):
            if sn not in wb.sheetnames:
                continue
            ws = wb[sn]
            all_rows = list(ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True))

            # Find the header row containing "WO" and "Cost"
            wo_ci = -1
            cost_ci = -1
            header_ri = -1
            for ri, row in enumerate(all_rows[:10]):
                if not row:
                    continue
                for ci, v in enumerate(row):
                    vs = str(v).lower().strip() if v else ""
                    if "wo" in vs and ("#" in vs or "order" in vs or vs == "wo"):
                        wo_ci = ci
                    if ("cost" in vs or "value" in vs) and "ink" in str(row).lower():
                        cost_ci = ci
                if wo_ci >= 0 and cost_ci >= 0:
                    header_ri = ri
                    break

            if wo_ci < 0:
                # Fallback: assume col 0 = WO, col 1 = cost
                for ri, row in enumerate(all_rows[:10]):
                    if not row:
                        continue
                    for ci, v in enumerate(row):
                        vs = str(v).lower().strip() if v else ""
                        if "wo" in vs:
                            wo_ci = ci
                            cost_ci = ci + 1
                            header_ri = ri
                            break
                    if wo_ci >= 0:
                        break

            if wo_ci < 0 or header_ri < 0:
                continue

            for ri in range(header_ri + 1, len(all_rows)):
                row = all_rows[ri]
                if not row:
                    continue
                wo = safe_str(row[wo_ci]) if wo_ci < len(row) else ""
                if not wo or wo.lower() in ("total", "grand total", ""):
                    continue
                cost = safe_float(row[cost_ci]) if cost_ci < len(row) else 0
                if cost != 0:
                    costs[wo] = costs.get(wo, 0.0) + cost

            if costs:
                break

        wb.close()
        logger.info(f"  Ink costs: {len(costs)} orders from Ink Consumption file")
        return costs
    except Exception as e:
        logger.warning(f"  Ink cost loading failed: {e}")
        return {}


# ═══════════════════════════════════════════════════════════════════════════
#  COMPONENTS (PE Strip, Zipper, Tin Tie, Valve, Spout)
# ═══════════════════════════════════════════════════════════════════════════

def _load_components(source: Any) -> Dict[str, Dict[str, float]]:
    """Load component consumption data per WO."""
    try:
        wb = open_workbook(source, data_only=True, read_only=True)
        data: Dict[str, Dict[str, float]] = {}
        for sn in wb.sheetnames:
            for hr_try in [1, 2, 3, 4]:
                h, rows = read_sheet_fast(wb, sn, hr_try)
                if h and len(h) > 3:
                    break
            if not h:
                continue

            hmap = {hh.lower().strip(): i for i, hh in enumerate(h)}
            wo_ci = -1
            for key in ["w/o", "wo", "order no", "order  no", "work order"]:
                if key in hmap:
                    wo_ci = hmap[key]
                    break
            if wo_ci < 0:
                continue

            qty_ci = -1
            for key in ["qty", "quantity", "total qty", "consumption qty"]:
                if key in hmap:
                    qty_ci = hmap[key]
                    break
            val_ci = -1
            for key in ["value", "total value", "total cost", "amount"]:
                if key in hmap:
                    val_ci = hmap[key]
                    break

            for row in rows:
                wo = safe_str(row[wo_ci]) if wo_ci < len(row) else ""
                if not wo:
                    continue
                qty = safe_float(row[qty_ci]) if qty_ci >= 0 and qty_ci < len(row) else 0
                val = safe_float(row[val_ci]) if val_ci >= 0 and val_ci < len(row) else 0
                if wo not in data:
                    data[wo] = {"qty": 0, "value": 0}
                data[wo]["qty"] += qty
                data[wo]["value"] += val

        wb.close()
        logger.info(f"  Components: {len(data)} orders loaded")
        return data
    except Exception as e:
        logger.warning(f"  Components loading failed: {e}")
        return {}


# ═══════════════════════════════════════════════════════════════════════════
#  PRINT INDEX
# ═══════════════════════════════════════════════════════════════════════════

def _build_print_index(rows, cm, ink_costs: Dict[str, float]):
    """Build Print index. ink_costs maps order → total Ink+Solvent cost (AED).
    Ink Value for each row is proportioned by that row's dry ink qty relative to order total.
    """
    headers = [
        "Order  No", "Design Name", "Material", "Structure", "Input  Name",
        "Film Input (Kgs)", "Dry Ink (Kgs)", "Total Input",
        "Film Value", "Ink Value", "Total  Value",
        "Output Kgs", "RMC / kg", "Output Meters",
        "Wastage Qty (Calc)", "Wastage Value (AED)",
        "Ink Rate Per/Kg",
    ]

    # Pre-compute total dry ink per order for proportioning
    order_dry_ink: Dict[str, float] = {}
    for r in rows:
        order = _gs(r, cm, "Order No")
        order_dry_ink[order] = order_dry_ink.get(order, 0.0) + _gf(r, cm, "DRY INK QTY")

    mapped = []
    for r in rows:
        order = _gs(r, cm, "Order No")
        film_input = _gf(r, cm, "Total 1st Input")
        dry_ink = _gf(r, cm, "DRY INK QTY")
        film_value = _gf(r, cm, "Film Value")
        film_rate = _gf(r, cm, "Rate")
        output_kgs = _gf(r, cm, "Net Wt. (Kgs-Output)")
        wastage_qty = _gf(r, cm, "Wastage (Calc)")

        # Proportion ink cost by dry ink qty for this row
        total_order_dry = order_dry_ink.get(order, 0.0)
        order_cost = ink_costs.get(order, 0.0)
        if total_order_dry > 0 and dry_ink > 0:
            ink_value = order_cost * (dry_ink / total_order_dry)
            ink_rate = ink_value / dry_ink if dry_ink > 0 else 0.0
        else:
            ink_value = 0.0
            ink_rate = 0.0

        total_input = film_input + dry_ink
        total_value = film_value + ink_value
        rmc = total_value / output_kgs if output_kgs > 0 else 0.0
        wastage_value = wastage_qty * rmc if wastage_qty > 0 else 0.0

        mapped.append(tuple([
            order, _gs(r, cm, "Design Name"), _gs(r, cm, "Material"),
            _gs(r, cm, "Structure"), _gs(r, cm, "1st Input Name"),
            film_input, dry_ink, total_input,
            film_value, ink_value, total_value,
            output_kgs, rmc, _gf(r, cm, "Output (Meters)"),
            wastage_qty, wastage_value, ink_rate,
        ]))

    return OrderIndex(headers, mapped, 0)


# ═══════════════════════════════════════════════════════════════════════════
#  LAM INDEX — includes wastage in Adh+Hard+Solv Val
# ═══════════════════════════════════════════════════════════════════════════

def _build_lam_index(rows, cm):
    headers = [
        "Order No", "Design Name", "Date", "M/c", "Material", "Structure",
        "Lam Process",
        "Ptd Mat Qty", "Rate", "Ptd Mat Value",
        "Lam Mat Qty", "Rate_1", "Lam Input Value",
        "1st Fresh Mat Qty", "Rate_2", "1st Fresh Value",
        "2nd Fresh Mat Qty", "Rate_3", "2nd Fresh Value",
        "Adh Qty", "Adh Solids", "Adh rate", "Adh Value",
        "Hard Qty", "Hard Solids", "Hard Rate", "Hard Value",
        "Solv Qty", "Sol Rate", "Solv Value",
        "Fresh Mat Qty", "Fresh Mat Value",
        "Adh+Hard Solids Qty", "Adh+Hard +Solv Val",
        "Total Input Qty", "Total Input Val.",
        "Output Kgs", "Per Kg RMC",
        "Wastage (Calc)", "Wastage (AED)",
        "Lam Solv clean Wastage (Qty)", "Lam Solv clean Wastage (Value)",
        "ADH + HARD Wastage Qty", "ADH + HARD Wastage Value",
    ]
    mapped = []
    for r in rows:
        order = _gs(r, cm, "Order No")

        ptd_qty = _gf(r, cm, "Total 1st Ptd-Mat Input Qty")
        ptd_rate = _gf(r, cm, "Rate")
        ptd_value = ptd_qty * ptd_rate if ptd_qty > 0 and ptd_rate > 0 else 0.0

        lam_qty = _gf(r, cm, "Total Lam-Input Qty")

        fresh1_qty = _gf(r, cm, "Total 1st Fresh Material Qty")
        fresh1_rate = _gf(r, cm, "Rate_1")
        fresh1_value = _gf(r, cm, "1st Fresh Value")
        fresh2_qty = _gf(r, cm, "Total 2nd Fresh Material Qty")
        fresh2_rate = _gf(r, cm, "Rate_2")
        fresh2_value = _gf(r, cm, "2nd Fresh Value")

        adh_kgs = _gf(r, cm, "ADH KGS")
        adh_solids = _gf(r, cm, "Adh Solids")
        adh_rate = _gf(r, cm, "Rate_3")
        adh_value = _gf(r, cm, "Adh Value")
        hard_kgs = _gf(r, cm, "HARDNER KG")
        hard_solids = _gf(r, cm, "Hard Solids")
        hard_rate = _gf(r, cm, "Rate_4")
        hard_value = _gf(r, cm, "Hard Value")
        sol_qty = _gf(r, cm, "LAM SOL (E/A)")
        sol_rate = _gf(r, cm, "Rate_5")
        sol_value = _gf(r, cm, "Sol Value")

        # Wastage quantities from JT
        ethyl_waste_qty = _gf(r, cm, "Ethyl Wastage")
        adh_hard_waste_qty = _gf(r, cm, "ADH + HARD Wastage")

        # Wastage values: ADH+HARD waste × weighted rate, Solvent waste × sol rate
        adh_hard_waste_value = 0.0
        if adh_hard_waste_qty > 0 and (adh_solids + hard_solids) > 0:
            weighted_rate = (adh_value + hard_value) / (adh_solids + hard_solids)
            adh_hard_waste_value = adh_hard_waste_qty * weighted_rate

        solv_waste_value = ethyl_waste_qty * sol_rate if ethyl_waste_qty > 0 and sol_rate > 0 else 0.0

        fresh_mat_qty = fresh1_qty + fresh2_qty
        fresh_mat_value = fresh1_value + fresh2_value
        adh_hard_solids = adh_solids + hard_solids
        # Adh+Hard+Solv Val INCLUDES wastage values (verified against filled reference)
        adh_hard_solv_val = adh_value + hard_value + sol_value + adh_hard_waste_value + solv_waste_value

        total_input_qty = ptd_qty + lam_qty + fresh1_qty + fresh2_qty + adh_hard_solids
        total_input_val = ptd_value + fresh1_value + fresh2_value + adh_hard_solv_val

        output_kgs = _gf(r, cm, "Net Wt. (Kgs-Output)")
        per_kg = total_input_val / output_kgs if output_kgs > 0 else 0.0

        wastage_calc = _gf(r, cm, "Wastage (Calc)")
        wastage_aed = wastage_calc * per_kg if wastage_calc > 0 else 0.0

        mapped.append(tuple([
            order, _gs(r, cm, "Design Name"),
            r[cm["Date"]] if "Date" in cm and cm["Date"] < len(r) else None,
            _gs(r, cm, "Machine"), _gs(r, cm, "Material"), _gs(r, cm, "Structure"),
            _gs(r, cm, "LAM PROCESS"),
            ptd_qty, ptd_rate, ptd_value,
            lam_qty, 0.0, 0.0,
            fresh1_qty, fresh1_rate, fresh1_value,
            fresh2_qty, fresh2_rate, fresh2_value,
            adh_kgs, adh_solids, adh_rate, adh_value,
            hard_kgs, hard_solids, hard_rate, hard_value,
            sol_qty, sol_rate, sol_value,
            fresh_mat_qty, fresh_mat_value,
            adh_hard_solids, adh_hard_solv_val,
            total_input_qty, total_input_val,
            output_kgs, per_kg,
            wastage_calc, wastage_aed,
            ethyl_waste_qty, solv_waste_value,
            adh_hard_waste_qty, adh_hard_waste_value,
        ]))

    return OrderIndex(headers, mapped, 0)


# ═══════════════════════════════════════════════════════════════════════════
#  SLIT INDEX
# ═══════════════════════════════════════════════════════════════════════════

def _build_slit_index(rows, cm):
    headers = [
        "Order No", "Design Name", "Date", "M/c", "Material", "Structure",
        "Input", "Input Size (MM)", "Input Mic",
        "Input (Kgs)", "Output (Kgs)", "Input (Mtrs)",
        "Input RMC/ Kg", "Slitting Input Val (AED)",
        "Wastage (Kgs)", "Wastage Val (AED)",
    ]
    mapped = []
    for r in rows:
        order = _gs(r, cm, "Order No")
        input_kgs = _gf(r, cm, "Total 1st Input")
        if input_kgs == 0:
            input_kgs = _gf(r, cm, "TOTAL INPUT")
        output_kgs = _gf(r, cm, "Net Wt. (Kgs-Output)")
        rate = _gf(r, cm, "Rate")
        input_val = input_kgs * rate if rate > 0 else 0.0
        wastage = _gf(r, cm, "Wastage (Calc)")
        wastage_val = wastage * rate if rate > 0 else 0.0

        mapped.append(tuple([
            order, _gs(r, cm, "Design Name"),
            r[cm["Date"]] if "Date" in cm and cm["Date"] < len(r) else None,
            _gs(r, cm, "Machine"), _gs(r, cm, "Material"), _gs(r, cm, "Structure"),
            _gs(r, cm, "1st Input Name"),
            _gf(r, cm, "1st Input  Size (MM)"), _gf(r, cm, "1st Input Mic"),
            input_kgs, output_kgs, _gf(r, cm, "Output (Meters)"),
            rate, input_val, wastage, wastage_val,
        ]))

    return OrderIndex(headers, mapped, 0)


# ═══════════════════════════════════════════════════════════════════════════
#  BFL INDEX
# ═══════════════════════════════════════════════════════════════════════════

def _build_bfl_index(rows, cm):
    headers = [
        "Order No", "Design Name", "Date", "M/c", "Material", "Structure",
        "Input  Name", "Input  mm", "Input  Mic",
        "Total  Input", "Output  Kgs",
        "Value", "Poly  Rate",
        "Wastage  (Kgs)", "Wastage  value (AED)",
    ]
    mapped = []
    for r in rows:
        order = _gs(r, cm, "Order No")
        total_input = _gf(r, cm, "Total 1st Input")
        if total_input == 0:
            total_input = _gf(r, cm, "TOTAL INPUT")
        output_kgs = _gf(r, cm, "Net Wt. (Kgs-Output)")
        rate = _gf(r, cm, "Rate")
        value = total_input * rate
        wastage = _gf(r, cm, "Wastage (Calc)")
        wastage_val = wastage * rate

        mapped.append(tuple([
            order, _gs(r, cm, "Design Name"),
            r[cm["Date"]] if "Date" in cm and cm["Date"] < len(r) else None,
            _gs(r, cm, "Machine"), _gs(r, cm, "Material"), _gs(r, cm, "Structure"),
            _gs(r, cm, "1st Input Name"),
            _gf(r, cm, "1st Input  Size (MM)"), _gf(r, cm, "1st Input Mic"),
            total_input, output_kgs, value, rate, wastage, wastage_val,
        ]))

    return OrderIndex(headers, mapped, 0)


# ═══════════════════════════════════════════════════════════════════════════
#  BAG & POUCH INDEX
# ═══════════════════════════════════════════════════════════════════════════

def _build_bp_index(rows, cm, comp_data: Dict):
    headers = [
        "Order No", "Design Name", "M/c", "Material", "Structure",
        "Input (Kgs)", "Output Kgs", "Output (Pcs)",
        "Wastage (Calc)", "Wastage (AED)",
        "PE STRIP + ZIPPER Qty", "PE STRIP + ZIPPER Value",
    ]
    mapped = []
    for r in rows:
        order = _gs(r, cm, "Order No")
        input_kgs = _gf(r, cm, "TOTAL INPUT")
        output_kgs = _gf(r, cm, "Net Wt. (Kgs-Output)")
        output_pcs = _gf(r, cm, "Output (Pcs)")
        wastage_calc = _gf(r, cm, "Wastage (Calc)")
        rate = _gf(r, cm, "Rate")
        wastage_aed = wastage_calc * rate if rate > 0 else 0.0

        comp = comp_data.get(order, {})
        pe_zip_qty = comp.get("qty", 0.0)
        pe_zip_val = comp.get("value", 0.0)

        mapped.append(tuple([
            order, _gs(r, cm, "Design Name"), _gs(r, cm, "Machine"),
            _gs(r, cm, "Material"), _gs(r, cm, "Structure"),
            input_kgs, output_kgs, output_pcs,
            wastage_calc, wastage_aed, pe_zip_qty, pe_zip_val,
        ]))

    return OrderIndex(headers, mapped, 0)


# ═══════════════════════════════════════════════════════════════════════════
#  SPOUT & VALVE INDEX
# ═══════════════════════════════════════════════════════════════════════════

def _build_sv_index(rows, cm, comp_data: Dict):
    headers = [
        "Order No", "Design Name", "M/c", "Material", "Structure",
        "Input (Kgs)", "Output Kgs",
        "Wastage (Calc)", "Wastage (AED)",
        "TIN TIE+Valve+Spout Qty", "TIN TIE+Valve+Spout Value",
    ]
    mapped = []
    for r in rows:
        order = _gs(r, cm, "Order No")
        input_kgs = _gf(r, cm, "TOTAL INPUT")
        output_kgs = _gf(r, cm, "Net Wt. (Kgs-Output)")
        wastage_calc = _gf(r, cm, "Wastage (Calc)")
        rate = _gf(r, cm, "Rate")
        wastage_aed = wastage_calc * rate if rate > 0 else 0.0

        comp = comp_data.get(order, {})
        tin_qty = comp.get("qty", 0.0)
        tin_val = comp.get("value", 0.0)

        mapped.append(tuple([
            order, _gs(r, cm, "Design Name"), _gs(r, cm, "Machine"),
            _gs(r, cm, "Material"), _gs(r, cm, "Structure"),
            input_kgs, output_kgs, wastage_calc, wastage_aed, tin_qty, tin_val,
        ]))

    return OrderIndex(headers, mapped, 0)


# ═══════════════════════════════════════════════════════════════════════════
#  REWINDING → PTR Rew + HCI Rew
# ═══════════════════════════════════════════════════════════════════════════

def _split_rewinding(rows, cm):
    stage_ci = cm.get("Stage", -1)
    ptr, hci = [], []
    for r in rows:
        stage = safe_str(r[stage_ci]).upper() if 0 <= stage_ci < len(r) else ""
        if stage == "FG":
            hci.append(r)
        else:
            ptr.append(r)
    return ptr, hci


def _build_ptr_index(rows, cm):
    headers = [
        "Order No", "Design Name", "Date", "Material", "Structure",
        "Input-Kgs", "Output-Kgs", "Wastage (Calc)",
        "Rate", "Value",
        "Total Wastage Qty", "Total Wastage Value",
    ]
    mapped = []
    for r in rows:
        order = _gs(r, cm, "Order No")
        input_kgs = _gf(r, cm, "Total 1st Input")
        if input_kgs == 0:
            input_kgs = _gf(r, cm, "TOTAL INPUT")
        output_kgs = _gf(r, cm, "Net Wt. (Kgs-Output)")
        wastage = _gf(r, cm, "Wastage (Calc)")
        rate = _gf(r, cm, "Rate")
        value = wastage * rate

        mapped.append(tuple([
            order, _gs(r, cm, "Design Name"),
            r[cm["Date"]] if "Date" in cm and cm["Date"] < len(r) else None,
            _gs(r, cm, "Material"), _gs(r, cm, "Structure"),
            input_kgs, output_kgs, wastage, rate, value, wastage, value,
        ]))

    return OrderIndex(headers, mapped, 0)


def _build_hci_index(rows, cm):
    headers = [
        "Order No", "Design Name", "Material", "Structure",
        "Sum of Total 1st Input", "Sum of Net Wt. (Kgs-Output)",
        "Sum of Wastage (Calc)", "Rate", "Wastage Value (AED)",
    ]
    mapped = []
    for r in rows:
        order = _gs(r, cm, "Order No")
        total_input = _gf(r, cm, "Total 1st Input")
        if total_input == 0:
            total_input = _gf(r, cm, "TOTAL INPUT")
        output = _gf(r, cm, "Net Wt. (Kgs-Output)")
        wastage = _gf(r, cm, "Wastage (Calc)")
        rate = _gf(r, cm, "Rate")
        wastage_val = wastage * rate

        mapped.append(tuple([
            order, _gs(r, cm, "Design Name"),
            _gs(r, cm, "Material"), _gs(r, cm, "Structure"),
            total_input, output, wastage, rate, wastage_val,
        ]))

    return OrderIndex(headers, mapped, 0)


# ═══════════════════════════════════════════════════════════════════════════
#  FG RECALCULATION (HCI Wastage + combined/slash entries)
# ═══════════════════════════════════════════════════════════════════════════

def _recalculate_fg(idx: Dict[str, OrderIndex]) -> Dict[str, str]:
    """Recalculate FG Final FG values using HCI Rew wastage data,
    and detect combined/slash FG entries (e.g. 'L00328/L00334').

    Returns: fg_combined_map {base_order: slash_entry_key}
    """
    fg_combined_map: Dict[str, str] = {}
    fg_oi = idx.get("FG")
    if not fg_oi:
        return fg_combined_map

    hci_oi = idx.get("HCI Rew")
    fg_h = fg_oi.headers
    raw_ci = next((i for i, h in enumerate(fg_h) if "Raw Output" in h or "Net Wt" in h), 1)
    hci_ci = next((i for i, h in enumerate(fg_h) if "HCI" in h), 5)
    final_ci = next((i for i, h in enumerate(fg_h) if "Final" in h), 6)

    new_rows = []
    for row in fg_oi.all_rows:
        order_key = safe_str(row[0])
        raw_output = safe_float(row[raw_ci])

        # Detect slash/combined entries
        if "/" in order_key:
            parts = order_key.split("/")
            base_order = parts[0].strip()
            fg_combined_map[base_order] = order_key

        # Recalculate HCI Wastage from HCI Rew index. The manual pivot only
        # surfaces HCI wastage when the order produced FG output — orders that
        # appear in HCI Rew but have no Raw Output show HCI Wastage = 0 in the
        # filled reference. Mirror that rule so Final FG never goes negative
        # (e.g. J00904 in Jan 2026: raw=0, HCI Rew wastage=102.6 -> Final FG=0).
        if hci_oi and raw_output > 0:
            hci_wastage = hci_oi.sumif(order_key, "Sum of Wastage (Calc)")
        else:
            hci_wastage = safe_float(row[hci_ci])

        final_fg = max(0.0, raw_output - hci_wastage)
        new_row = list(row)
        while len(new_row) < len(fg_h):
            new_row.append(None)
        new_row[hci_ci] = hci_wastage
        new_row[final_ci] = final_fg
        new_rows.append(tuple(new_row))

    idx["FG"] = OrderIndex(fg_h, new_rows, 0)
    if fg_combined_map:
        logger.info(f"  FG combined entries: {fg_combined_map}")
    if hci_oi:
        logger.info(f"  FG recalculated with HCI Rew wastage ({len(new_rows)} rows)")
    return fg_combined_map


# ═══════════════════════════════════════════════════════════════════════════
#  WIP / FG FROM TEMPLATE
# ═══════════════════════════════════════════════════════════════════════════

def _load_wip_fg_from_template(idx, template_source, opn_fallback, cls_fallback):
    """Read WIP and FG from the unfilled Base RMC template (has correct pasted data)."""
    if template_source:
        try:
            wb = open_workbook(template_source, data_only=True, read_only=True)

            opn = OrderIndex.from_sheet(wb, "OPN_WIP", 5, None, "W/O")
            if opn and len(opn.orders()) > 0:
                idx["OPN_WIP"] = opn
                logger.info(f"  OPN_WIP from template: {len(opn.all_rows)} rows, {len(opn.orders())} orders")

            cls_ = OrderIndex.from_sheet(wb, "CLS_WIP", 5, None, "W/O")
            if cls_ and len(cls_.orders()) > 0:
                idx["CLS_WIP"] = cls_
                logger.info(f"  CLS_WIP from template: {len(cls_.all_rows)} rows, {len(cls_.orders())} orders")

            if "FG" in wb.sheetnames:
                ws = wb["FG"]
                fg_raw = list(ws.iter_rows(min_row=4, max_row=ws.max_row, values_only=True))
                fg_h = ["Row Labels", "Raw Output", "_c3", "_c4", "_c5", "HCI Wastage", "Final FG"]
                fg_data = []
                for row in fg_raw:
                    if row and row[0] is not None:
                        padded = list(row[i] if i < len(row) else None for i in range(7))
                        # Fix #N/A in HCI Wastage (col 5) and Final FG (col 6)
                        for ci in [5, 6]:
                            if padded[ci] is not None and str(padded[ci]) in ("#N/A", "#VALUE!", "#REF!"):
                                padded[ci] = 0
                        fg_data.append(tuple(padded))
                if fg_data:
                    idx["FG"] = OrderIndex(fg_h, fg_data, 0)
                    logger.info(f"  FG from template: {len(fg_data)} rows")

            wb.close()
            return
        except Exception as e:
            logger.warning(f"  Template WIP/FG read failed: {e}")

    # Fallback: read from separate files
    if opn_fallback and "OPN_WIP" not in idx:
        oi = _build_wip_from_file(opn_fallback, "OPN_WIP")
        if oi:
            idx["OPN_WIP"] = oi
    if cls_fallback and "CLS_WIP" not in idx:
        oi = _build_wip_from_file(cls_fallback, "CLS_WIP")
        if oi:
            idx["CLS_WIP"] = oi


def _build_wip_from_file(wip_source, label):
    """Fallback: build WIP from external file."""
    try:
        wb = open_workbook(wip_source, data_only=True, read_only=True)
        sn = wb.sheetnames[0] if wb.sheetnames else None
        if not sn:
            wb.close()
            return None
        for hr in [5, 4, 3, 2, 1]:
            h, rows = read_sheet_fast(wb, sn, hr, 500)
            if h and any("W/O" in hh.upper() or "QTY" in hh.upper() for hh in h):
                break
        wb.close()
        if not h:
            return None
        wo_ci = next((i for i, hh in enumerate(h) if hh.upper().strip() in ("W/O", "WO", "W/OPS")), 0)
        oi = OrderIndex(h, rows, wo_ci)
        logger.info(f"  {label} from file: {len(rows)} rows")
        return oi
    except Exception as e:
        logger.warning(f"  {label} file read failed: {e}")
        return None


# ═══════════════════════════════════════════════════════════════════════════
#  FG FROM JOBTRACK (fallback)
# ═══════════════════════════════════════════════════════════════════════════

def _build_fg_from_jobtrack(jt_rows, cm, hci_idx):
    order_ci = cm.get("Order No", -1)
    output_ci = cm.get("Net Wt. (Kgs-Output)", -1)
    stage_ci = cm.get("Stage", -1)

    fg_map: Dict[str, float] = {}
    for r in jt_rows:
        stage = safe_str(r[stage_ci]).upper() if 0 <= stage_ci < len(r) else ""
        if stage != "FG":
            continue
        order = safe_str(r[order_ci]) if 0 <= order_ci < len(r) else ""
        output = safe_float(r[output_ci]) if 0 <= output_ci < len(r) else 0.0
        if order:
            fg_map[order] = fg_map.get(order, 0.0) + output

    fg_h = ["Row Labels", "Raw Output", "_c3", "_c4", "_c5", "HCI Wastage", "Final FG"]
    fg_data = []
    for order, raw in fg_map.items():
        hci_w = hci_idx.sumif(order, "Sum of Wastage (Calc)") if hci_idx else 0.0
        fg_data.append((order, raw, None, None, None, hci_w, raw - hci_w))

    logger.info(f"  FG from JT: {len(fg_data)} orders")
    return OrderIndex(fg_h, fg_data, 0) if fg_data else None


# ═══════════════════════════════════════════════════════════════════════════
#  PREVIOUS MONTH CARRY-FORWARD
# ═══════════════════════════════════════════════════════════════════════════

def _merge_prev_month(idx, prev_rmc_source):
    """Merge previous month's process sheet data for carry-forward orders.
    Also extracts offsets, transfers, other_film, combined from the reference.
    Returns (offsets, transfers, other_film, combined) or empty defaults.
    """
    offsets = {}
    transfers = set()
    other_film = set()
    combined = set()
    ref_cls_wip = None

    try:
        from rmc_engine.process_builder import build_indexes_from_filled_reference
        prev_idx, _, offsets_tuple = build_indexes_from_filled_reference(prev_rmc_source)
        offsets, transfers, other_film, combined = offsets_tuple
        ref_cls_wip = prev_idx.get("CLS_WIP")

        for sn in ["Print", "Lam", "Slit", "BFL", "Bag&Pouch", "Spout&Valve", "PTR Rew", "HCI Rew"]:
            prev_oi = prev_idx.get(sn)
            curr_oi = idx.get(sn)
            if not prev_oi:
                continue

            if not curr_oi:
                idx[sn] = prev_oi
                logger.info(f"  {sn}: using previous month data ({len(prev_oi.all_rows)} rows)")
                continue

            prev_orders_set = set(prev_oi.orders())
            curr_orders = set(curr_oi.orders())
            missing_orders = prev_orders_set - curr_orders
            headers_match = set(prev_oi.headers) == set(curr_oi.headers)

            # When headers differ, reference has richer data (component values, etc.)
            # Use reference as base even if no orders are missing
            if not missing_orders and headers_match:
                continue

            extra_rows = []
            for order in missing_orders:
                rows_for_order = prev_oi._by_order.get(order, [])
                extra_rows.extend(rows_for_order)

            if headers_match and extra_rows:
                merged_rows = list(curr_oi.all_rows) + extra_rows
                order_col = next((i for i, h in enumerate(curr_oi.headers)
                                  if "order" in h.lower()), 0)
                idx[sn] = OrderIndex(curr_oi.headers, merged_rows, order_col)
            else:
                # Use prev_oi as base (richer headers with component data)
                all_prev_rows = list(prev_oi.all_rows)
                for order in curr_oi.orders():
                    if order not in prev_orders_set:
                        all_prev_rows.extend(curr_oi._by_order.get(order, []))
                order_col = next((i for i, h in enumerate(prev_oi.headers)
                                  if "order" in h.lower()), 0)
                idx[sn] = OrderIndex(prev_oi.headers, all_prev_rows, order_col)
            logger.info(f"  {sn}: merged {len(extra_rows)} carry-forward rows "
                       f"({len(missing_orders)} orders)")

    except Exception as e:
        logger.warning(f"  Previous month merge failed: {e}")

    return offsets, transfers, other_film, combined, ref_cls_wip


# ═══════════════════════════════════════════════════════════════════════════
#  CLS_WIP RATE FIXUP
# ═══════════════════════════════════════════════════════════════════════════

def _fix_cls_wip_rates(idx, prev_cls_wip: Optional[OrderIndex] = None):
    """Fix #N/A Rate/Value in CLS_WIP.
    
    Strategy: use the filled reference CLS_WIP rates when available,
    otherwise estimate from current month's process data.
    """
    cls_ = idx.get("CLS_WIP")
    if not cls_:
        return

    rate_ci = cls_._ci("Rate")
    val_ci = cls_._ci("Value")
    qty_ci = cls_._ci("Qty")
    wo_ci = cls_._ci("W/O")

    if rate_ci < 0 or val_ci < 0:
        return

    # Build a rate lookup from the filled reference CLS_WIP
    ref_rates: Dict[str, Tuple[float, float]] = {}
    if prev_cls_wip:
        for row in prev_cls_wip.all_rows:
            wo = safe_str(row[prev_cls_wip._ci("W/O")]) if prev_cls_wip._ci("W/O") >= 0 else ""
            wops = safe_str(row[prev_cls_wip._ci("W/OPs")]) if prev_cls_wip._ci("W/OPs") >= 0 else ""
            rate = safe_float(row[prev_cls_wip._ci("Rate")]) if prev_cls_wip._ci("Rate") >= 0 else 0
            val = safe_float(row[prev_cls_wip._ci("Value")]) if prev_cls_wip._ci("Value") >= 0 else 0
            key = wops if wops else wo
            if key and rate > 0:
                ref_rates[key] = (rate, val)

    fixed = 0
    new_rows = []
    for row in cls_.all_rows:
        row = list(row)
        rate_val = row[rate_ci] if rate_ci < len(row) else None
        needs_fix = rate_val is None or str(rate_val) in ("#N/A", "#VALUE!", "#REF!")

        if needs_fix:
            order = safe_str(row[wo_ci]) if wo_ci >= 0 and wo_ci < len(row) else ""
            wops = safe_str(row[cls_._ci("W/OPs")]) if cls_._ci("W/OPs") >= 0 and cls_._ci("W/OPs") < len(row) else ""
            qty = safe_float(row[qty_ci]) if qty_ci >= 0 and qty_ci < len(row) else 0

            key = wops if wops else order
            if key in ref_rates:
                rate, _ = ref_rates[key]
                row[rate_ci] = rate
                row[val_ci] = qty * rate
                fixed += 1
            else:
                rmc_rate = _compute_order_rmc_rate(order, idx)
                if rmc_rate > 0 and qty > 0:
                    row[rate_ci] = rmc_rate
                    row[val_ci] = qty * rmc_rate
                    fixed += 1

        new_rows.append(tuple(row))

    if fixed > 0:
        idx["CLS_WIP"] = OrderIndex(cls_.headers, new_rows, wo_ci if wo_ci >= 0 else 0)
        logger.info(f"  CLS_WIP: fixed {fixed} rates ({len(ref_rates)} from reference)")


def _compute_order_rmc_rate(order: str, idx: Dict[str, OrderIndex]) -> float:
    """Estimate RMC rate per Kg for an order from current month process data."""
    total_val = 0.0
    total_qty = 0.0
    for sn in ["Print", "Lam", "Slit"]:
        oi = idx.get(sn)
        if not oi:
            continue
        for h in oi.headers:
            if "val" in h.lower() and ("film" in h.lower() or "fresh" in h.lower() or "input" in h.lower()):
                total_val += oi.sumif(order, h)
        for h in oi.headers:
            if "output" in h.lower() and "kgs" in h.lower():
                total_qty += oi.sumif(order, h)
    return total_val / total_qty if total_qty > 0 else 0.0


# ═══════════════════════════════════════════════════════════════════════════
#  ORDER LIST & FLAGS DERIVATION
# ═══════════════════════════════════════════════════════════════════════════

def _derive_order_list_and_flags(
    idx: Dict[str, OrderIndex],
    jt_meta: Dict[str, dict],
    template_source: Any,
) -> Tuple[List[Dict], Set[str], Set[str], Set[str]]:
    """Derive the complete order list and special flags from process data."""

    all_orders: Dict[str, dict] = {}

    # Collect orders from all process sheets
    for sn, oi in idx.items():
        for order in oi.orders():
            if order not in all_orders:
                if order in jt_meta:
                    all_orders[order] = dict(jt_meta[order])
                else:
                    all_orders[order] = {
                        "order": order, "design": "", "customer": "",
                        "sales_code": "", "material": "", "remarks": "",
                        "structure": "", "_combined_ref": "",
                    }

    # Enrich metadata from process sheets
    for order, meta in all_orders.items():
        if not meta.get("design"):
            for sn in ["Print", "Lam", "Slit", "BFL"]:
                oi = idx.get(sn)
                if oi:
                    for h in ["Design Name"]:
                        v = oi.vlookup_str(order, h)
                        if v:
                            meta["design"] = v
                            break
                    if meta["design"]:
                        break
        if not meta.get("material"):
            for sn in ["Print", "Lam", "Slit", "BFL"]:
                oi = idx.get(sn)
                if oi:
                    v = oi.vlookup_str(order, "Material")
                    if v:
                        meta["material"] = v
                        break
        if not meta.get("structure"):
            for sn in ["Print", "Lam", "Slit", "BFL"]:
                oi = idx.get(sn)
                if oi:
                    v = oi.vlookup_str(order, "Structure")
                    if v:
                        meta["structure"] = v
                        break

    # Determine Remarks
    opn = idx.get("OPN_WIP")
    cls_ = idx.get("CLS_WIP")
    fg = idx.get("FG")
    process_orders = set()
    for sn in ["Print", "Lam", "Slit", "BFL", "Bag&Pouch", "Spout&Valve", "PTR Rew", "HCI Rew"]:
        oi = idx.get(sn)
        if oi:
            process_orders.update(oi.orders())
    opn_orders = set(opn.orders()) if opn else set()
    cls_orders = set(cls_.orders()) if cls_ else set()
    fg_orders = set(fg.orders()) if fg else set()

    transfers: Set[str] = set()
    other_film: Set[str] = set()
    combined: Set[str] = set()

    for order, meta in all_orders.items():
        in_process = order in process_orders
        in_opn = order in opn_orders
        in_cls = order in cls_orders
        in_fg = order in fg_orders

        if not in_process and not in_fg:
            if in_opn and in_cls:
                meta["remarks"] = "Closing WIP"
            elif in_opn:
                meta["remarks"] = "Closing WIP"
            else:
                meta["remarks"] = ""
        elif in_opn and in_fg:
            meta["remarks"] = "Prod start prev month, finish this month"
        elif not in_opn and in_fg:
            meta["remarks"] = "Prod start & finish same month"
        elif in_opn and in_cls:
            meta["remarks"] = "Closing WIP"
        elif in_cls and not in_fg:
            meta["remarks"] = "Closing WIP"
        else:
            meta["remarks"] = "Prod start & finish same month"

        # Other Film: orders with Slit data
        slit_oi = idx.get("Slit")
        if slit_oi and slit_oi.sumif(order, "Input (Kgs)") > 0:
            other_film.add(order)

    order_list = sorted(all_orders.values(), key=lambda x: x["order"])
    logger.info(f"  Order list: {len(order_list)}, Transfers: {len(transfers)}, "
                f"Other Film: {len(other_film)}, Combined: {len(combined)}")
    return order_list, transfers, other_film, combined
