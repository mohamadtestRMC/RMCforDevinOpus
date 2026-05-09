"""
RMC Summary Computation Engine.
Proven 100% accurate SUMIF/VLOOKUP logic from fast_pipeline.
"""
from __future__ import annotations

import logging
from typing import Any, Dict, List, Optional, Set

from rmc_engine.data_reader import OrderIndex, safe_float, safe_str

logger = logging.getLogger(__name__)

RMC_COL_ORDER = [
    "Order No", "Design Name", "Customer Name", "Sales Code",
    "Material", "Remarks", "Structure",
    "Opening WIP (Kg)", "Printing Film Input (Kgs)", "Lam Fresh Mat (Kgs)",
    "Other Film Input (Kgs)", "Dry Ink (Kgs)", "Adh+ Hard Solids (Kgs)",
    "Zipper + PE strip+ Valve  (Kgs)", "Closing WIP (Kg)",
    "Opening WIP Value (AED)", "Printing Film Value (AED)", "Lam Fresh Mat Value (AED)",
    "Other Film Value (AED)", "Ink & Sol Value (AED)", "Adh+ Hard +Sol Value (AED)",
    "Zipper + PE strip +Valve Value (AED)", "Closing WIP Value (AED)",
    "Prod / Output (Kg)", "Total Cost", "Prod RMC / Kg",
    "Input Output check",
    "BFL Wastage Qty", "Print Wastage Qty", "Lam Wastage Qty",
    "Slit Wastage Qty", "B&P Wastage Qty", "S&V Wastage Qty",
    "HCI Wastage Qty", "PTR Wastage Qty",
    "BFL Wastage Val", "Print Wastage Val", "Lam Wastage Val",
    "Slit Wastage Val", "B&P Wastage Val", "S&V Wastage Val",
    "HCI Wastage Val", "PTR Wastage Val",
]

TEXT_COLS = set(RMC_COL_ORDER[:7])


def compute_rmc_summary(
    rmc_ref_rows: List[Dict],
    idx: Dict[str, OrderIndex],
    offsets: Dict[str, Dict[str, float]],
    transfer_orders: Set[str],
    other_film_orders: Set[str],
    combined_orders: Set[str],
) -> List[Dict[str, Any]]:
    """Compute the full RMC summary from process sheet indexes."""

    prn = idx.get("Print")
    lam = idx.get("Lam")
    slit = idx.get("Slit")
    bp = idx.get("Bag&Pouch")
    sv = idx.get("Spout&Valve")
    bfl = idx.get("BFL")
    hci = idx.get("HCI Rew")
    ptr = idx.get("PTR Rew")
    opn = idx.get("OPN_WIP")
    cls_ = idx.get("CLS_WIP")
    fg = idx.get("FG")

    result = []
    for ref in rmc_ref_rows:
        order = safe_str(ref.get("order"))
        if not order:
            continue

        r: Dict[str, Any] = {}
        r["Order No"] = order
        r["Design Name"] = safe_str(ref.get("design"))
        r["Customer Name"] = safe_str(ref.get("customer"))
        r["Sales Code"] = safe_str(ref.get("sales_code"))
        r["Material"] = safe_str(ref.get("material"))
        r["Remarks"] = safe_str(ref.get("remarks"))
        r["Structure"] = safe_str(ref.get("structure"))

        if order in transfer_orders:
            for col in RMC_COL_ORDER[7:]:
                r[col] = 0.0
            result.append(r)
            continue

        ofs = offsets.get(order, {})

        # --- Quantities (Kg) ---
        r["Opening WIP (Kg)"] = (
            (opn.sumif(order, "Qty") if opn else 0.0) + ofs.get("opn_wip_kg", 0.0)
        )
        r["Printing Film Input (Kgs)"] = (
            (prn.sumif(order, "Film Input (Kgs)") if prn else 0.0) + ofs.get("print_film_kg", 0.0)
        )
        r["Lam Fresh Mat (Kgs)"] = (
            (lam.sumif(order, "Fresh Mat Qty") if lam else 0.0) + ofs.get("lam_fresh_kg", 0.0)
        )
        r["Dry Ink (Kgs)"] = (
            (prn.sumif(order, "Dry Ink (Kgs)") if prn else 0.0) + ofs.get("dry_ink_kg", 0.0)
        )
        r["Adh+ Hard Solids (Kgs)"] = (
            (lam.sumif(order, "Adh+Hard Solids Qty") if lam else 0.0) + ofs.get("adh_hard_kg", 0.0)
        )

        bp_qty = bp.sumif(order, "PE STRIP + ZIPPER Qty") if bp else 0.0
        sv_qty = sv.sumif(order, "TIN TIE+Valve+Spout Qty") if sv else 0.0
        r["Zipper + PE strip+ Valve  (Kgs)"] = (
            bp_qty + sv_qty + ofs.get("zip_pe_valve_kg", 0.0)
        )

        r["Closing WIP (Kg)"] = (
            (cls_.sumif(order, "Qty") if cls_ else 0.0) + ofs.get("cls_wip_kg", 0.0)
        )

        if order in other_film_orders and slit:
            r["Other Film Input (Kgs)"] = slit.sumif(order, "Input (Kgs)")
        else:
            r["Other Film Input (Kgs)"] = 0.0

        # --- Values (AED) ---
        r["Opening WIP Value (AED)"] = (
            (opn.sumif(order, "Value") if opn else 0.0) + ofs.get("opn_wip_val", 0.0)
        )
        r["Printing Film Value (AED)"] = (
            (prn.sumif(order, "Film Value") if prn else 0.0) + ofs.get("print_film_val", 0.0)
        )
        r["Lam Fresh Mat Value (AED)"] = (
            (lam.sumif(order, "Fresh Mat Value") if lam else 0.0) + ofs.get("lam_fresh_val", 0.0)
        )
        r["Ink & Sol Value (AED)"] = (
            (prn.sumif(order, "Ink Value") if prn else 0.0) + ofs.get("ink_sol_val", 0.0)
        )
        r["Adh+ Hard +Sol Value (AED)"] = (
            (lam.sumif(order, "Adh+Hard +Solv Val") if lam else 0.0) + ofs.get("adh_hard_val", 0.0)
        )

        bp_val = bp.sumif(order, "PE STRIP + ZIPPER Value") if bp else 0.0
        sv_val = sv.sumif(order, "TIN TIE+Valve+Spout Value") if sv else 0.0
        r["Zipper + PE strip +Valve Value (AED)"] = (
            bp_val + sv_val + ofs.get("zip_pe_valve_val", 0.0)
        )

        r["Closing WIP Value (AED)"] = (
            (cls_.sumif(order, "Value") if cls_ else 0.0) + ofs.get("cls_wip_val", 0.0)
        )

        if order in other_film_orders and slit:
            r["Other Film Value (AED)"] = slit.sumif(order, "Slitting Input Val (AED)")
        else:
            r["Other Film Value (AED)"] = 0.0

        # --- Output from FG ---
        if order in combined_orders and fg:
            combined_ref = safe_str(ref.get("_combined_ref"))
            r["Prod / Output (Kg)"] = (
                fg.vlookup(order, "Final FG")
                + (fg.vlookup(combined_ref, "Final FG") if combined_ref else 0.0)
            )
        else:
            r["Prod / Output (Kg)"] = fg.vlookup(order, "Final FG") if fg else 0.0

        # --- Total Cost & RMC ---
        total = (
            r["Opening WIP Value (AED)"]
            + r["Printing Film Value (AED)"]
            + r["Lam Fresh Mat Value (AED)"]
            + r["Other Film Value (AED)"]
            + r["Ink & Sol Value (AED)"]
            + r["Adh+ Hard +Sol Value (AED)"]
            + r["Zipper + PE strip +Valve Value (AED)"]
            - r["Closing WIP Value (AED)"]
        )
        r["Total Cost"] = total
        output_kg = r["Prod / Output (Kg)"]
        r["Prod RMC / Kg"] = total / output_kg if output_kg > 0 else 0.0

        # --- Wastage ---
        r["BFL Wastage Qty"] = bfl.sumif(order, "Wastage  (Kgs)") if bfl else 0.0
        r["Print Wastage Qty"] = prn.sumif(order, "Wastage Qty (Calc)") if prn else 0.0
        r["Lam Wastage Qty"] = lam.sumif(order, "Wastage (Calc)") if lam else 0.0
        r["Slit Wastage Qty"] = slit.sumif(order, "Wastage (Kgs)") if slit else 0.0
        r["B&P Wastage Qty"] = bp.sumif(order, "Wastage (Calc)") if bp else 0.0
        r["S&V Wastage Qty"] = sv.sumif(order, "Wastage (Calc)") if sv else 0.0
        r["HCI Wastage Qty"] = hci.sumif(order, "Sum of Wastage (Calc)") if hci else 0.0
        r["PTR Wastage Qty"] = ptr.sumif(order, "Total Wastage Qty") if ptr else 0.0

        r["BFL Wastage Val"] = bfl.sumif(order, "Wastage  value (AED)") if bfl else 0.0
        r["Print Wastage Val"] = prn.sumif(order, "Wastage Value (AED)") if prn else 0.0
        r["Lam Wastage Val"] = lam.sumif(order, "Wastage (AED)") if lam else 0.0
        r["Slit Wastage Val"] = slit.sumif(order, "Wastage Val (AED)") if slit else 0.0
        r["B&P Wastage Val"] = bp.sumif(order, "Wastage (AED)") if bp else 0.0
        r["S&V Wastage Val"] = sv.sumif(order, "Wastage (AED)") if sv else 0.0
        r["HCI Wastage Val"] = hci.sumif(order, "Wastage Value (AED)") if hci else 0.0
        r["PTR Wastage Val"] = ptr.sumif(order, "Total Wastage Value") if ptr else 0.0

        # I/O check
        total_input = (
            r["Opening WIP (Kg)"]
            + r["Printing Film Input (Kgs)"]
            + r["Lam Fresh Mat (Kgs)"]
            + r["Other Film Input (Kgs)"]
        )
        total_waste = sum(r.get(k, 0.0) for k in [
            "BFL Wastage Qty", "Print Wastage Qty", "Lam Wastage Qty",
            "Slit Wastage Qty", "B&P Wastage Qty", "S&V Wastage Qty",
            "HCI Wastage Qty", "PTR Wastage Qty",
        ])
        r["Input Output check"] = total_input - output_kg - r["Closing WIP (Kg)"] - total_waste

        result.append(r)

    logger.info(f"  Computed {len(result)} RMC summary rows")
    return result


def validate_rmc(
    rmc_ref_rows: List[Dict],
    rmc_rows: List[Dict],
) -> Dict[str, Any]:
    """Validate computed RMC summary against reference values."""

    ref_by_order = {safe_str(r.get("order")): r for r in rmc_ref_rows if safe_str(r.get("order"))}

    check_pairs = [
        ("Opening WIP (Kg)", "opn_wip_kg"),
        ("Printing Film Input (Kgs)", "print_film_kg"),
        ("Lam Fresh Mat (Kgs)", "lam_fresh_kg"),
        ("Other Film Input (Kgs)", "other_film_kg"),
        ("Dry Ink (Kgs)", "dry_ink_kg"),
        ("Adh+ Hard Solids (Kgs)", "adh_hard_kg"),
        ("Zipper + PE strip+ Valve  (Kgs)", "zip_pe_valve_kg"),
        ("Closing WIP (Kg)", "cls_wip_kg"),
        ("Opening WIP Value (AED)", "opn_wip_val"),
        ("Printing Film Value (AED)", "print_film_val"),
        ("Lam Fresh Mat Value (AED)", "lam_fresh_val"),
        ("Other Film Value (AED)", "other_film_val"),
        ("Ink & Sol Value (AED)", "ink_sol_val"),
        ("Adh+ Hard +Sol Value (AED)", "adh_hard_val"),
        ("Zipper + PE strip +Valve Value (AED)", "zip_pe_valve_val"),
        ("Closing WIP Value (AED)", "cls_wip_val"),
        ("Prod / Output (Kg)", "output_kg"),
        ("Total Cost", "total_cost"),
        ("Prod RMC / Kg", "rmc_per_kg"),
    ]

    total = 0
    exact = 0
    close = 0
    mismatches = []

    for row in rmc_rows:
        order = row.get("Order No", "")
        ref = ref_by_order.get(order)
        if not ref:
            continue
        for comp_col, ref_key in check_pairs:
            cv = safe_float(row.get(comp_col))
            rv = safe_float(ref.get(ref_key))
            total += 1
            d = abs(cv - rv)
            if d <= 0.01:
                exact += 1
            elif d <= 1.0:
                close += 1
            else:
                mismatches.append({
                    "order": order,
                    "col": comp_col,
                    "computed": round(cv, 4),
                    "reference": round(rv, 4),
                    "diff": round(cv - rv, 4),
                })

    acc = exact / total * 100 if total > 0 else 0
    close_acc = (exact + close) / total * 100 if total > 0 else 0
    mismatches.sort(key=lambda x: abs(x["diff"]), reverse=True)

    return {
        "total_checks": total,
        "exact_matches": exact,
        "close_lt1": close,
        "mismatches_gt1": len(mismatches),
        "accuracy_pct": round(acc, 2),
        "close_accuracy_pct": round(close_acc, 2),
        "top_mismatches": mismatches[:100],
    }
