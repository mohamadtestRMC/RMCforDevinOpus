"""
WIP Composite Key Generator.

Reverse-engineered from the filled Base RMC:
  OPN_WIP!A = B & LEFT(E,1) & RIGHT(E,1) & IF(AND(LEFT(E,1)&RIGHT(E,1)="LM", LEFT(G,1)="L"), G, "")
  CLS_WIP!A = same formula

Examples:
  B="B01065", E="Printed, Waiting for Lam"          → "B01065Pm"
  B="L00328", E="Laminated, Waiting for Slitting"   → "L00328Lg"
  B="J00877", E="Laminated, Waiting for ...", G="L2" → "J00877LmL2"
  B="C01480", E="..., Waiting for Pouching"          → "C01480pg"
"""
from __future__ import annotations
import logging
import pandas as pd

logger = logging.getLogger(__name__)


def compute_wip_composite_key(order_no: str, process: str, lam_pass: str = "") -> str:
    """Compute the composite key for a WIP row.

    This replicates the Excel formula:
      =B & LEFT(E,1) & RIGHT(E,1) & IF(AND(LEFT(E,1)&RIGHT(E,1)="LM", LEFT(G,1)="L"), G, "")

    Args:
        order_no: W/O number (col B), e.g. "B01065"
        process: Process description (col E), e.g. "Printed, Waiting for Lam"
        lam_pass: Lam Pass (col G), e.g. "L1", "L2", "L3"

    Returns:
        Composite key string, e.g. "B01065Pm"
    """
    if not order_no or not process:
        return ""

    wo = str(order_no).strip()
    proc = str(process).strip()

    if not proc:
        return wo

    first_char = proc[0]     # LEFT(E,1)
    last_char = proc[-1]     # RIGHT(E,1)
    suffix = first_char + last_char

    # Check: IF(AND(LEFT(E,1)&RIGHT(E,1)="LM", LEFT(G,1)="L"), G, "")
    # When suffix is "Lm" (Laminated, Waiting for ... something ending in 'm')
    # AND lam_pass starts with "L", append lam_pass
    lp = str(lam_pass).strip() if lam_pass and not pd.isna(lam_pass) else ""

    if suffix.upper() == "LM" and lp and lp[0].upper() == "L":
        return wo + suffix + lp
    else:
        return wo + suffix


def build_wip_index(wip_df: pd.DataFrame) -> dict:
    """Build a {composite_key: {qty, rate, value, order, process, lam_pass}} index from WIP data.

    Args:
        wip_df: DataFrame with columns like W/O, Process, Lam Pass, Qty, Rate, Value

    Returns:
        Dict keyed by composite key with row data
    """
    if wip_df is None or wip_df.empty:
        return {}

    # Find columns dynamically
    cols = {}
    for c in wip_df.columns:
        cl = str(c).lower().strip()
        if cl == 'w/o' or cl == 'wo' or cl == 'order':
            cols['wo'] = c
        elif cl == 'process':
            cols['process'] = c
        elif cl == 'lam pass' or cl == 'lampass':
            cols['lam_pass'] = c
        elif cl == 'qty':
            cols['qty'] = c
        elif cl == 'rate':
            cols['rate'] = c
        elif cl == 'value':
            cols['value'] = c
        elif 'design' in cl:
            cols['design'] = c
        elif 'mat' in cl and 'structure' in cl:
            cols['mat_structure'] = c
        elif 'substrate' in cl:
            cols['substrate'] = c

    if 'wo' not in cols:
        logger.warning(f"WIP DataFrame missing W/O column. Available: {list(wip_df.columns)}")
        return {}

    index = {}
    for _, row in wip_df.iterrows():
        wo = str(row.get(cols['wo'], '')).strip()
        if not wo or wo == 'nan':
            continue

        process = str(row.get(cols.get('process', ''), '')).strip()
        lam_pass = str(row.get(cols.get('lam_pass', ''), '')).strip()

        key = compute_wip_composite_key(wo, process, lam_pass)
        if not key:
            continue

        qty = _safe_float(row.get(cols.get('qty', '')))
        rate = _safe_float(row.get(cols.get('rate', '')))
        value = _safe_float(row.get(cols.get('value', '')))

        # If same key appears multiple times, sum quantities
        if key in index:
            index[key]['qty'] += qty
            index[key]['value'] += value
            # Rate: use weighted average
            total_qty = index[key]['qty']
            if total_qty > 0:
                index[key]['rate'] = index[key]['value'] / total_qty
        else:
            index[key] = {
                'order': wo,
                'process': process,
                'lam_pass': lam_pass,
                'qty': qty,
                'rate': rate,
                'value': value,
                'key': key,
            }

    logger.info(f"WIP index: {len(index)} entries")
    # Log a few examples
    for i, (k, v) in enumerate(index.items()):
        if i >= 5:
            break
        logger.debug(f"  WIP[{k}] = qty={v['qty']:.2f}, rate={v['rate']:.6f}")

    return index


def _safe_float(val) -> float:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0
