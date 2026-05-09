"""
OPN_WIP Filler — Populates the OPN_WIP sheet in the Base RMC workbook.

Reverse-engineered from filled output:
  - Data is pasted from file #9 (Opening WIP Stock) starting at row 6
  - Column A = composite key formula: =B&LEFT(E,1)&RIGHT(E,1)&IF(AND(...), G, "")
  - Column J = I * H  (Value = Rate × Qty)
  - Row 3 has totals: H3=sum(qty), J3=sum(value)
  - Row 5 has headers: W/O, Design Name, Mat Structure, Process, Substrate, Lam Pass, Qty, Rate, Value
"""
from __future__ import annotations
import logging
import pandas as pd

from engine.base_rmc.context import RMCContext
from engine.base_rmc.wip_keys import compute_wip_composite_key, build_wip_index

logger = logging.getLogger(__name__)

# Column positions in the OPN_WIP sheet (1-indexed)
OPN_COLS = {
    'key': 1,           # A - composite key
    'wo': 2,            # B - W/O
    'design': 3,        # C - Design Name
    'mat_structure': 4, # D - Mat Structure
    'process': 5,       # E - Process
    'substrate': 6,     # F - Substrate
    'lam_pass': 7,      # G - Lam Pass
    'qty': 8,           # H - Qty
    'rate': 9,          # I - Rate
    'value': 10,        # J - Value
}

HEADER_ROW = 5
DATA_START = 6


def _safe_float(val) -> float:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def _safe_str(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return str(val).strip()


def fill_opn_wip(ctx: RMCContext) -> None:
    """Fill the OPN_WIP sheet from the loaded Opening WIP data."""
    ctx._log("Filling OPN_WIP sheet...")

    if ctx.opn_wip_df is None or ctx.opn_wip_df.empty:
        ctx._log("  No OPN_WIP data to fill")
        return

    ws = ctx.wb['OPN_WIP'] if 'OPN_WIP' in ctx.wb.sheetnames else None
    if ws is None:
        ctx._error("  OPN_WIP sheet not found in template!")
        return

    df = ctx.opn_wip_df

    # Map DataFrame columns to our column layout
    col_map = {}
    for c in df.columns:
        cl = str(c).lower().strip()
        if cl == 'w/o' or cl == 'wo':
            col_map['wo'] = c
        elif 'design' in cl:
            col_map['design'] = c
        elif 'mat' in cl and 'structure' in cl:
            col_map['mat_structure'] = c
        elif cl == 'process':
            col_map['process'] = c
        elif 'substrate' in cl:
            col_map['substrate'] = c
        elif 'lam' in cl and 'pass' in cl:
            col_map['lam_pass'] = c
        elif cl == 'qty':
            col_map['qty'] = c
        elif cl == 'rate':
            col_map['rate'] = c
        elif cl == 'value':
            col_map['value'] = c

    # Write header row
    headers = ['', 'W/O', 'Design Name', 'Mat Structure', 'Process',
               'Substrate', 'Lam Pass', 'Qty', 'Rate', 'Value']
    for i, h in enumerate(headers):
        if h:
            ws.cell(row=HEADER_ROW, column=i + 1, value=h)

    # Write data rows
    total_qty = 0.0
    total_value = 0.0
    row_idx = DATA_START

    for _, data_row in df.iterrows():
        wo = _safe_str(data_row.get(col_map.get('wo', ''), ''))
        if not wo:
            continue

        process = _safe_str(data_row.get(col_map.get('process', ''), ''))
        lam_pass = _safe_str(data_row.get(col_map.get('lam_pass', ''), ''))
        qty = _safe_float(data_row.get(col_map.get('qty', ''), 0))
        rate = _safe_float(data_row.get(col_map.get('rate', ''), 0))
        value = rate * qty  # Always compute: J = I * H

        # Compute composite key
        key = compute_wip_composite_key(wo, process, lam_pass)

        # Write cells
        ws.cell(row=row_idx, column=OPN_COLS['key'], value=key)
        ws.cell(row=row_idx, column=OPN_COLS['wo'], value=wo)
        ws.cell(row=row_idx, column=OPN_COLS['design'],
                value=_safe_str(data_row.get(col_map.get('design', ''), '')))
        ws.cell(row=row_idx, column=OPN_COLS['mat_structure'],
                value=_safe_str(data_row.get(col_map.get('mat_structure', ''), '')))
        ws.cell(row=row_idx, column=OPN_COLS['process'], value=process)
        ws.cell(row=row_idx, column=OPN_COLS['substrate'],
                value=_safe_str(data_row.get(col_map.get('substrate', ''), '')))
        ws.cell(row=row_idx, column=OPN_COLS['lam_pass'], value=lam_pass)
        ws.cell(row=row_idx, column=OPN_COLS['qty'], value=qty if qty else None)
        ws.cell(row=row_idx, column=OPN_COLS['rate'], value=rate if rate else None)
        ws.cell(row=row_idx, column=OPN_COLS['value'], value=value if value else None)

        total_qty += qty
        total_value += value
        row_idx += 1

    # Write title / totals in row 3
    ws.cell(row=3, column=8, value=total_qty)
    ws.cell(row=3, column=10, value=total_value)

    # Build WIP index for context
    ctx.opn_wip_by_key = build_wip_index(ctx.opn_wip_df)

    ctx._log(f"  OPN_WIP filled: {row_idx - DATA_START} rows, "
             f"total_qty={total_qty:.2f}, total_value={total_value:.2f}")
    ctx._log(f"  OPN_WIP index: {len(ctx.opn_wip_by_key)} composite keys")
