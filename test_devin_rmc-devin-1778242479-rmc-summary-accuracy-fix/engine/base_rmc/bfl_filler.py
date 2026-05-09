"""
BFL Filler — Fills the BFL (Blown Film / Extrusion) sheet FROM Jobtrack data.

The unfilled template has an EMPTY BFL sheet (only headers).
This filler:
  1. Reads BFL rows from the enriched Jobtrack (ctx.jobtrack_df, Process='BFL')
  2. WRITES data rows into the BFL sheet
  3. Looks up Poly Rate from Granules Recipe / Purchase Register
  4. Computes Wastage, Wastage Value, Output RMC/KG

  BFL Sheet Layout (filled reference):
    Row 2: C2 = month text
    Row 3: C3 = "BFL"
    Row 5: subtotal formulas
    Row 6: headers
    Row 7+: data rows

  Columns (1-indexed):
    B=Order No, C=Design Name, D=Date, E=M/c, F=Material, G=Structure,
    H=Input Name, I=Input mm, J=Input Mic, K=Total Input, L=Output Kgs,
    M=Output Mtrs, N=Value, O=Poly Rate, P=Wastage Kgs, Q=Wastage Val,
    R=check, S=Sum of Waste, U=Product Type, V=Wastage(AED)
"""
from __future__ import annotations
import logging
import pandas as pd

from engine.base_rmc.context import RMCContext

logger = logging.getLogger(__name__)

# BFL sheet column positions (1-indexed) — matches filled reference
COL = {
    'order_no': 2,       # B
    'design': 3,         # C
    'date': 4,           # D
    'machine': 5,        # E
    'material': 6,       # F
    'structure': 7,      # G
    'input_name': 8,     # H
    'input_mm': 9,       # I
    'input_mic': 10,     # J
    'total_input': 11,   # K - Total Input Kgs
    'output_kgs': 12,    # L
    'output_mtrs': 13,   # M
    'value': 14,         # N - Film Value = Output Kgs * Poly Rate
    'poly_rate': 15,     # O - TO BE FILLED
    'wastage_kgs': 16,   # P = K - L
    'wastage_val': 17,   # Q = P * O
    'check': 18,         # R = Order No (repeat)
    'sum_waste': 19,     # S = Waste (same as P)
    'product_type': 21,  # U
    'wastage_aed': 22,   # V = same as Q
}

HEADER_ROW = 6
DATA_START = 7
SUBTOTAL_ROW = 5

# Jobtrack column names (from header row 4)
JT = {
    'process': 'Process',
    'order': 'Order No',
    'design': 'Design Name',
    'date': 'Date',
    'machine': 'Machine',
    'application': 'Application',
    'input_name': '1st Input Name',
    'input_mm': '1st Input  Size (MM)',
    'input_mic': '1st Input Mic',
    'total_input': 'TOTAL INPUT',
    'output_kgs': 'Net Wt. (Kgs-Output)',
    'output_mtrs': 'Output (Meters)',
    'waste': 'Waste',
    'stage': 'Stage',
}


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


def fill_bfl(ctx: RMCContext) -> None:
    """Fill the BFL sheet by writing BFL rows from Jobtrack + computing Poly Rates."""
    ctx._log("Filling BFL sheet...")

    ws = ctx.wb['BFL'] if 'BFL' in ctx.wb.sheetnames else None
    if ws is None:
        ctx._error("  BFL sheet not found!")
        return

    if ctx.jobtrack_df is None or ctx.jobtrack_df.empty:
        ctx._error("  No Jobtrack data for BFL")
        return

    # Filter Jobtrack for BFL rows
    df = ctx.jobtrack_df
    proc_col = None
    for c in df.columns:
        if str(c).strip().lower() == 'process':
            proc_col = c
            break
    if proc_col is None:
        ctx._error("  Process column not found in Jobtrack")
        return

    bfl_df = df[df[proc_col].astype(str).str.strip().str.upper() == 'BFL'].copy()
    ctx._log(f"  BFL rows from Jobtrack: {len(bfl_df)}")

    if bfl_df.empty:
        ctx._log("  No BFL data in Jobtrack")
        return

    # Find column names dynamically
    def find_col(target):
        for c in df.columns:
            if str(c).strip() == target:
                return c
        return None

    col_order = find_col('Order No')
    col_design = find_col('Design Name')
    col_date = find_col('Date')
    col_machine = find_col('Machine')
    col_input_name = find_col('1st Input Name')
    col_input_mm = find_col('1st Input  Size (MM)')
    col_input_mic = find_col('1st Input Mic')
    col_input_qty = find_col('1st Input  Qty')  # AY - simple additive formula
    col_balance_qty = find_col('Balance Qty')     # AZ
    col_total_input = find_col('TOTAL INPUT')     # AC - has cell refs, may be None
    col_output_kgs = find_col('Net Wt. (Kgs-Output)')  # simple additive formula
    col_output_mtrs = find_col('Output (Meters)')
    col_waste = find_col('Waste')
    col_stage = find_col('Stage')

    # Write data rows
    filled_count = 0
    bfl_cache = {}
    row_idx = DATA_START

    for _, jt_row in bfl_df.iterrows():
        order_no = _safe_str(jt_row.get(col_order, ''))
        if not order_no:
            continue

        design = _safe_str(jt_row.get(col_design, ''))
        date_val = jt_row.get(col_date, None)
        machine = _safe_str(jt_row.get(col_machine, ''))
        input_name = _safe_str(jt_row.get(col_input_name, ''))
        input_mm = _safe_float(jt_row.get(col_input_mm, 0))
        input_mic = _safe_float(jt_row.get(col_input_mic, 0))
        # Compute total_input from component columns (TOTAL INPUT has cell refs)
        total_input = _safe_float(jt_row.get(col_total_input, 0))
        if total_input == 0:
            qty = _safe_float(jt_row.get(col_input_qty, 0))
            bal = _safe_float(jt_row.get(col_balance_qty, 0))
            total_input = qty + bal
        output_kgs = _safe_float(jt_row.get(col_output_kgs, 0))
        output_mtrs = _safe_float(jt_row.get(col_output_mtrs, 0))
        waste = _safe_float(jt_row.get(col_waste, 0))

        # ── Determine Poly Rate ──
        poly_rate = 0.0
        order_upper = order_no.upper()

        # Priority 1: Granules Recipe (by WO#)
        if ctx.granules_rates and order_upper in ctx.granules_rates:
            poly_rate = ctx.granules_rates[order_upper]

        # Priority 2: Previous month Granules
        if poly_rate == 0 and ctx.prev_granules_rates and order_upper in ctx.prev_granules_rates:
            poly_rate = ctx.prev_granules_rates[order_upper]

        # Priority 3: Purchase Register (by material name + month)
        if poly_rate == 0 and input_name:
            try:
                from engine.rate_lookup import lookup_material_rate_for_month
                poly_rate = lookup_material_rate_for_month(
                    ctx.purchase_register, input_name, report_month=ctx.report_month
                )
            except Exception:
                pass

        # ── Compute derived ──
        wastage_kgs = total_input - output_kgs if total_input > 0 else waste
        wastage_val = wastage_kgs * poly_rate if poly_rate > 0 else 0
        film_value = output_kgs * poly_rate if poly_rate > 0 else 0

        # ── Write to BFL sheet ──
        ws.cell(row=row_idx, column=COL['order_no'], value=order_no)
        ws.cell(row=row_idx, column=COL['design'], value=design)
        if date_val is not None and not (isinstance(date_val, float) and pd.isna(date_val)):
            ws.cell(row=row_idx, column=COL['date'], value=date_val)
        ws.cell(row=row_idx, column=COL['machine'], value=machine)
        ws.cell(row=row_idx, column=COL['input_name'], value=input_name)
        if input_mm > 0:
            ws.cell(row=row_idx, column=COL['input_mm'], value=input_mm)
        if input_mic > 0:
            ws.cell(row=row_idx, column=COL['input_mic'], value=input_mic)
        ws.cell(row=row_idx, column=COL['total_input'], value=total_input)
        ws.cell(row=row_idx, column=COL['output_kgs'], value=output_kgs)
        ws.cell(row=row_idx, column=COL['output_mtrs'], value=output_mtrs)
        ws.cell(row=row_idx, column=COL['value'], value=film_value)
        ws.cell(row=row_idx, column=COL['poly_rate'], value=poly_rate)
        ws.cell(row=row_idx, column=COL['wastage_kgs'], value=wastage_kgs)
        ws.cell(row=row_idx, column=COL['wastage_val'], value=wastage_val)
        ws.cell(row=row_idx, column=COL['check'], value=order_no)
        ws.cell(row=row_idx, column=COL['sum_waste'], value=wastage_kgs)
        ws.cell(row=row_idx, column=COL['wastage_aed'], value=wastage_val)

        if poly_rate > 0:
            filled_count += 1

        # ── Accumulate for RMC summary ──
        if order_upper not in bfl_cache:
            bfl_cache[order_upper] = {
                'input_kgs': 0, 'output_kgs': 0, 'wastage_kgs': 0,
                'wastage_val': 0, 'poly_rate': poly_rate,
                'total_value': 0,
            }
        bfl_cache[order_upper]['input_kgs'] += total_input
        bfl_cache[order_upper]['output_kgs'] += output_kgs
        bfl_cache[order_upper]['wastage_kgs'] += wastage_kgs
        bfl_cache[order_upper]['wastage_val'] += wastage_val
        bfl_cache[order_upper]['total_value'] += film_value

        row_idx += 1

    total_rows = row_idx - DATA_START

    # ── Write subtotals (row 5) ──
    if total_rows > 0:
        last = DATA_START + total_rows - 1
        ws.cell(row=SUBTOTAL_ROW, column=COL['total_input'],
                value=f"=SUBTOTAL(9,K{DATA_START}:K{last})")
        ws.cell(row=SUBTOTAL_ROW, column=COL['output_kgs'],
                value=f"=SUBTOTAL(9,L{DATA_START}:L{last})")
        ws.cell(row=SUBTOTAL_ROW, column=COL['wastage_kgs'],
                value=f"=SUBTOTAL(9,P{DATA_START}:P{last})")
        ws.cell(row=SUBTOTAL_ROW, column=COL['wastage_val'],
                value=f"=SUBTOTAL(9,Q{DATA_START}:Q{last})")

    # Store in context
    ctx.bfl_by_order = bfl_cache

    ctx._log(f"  BFL filled: {filled_count}/{total_rows} rows with rates, "
             f"{len(bfl_cache)} orders cached")
