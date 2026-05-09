"""
Bag&Pouch Filler — Fills the Bag&Pouch sheet FROM Jobtrack data.
Processes: POUCHING, BAG from Jobtrack.
"""
from __future__ import annotations
import logging
import pandas as pd
from engine.base_rmc.context import RMCContext

logger = logging.getLogger(__name__)

BP_COLS = {
    'order_no': 2, 'design': 3, 'machine': 4, 'material': 5,
    'structure': 6, 'input_lam_pass': 7, 'bag_type': 8,
    'input_size': 9, 'input_mic': 10, 'input_kgs2': 11,
    'zipper_rate': 12, 'input_kgs': 13, 'output_kgs': 14,
    'output_pcs': 15, 'wastage_kgs': 16, 'rmc_rate': 17,
    'input_val': 18, 'wastage_val': 19,
    'pe_strip_kgs': 20, 'pe_strip_rate': 21, 'pe_strip_val': 22,
    'zipper_kgs': 23, 'zipper_val': 24,
    'total_input_val': 25, 'output_kgs2': 26, 'output_rmc_kg': 27,
}
DATA_START = 7
PE_STRIP_RATE = 11.85


def _sf(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    try:
        return float(v)
    except:
        return 0.0


def _ss(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


def fill_bag_pouch(ctx: RMCContext) -> None:
    """Fill Bag&Pouch sheet by writing POUCHING/BAG rows from Jobtrack."""
    ctx._log("Filling Bag&Pouch sheet...")
    ws = ctx.wb['Bag&Pouch'] if 'Bag&Pouch' in ctx.wb.sheetnames else None
    if not ws:
        ctx._error("  Bag&Pouch sheet not found!")
        return

    if ctx.jobtrack_df is None or ctx.jobtrack_df.empty:
        ctx._error("  No Jobtrack data for Bag&Pouch")
        return

    df = ctx.jobtrack_df
    proc_col = None
    for c in df.columns:
        if str(c).strip().lower() == 'process':
            proc_col = c
            break

    bp_df = df[df[proc_col].astype(str).str.strip().str.upper().isin(['POUCHING', 'BAG'])].copy()
    ctx._log(f"  Bag&Pouch rows from Jobtrack: {len(bp_df)}")

    if bp_df.empty:
        return

    def fc(name):
        for c in df.columns:
            if str(c).strip() == name:
                return c
        return None

    filled = 0
    bp_cache = {}
    row_idx = DATA_START

    for _, jt in bp_df.iterrows():
        order = _ss(jt.get(fc('Order No'), ''))
        if not order:
            continue

        ou = order.upper()
        design = _ss(jt.get(fc('Design Name'), ''))
        machine = _ss(jt.get(fc('Machine'), ''))
        material = _ss(jt.get(fc('Application'), ''))
        structure = _ss(jt.get(fc('Mat Structure'), ''))
        lam_pass = _ss(jt.get(fc('LAM PROCESS'), ''))
        input_kgs = _sf(jt.get(fc('1st Input  Qty'), 0)) + _sf(jt.get(fc('Balance Qty'), 0))
        output_kgs = _sf(jt.get(fc('Net Wt. (Kgs-Output)'), 0))

        # RMC Rate lookup
        rate = 0.0
        wh_key = ou + "Wh"
        pg_key = ou + "pg"
        if wh_key in ctx.opn_wip_by_key:
            rate = ctx.opn_wip_by_key[wh_key].get('rate', 0)
        elif pg_key in ctx.opn_wip_by_key:
            rate = ctx.opn_wip_by_key[pg_key].get('rate', 0)
        if rate == 0:
            rate = ctx.slit_rate_cache.get(ou, 0)
        if rate == 0 and lam_pass:
            rate = ctx.pivot_lam_rates.get(ou + lam_pass, 0)

        if rate > 0:
            filled += 1

        input_val = rate * input_kgs
        wastage_kgs = input_kgs - output_kgs
        wastage_val = wastage_kgs * rate
        output_rmc = input_val / output_kgs if output_kgs > 0 else 0

        # Write
        ws.cell(row=row_idx, column=BP_COLS['order_no'], value=order)
        ws.cell(row=row_idx, column=BP_COLS['design'], value=design)
        ws.cell(row=row_idx, column=BP_COLS['machine'], value=machine)
        ws.cell(row=row_idx, column=BP_COLS['material'], value=material)
        ws.cell(row=row_idx, column=BP_COLS['structure'], value=structure)
        ws.cell(row=row_idx, column=BP_COLS['input_lam_pass'], value=lam_pass)
        ws.cell(row=row_idx, column=BP_COLS['input_kgs'], value=input_kgs)
        ws.cell(row=row_idx, column=BP_COLS['output_kgs'], value=output_kgs)
        ws.cell(row=row_idx, column=BP_COLS['wastage_kgs'], value=wastage_kgs)
        ws.cell(row=row_idx, column=BP_COLS['rmc_rate'], value=rate)
        ws.cell(row=row_idx, column=BP_COLS['input_val'], value=input_val)
        ws.cell(row=row_idx, column=BP_COLS['wastage_val'], value=wastage_val)
        ws.cell(row=row_idx, column=BP_COLS['pe_strip_rate'], value=PE_STRIP_RATE)
        ws.cell(row=row_idx, column=BP_COLS['total_input_val'], value=input_val)
        ws.cell(row=row_idx, column=BP_COLS['output_kgs2'], value=output_kgs)
        ws.cell(row=row_idx, column=BP_COLS['output_rmc_kg'], value=output_rmc)

        ctx.bp_rate_cache[ou] = output_rmc
        if ou not in bp_cache:
            bp_cache[ou] = {'input_kgs': 0, 'output_kgs': 0, 'total_val': 0}
        bp_cache[ou]['input_kgs'] += input_kgs
        bp_cache[ou]['output_kgs'] += output_kgs
        bp_cache[ou]['total_val'] += input_val

        row_idx += 1

    total = row_idx - DATA_START
    ctx.bag_pouch_by_order = bp_cache
    ctx._log(f"  Bag&Pouch filled: {filled}/{total} rows, {len(bp_cache)} orders")
