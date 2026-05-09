"""
Slit Filler — Fills the Slitting sheet FROM Jobtrack data.

Slit sheet columns (1-indexed):
  B=Order No(2), C=Design(3), D=Date(4), E=M/c(5), F=Material(6),
  G=Structure(7), H=Lam Pass(8), I=Input(9), J=Input Size(10),
  K=Input Kgs(11), L=Output Kgs(12), M=Output Mtrs(13),
  N=Input RMC/Kg(14), O=Input Val(15), P=Wastage Kgs(16),
  Q=Wastage Val(17), R=Wastage %(18), S=Waste LS(19), T=Output RMC/KG(20)
"""
from __future__ import annotations
import logging
import pandas as pd
from engine.base_rmc.context import RMCContext

logger = logging.getLogger(__name__)

SLIT_COLS = {
    'order_no': 2, 'design': 3, 'date': 4, 'machine': 5, 'material': 6,
    'structure': 7, 'lam_pass': 8, 'input': 9, 'input_size': 10,
    'input_kgs': 11, 'output_kgs': 12, 'output_mtrs': 13,
    'input_rmc_kg': 14, 'input_val': 15, 'wastage_kgs': 16,
    'wastage_val': 17, 'wastage_pct': 18, 'waste_ls': 19, 'output_rmc_kg': 20,
}
DATA_START = 7


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


def fill_slit(ctx: RMCContext) -> None:
    """Fill Slit sheet by writing SLITTING rows from Jobtrack."""
    ctx._log("Filling Slit sheet...")
    ws = ctx.wb['Slit'] if 'Slit' in ctx.wb.sheetnames else None
    if not ws:
        ctx._error("  Slit sheet not found!")
        return

    if ctx.jobtrack_df is None or ctx.jobtrack_df.empty:
        ctx._error("  No Jobtrack data for Slit")
        return

    df = ctx.jobtrack_df
    proc_col = None
    for c in df.columns:
        if str(c).strip().lower() == 'process':
            proc_col = c
            break

    slit_df = df[df[proc_col].astype(str).str.strip().str.upper() == 'SLITTING'].copy()
    ctx._log(f"  Slit rows from Jobtrack: {len(slit_df)}")

    if slit_df.empty:
        return

    def fc(name):
        for c in df.columns:
            if str(c).strip() == name:
                return c
        return None

    filled = 0
    slit_cache = {}
    row_idx = DATA_START

    for _, jt in slit_df.iterrows():
        order = _ss(jt.get(fc('Order No'), ''))
        if not order:
            continue

        ou = order.upper()
        design = _ss(jt.get(fc('Design Name'), ''))
        date_val = jt.get(fc('Date'), None)
        machine = _ss(jt.get(fc('Machine'), ''))
        material = _ss(jt.get(fc('Application'), ''))
        structure = _ss(jt.get(fc('Mat Structure'), ''))
        lam_pass = _ss(jt.get(fc('LAM PROCESS'), ''))
        input_name = _ss(jt.get(fc('1st Input Name'), ''))
        input_size = _sf(jt.get(fc('1st Input  Size (MM)'), 0))
        # Input kgs = 1st Input Qty + Balance Qty
        input_kgs = _sf(jt.get(fc('1st Input  Qty'), 0)) + _sf(jt.get(fc('Balance Qty'), 0))
        output_kgs = _sf(jt.get(fc('Net Wt. (Kgs-Output)'), 0))
        output_mtrs = _sf(jt.get(fc('Output (Meters)'), 0))

        # ── Determine Input RMC/Kg ──
        rate = 0.0
        lg_key = ou + "Lg"
        if lg_key in ctx.opn_wip_by_key:
            rate = ctx.opn_wip_by_key[lg_key].get('rate', 0)
        if rate == 0:
            plr_key = ou + lam_pass
            rate = ctx.pivot_lam_rates.get(plr_key, 0)
        if rate == 0 and lam_pass:
            lm_key = ou + "Lm" + lam_pass
            if lm_key in ctx.opn_wip_by_key:
                rate = ctx.opn_wip_by_key[lm_key].get('rate', 0)

        if rate > 0:
            filled += 1

        input_val = rate * input_kgs
        wastage_kgs = input_kgs - output_kgs
        wastage_val = wastage_kgs * rate
        wastage_pct = wastage_kgs / input_kgs if input_kgs > 0 else 0
        output_rmc = input_val / output_kgs if output_kgs > 0 else 0

        # Write to sheet
        ws.cell(row=row_idx, column=SLIT_COLS['order_no'], value=order)
        ws.cell(row=row_idx, column=SLIT_COLS['design'], value=design)
        if date_val is not None and not (isinstance(date_val, float) and pd.isna(date_val)):
            ws.cell(row=row_idx, column=SLIT_COLS['date'], value=date_val)
        ws.cell(row=row_idx, column=SLIT_COLS['machine'], value=machine)
        ws.cell(row=row_idx, column=SLIT_COLS['material'], value=material)
        ws.cell(row=row_idx, column=SLIT_COLS['structure'], value=structure)
        ws.cell(row=row_idx, column=SLIT_COLS['lam_pass'], value=lam_pass)
        ws.cell(row=row_idx, column=SLIT_COLS['input'], value=input_name)
        ws.cell(row=row_idx, column=SLIT_COLS['input_size'], value=input_size)
        ws.cell(row=row_idx, column=SLIT_COLS['input_kgs'], value=input_kgs)
        ws.cell(row=row_idx, column=SLIT_COLS['output_kgs'], value=output_kgs)
        ws.cell(row=row_idx, column=SLIT_COLS['output_mtrs'], value=output_mtrs)
        ws.cell(row=row_idx, column=SLIT_COLS['input_rmc_kg'], value=rate)
        ws.cell(row=row_idx, column=SLIT_COLS['input_val'], value=input_val)
        ws.cell(row=row_idx, column=SLIT_COLS['wastage_kgs'], value=wastage_kgs)
        ws.cell(row=row_idx, column=SLIT_COLS['wastage_val'], value=wastage_val)
        ws.cell(row=row_idx, column=SLIT_COLS['wastage_pct'], value=wastage_pct)
        ws.cell(row=row_idx, column=SLIT_COLS['output_rmc_kg'], value=output_rmc)

        # Cache
        ctx.slit_rate_cache[ou] = output_rmc
        if ou not in slit_cache:
            slit_cache[ou] = {'input_kgs': 0, 'output_kgs': 0, 'input_val': 0}
        slit_cache[ou]['input_kgs'] += input_kgs
        slit_cache[ou]['output_kgs'] += output_kgs
        slit_cache[ou]['input_val'] += input_val

        row_idx += 1

    total = row_idx - DATA_START
    ctx.slit_by_order = slit_cache
    ctx._log(f"  Slit filled: {filled}/{total} rows, {len(slit_cache)} orders")
