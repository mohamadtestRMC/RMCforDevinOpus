"""
Remaining Process Sheet Fillers:
  - Spout&Valve (FROM Jobtrack SPOUT & VALVE rows)
  - HCI Rew (FROM Jobtrack REWINDING rows)
  - PTR Rew (FROM Jobtrack REWINDING rows)
  - Embossing (FROM Jobtrack EMBOSSING rows)
  - Pivot_Lam Rates builder
  - FG (Finished Goods - FROM Jobtrack unique orders)
  - CLS_WIP (Closing WIP rate cascade)
  - Overall Wastage
"""
from __future__ import annotations
import logging
import pandas as pd
from engine.base_rmc.context import RMCContext

logger = logging.getLogger(__name__)

def _sf(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    try: return float(v)
    except: return 0.0

def _ss(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return str(v).strip()


# ═══════════════════════════════════════════════════════════════
# SPOUT & VALVE
# ═══════════════════════════════════════════════════════════════
def fill_spout_valve(ctx: RMCContext) -> None:
    """Fill Spout&Valve sheet FROM Jobtrack SPOUT & VALVE rows."""
    ctx._log("Filling Spout&Valve sheet...")
    ws = ctx.wb['Spout&Valve'] if 'Spout&Valve' in ctx.wb.sheetnames else None
    if not ws:
        ctx._error("  Spout&Valve sheet not found!")
        return

    SV = {'order': 2, 'design': 3, 'machine': 4, 'material': 5, 'structure': 6,
          'input': 7, 'input_kgs': 8, 'rmc_rate': 9, 'input_val': 10,
          'output_kgs': 27, 'output_rmc': 32}

    if ctx.jobtrack_df is None or ctx.jobtrack_df.empty:
        ctx._log("  No Jobtrack data for Spout&Valve")
        return

    df = ctx.jobtrack_df
    proc_col = None
    for c in df.columns:
        if str(c).strip().lower() == 'process':
            proc_col = c
            break

    sv_df = df[df[proc_col].astype(str).str.strip().str.upper() == 'SPOUT & VALVE'].copy()
    ctx._log(f"  Spout&Valve rows from Jobtrack: {len(sv_df)}")

    def fc(name):
        for c in df.columns:
            if str(c).strip() == name:
                return c
        return None

    filled = 0
    sv_cache = {}
    row_idx = 7

    for _, jt in sv_df.iterrows():
        order = _ss(jt.get(fc('Order No'), ''))
        if not order: continue
        ou = order.upper()

        input_kgs = _sf(jt.get(fc('1st Input  Qty'), 0)) + _sf(jt.get(fc('Balance Qty'), 0))
        output_kgs = _sf(jt.get(fc('Net Wt. (Kgs-Output)'), 0))

        rate = ctx.bp_rate_cache.get(ou, 0)
        if rate == 0:
            rate = ctx.slit_rate_cache.get(ou, 0)
        if rate > 0:
            filled += 1

        input_val = rate * input_kgs
        output_rmc = input_val / output_kgs if output_kgs > 0 else 0

        ws.cell(row=row_idx, column=SV['order'], value=order)
        ws.cell(row=row_idx, column=SV['design'], value=_ss(jt.get(fc('Design Name'), '')))
        ws.cell(row=row_idx, column=SV['machine'], value=_ss(jt.get(fc('Machine'), '')))
        ws.cell(row=row_idx, column=SV['material'], value=_ss(jt.get(fc('Application'), '')))
        ws.cell(row=row_idx, column=SV['input_kgs'], value=input_kgs)
        ws.cell(row=row_idx, column=SV['rmc_rate'], value=rate)
        ws.cell(row=row_idx, column=SV['input_val'], value=input_val)
        ws.cell(row=row_idx, column=SV['output_kgs'], value=output_kgs)
        ws.cell(row=row_idx, column=SV['output_rmc'], value=output_rmc)

        if ou not in sv_cache:
            sv_cache[ou] = {'input_kgs': 0, 'output_kgs': 0, 'total_val': 0}
        sv_cache[ou]['input_kgs'] += input_kgs
        sv_cache[ou]['output_kgs'] += output_kgs
        sv_cache[ou]['total_val'] += input_val
        row_idx += 1

    total = row_idx - 7
    ctx.spout_valve_by_order = sv_cache
    ctx._log(f"  Spout&Valve filled: {filled}/{total} rows")


# ═══════════════════════════════════════════════════════════════
# HCI REW (Rewinder) — FROM Jobtrack REWINDING rows
# ═══════════════════════════════════════════════════════════════
def fill_hci_rew(ctx: RMCContext) -> None:
    """Fill HCI Rew sheet FROM Jobtrack REWINDING rows."""
    ctx._log("Filling HCI Rew sheet...")
    ws = ctx.wb['HCI Rew'] if 'HCI Rew' in ctx.wb.sheetnames else None
    if not ws:
        ctx._error("  HCI Rew sheet not found!")
        return

    HC = {'order': 2, 'design': 3, 'material': 4, 'structure': 5,
          'input': 6, 'input_kgs': 7, 'output_kgs': 8, 'wastage': 9,
          'rate': 10, 'wastage_val': 11}

    if ctx.jobtrack_df is None or ctx.jobtrack_df.empty:
        ctx._log("  No Jobtrack data for HCI Rew")
        return

    df = ctx.jobtrack_df
    proc_col = None
    for c in df.columns:
        if str(c).strip().lower() == 'process':
            proc_col = c
            break

    rew_df = df[df[proc_col].astype(str).str.strip().str.upper() == 'REWINDING'].copy()
    ctx._log(f"  HCI Rew rows from Jobtrack: {len(rew_df)}")

    def fc(name):
        for c in df.columns:
            if str(c).strip() == name:
                return c
        return None

    filled = 0
    hci_cache = {}
    row_idx = 7

    for _, jt in rew_df.iterrows():
        order = _ss(jt.get(fc('Order No'), ''))
        if not order: continue
        ou = order.upper()

        input_kgs = _sf(jt.get(fc('1st Input  Qty'), 0)) + _sf(jt.get(fc('Balance Qty'), 0))
        output_kgs = _sf(jt.get(fc('Net Wt. (Kgs-Output)'), 0))
        wastage = input_kgs - output_kgs

        rate = ctx.slit_rate_cache.get(ou, 0)
        if rate == 0:
            for suffix in ['Lg', 'Wh', 'pg']:
                key = ou + suffix
                if key in ctx.opn_wip_by_key:
                    rate = ctx.opn_wip_by_key[key].get('rate', 0)
                    if rate > 0: break

        if rate > 0:
            filled += 1

        wastage_val = wastage * rate

        ws.cell(row=row_idx, column=HC['order'], value=order)
        ws.cell(row=row_idx, column=HC['design'], value=_ss(jt.get(fc('Design Name'), '')))
        ws.cell(row=row_idx, column=HC['material'], value=_ss(jt.get(fc('Application'), '')))
        ws.cell(row=row_idx, column=HC['structure'], value=_ss(jt.get(fc('Mat Structure'), '')))
        ws.cell(row=row_idx, column=HC['input_kgs'], value=input_kgs)
        ws.cell(row=row_idx, column=HC['output_kgs'], value=output_kgs)
        ws.cell(row=row_idx, column=HC['wastage'], value=wastage)
        ws.cell(row=row_idx, column=HC['rate'], value=rate)
        ws.cell(row=row_idx, column=HC['wastage_val'], value=wastage_val)

        if ou not in hci_cache:
            hci_cache[ou] = {'wastage': 0, 'wastage_val': 0}
        hci_cache[ou]['wastage'] += wastage
        hci_cache[ou]['wastage_val'] += wastage_val
        row_idx += 1

    total = row_idx - 7
    ctx.hci_rew_by_order = hci_cache
    ctx._log(f"  HCI Rew filled: {filled}/{total} rows")


# ═══════════════════════════════════════════════════════════════
# PTR REW (Rewinder) — FROM Jobtrack REWINDING rows
# ═══════════════════════════════════════════════════════════════
def fill_ptr_rew(ctx: RMCContext) -> None:
    """Fill PTR Rew sheet FROM Jobtrack REWINDING rows.
    PTR Rew tracks Print and Lam wastage separately.
    """
    ctx._log("Filling PTR Rew sheet...")
    ws = ctx.wb['PTR Rew'] if 'PTR Rew' in ctx.wb.sheetnames else None
    if not ws:
        ctx._error("  PTR Rew sheet not found!")
        return

    PT = {'order': 2, 'design': 3, 'date': 4, 'material': 5, 'structure': 6,
          'input': 7, 'input_size': 8, 'input_mic': 9,
          'print_input': 10, 'print_output': 11, 'print_mtrs': 12,
          'print_wastage': 13, 'print_rate': 14, 'print_val': 15,
          'lam_input': 17, 'lam_output': 18, 'lam_mtrs': 19,
          'lam_wastage': 20, 'lam_rate': 21, 'lam_val': 22}

    if ctx.jobtrack_df is None or ctx.jobtrack_df.empty:
        ctx._log("  No Jobtrack data for PTR Rew")
        return

    df = ctx.jobtrack_df
    proc_col = None
    for c in df.columns:
        if str(c).strip().lower() == 'process':
            proc_col = c
            break

    rew_df = df[df[proc_col].astype(str).str.strip().str.upper() == 'REWINDING'].copy()

    def fc(name):
        for c in df.columns:
            if str(c).strip() == name:
                return c
        return None

    filled = 0
    ptr_cache = {}
    row_idx = 7

    for _, jt in rew_df.iterrows():
        order = _ss(jt.get(fc('Order No'), ''))
        if not order: continue
        ou = order.upper()

        input_kgs = _sf(jt.get(fc('1st Input  Qty'), 0)) + _sf(jt.get(fc('Balance Qty'), 0))
        output_kgs = _sf(jt.get(fc('Net Wt. (Kgs-Output)'), 0))
        wastage = input_kgs - output_kgs

        # Print rate
        print_rate = ctx.print_rate_cache.get(ou, 0)
        if print_rate > 0:
            ws.cell(row=row_idx, column=PT['print_rate'], value=print_rate)
            ws.cell(row=row_idx, column=PT['print_val'], value=wastage * print_rate)

        # Lam rate
        lam_rate = 0
        for lp in ['L1', 'L2', 'L3']:
            lam_rate = ctx.pivot_lam_rates.get(ou + lp, 0)
            if lam_rate > 0: break
        if lam_rate > 0:
            ws.cell(row=row_idx, column=PT['lam_rate'], value=lam_rate)
            ws.cell(row=row_idx, column=PT['lam_val'], value=wastage * lam_rate)
            filled += 1

        ws.cell(row=row_idx, column=PT['order'], value=order)
        ws.cell(row=row_idx, column=PT['design'], value=_ss(jt.get(fc('Design Name'), '')))
        ws.cell(row=row_idx, column=PT['material'], value=_ss(jt.get(fc('Application'), '')))
        ws.cell(row=row_idx, column=PT['structure'], value=_ss(jt.get(fc('Mat Structure'), '')))
        ws.cell(row=row_idx, column=PT['print_input'], value=input_kgs)
        ws.cell(row=row_idx, column=PT['print_output'], value=output_kgs)
        ws.cell(row=row_idx, column=PT['print_wastage'], value=wastage)

        if ou not in ptr_cache:
            ptr_cache[ou] = {'print_waste_val': 0, 'lam_waste_val': 0}
        ptr_cache[ou]['print_waste_val'] += wastage * print_rate
        ptr_cache[ou]['lam_waste_val'] += wastage * lam_rate
        row_idx += 1

    total = row_idx - 7
    ctx.ptr_rew_by_order = ptr_cache
    ctx._log(f"  PTR Rew filled: {filled}/{total} rows")


# ═══════════════════════════════════════════════════════════════
# EMBOSSING — FROM Jobtrack EMBOSSING rows
# ═══════════════════════════════════════════════════════════════
def fill_embossing(ctx: RMCContext) -> None:
    """Fill Embossing sheet FROM Jobtrack EMBOSSING rows."""
    ctx._log("Filling Embossing sheet...")
    ws = ctx.wb['Embossing'] if 'Embossing' in ctx.wb.sheetnames else None
    if not ws:
        ctx._log("  Embossing sheet not found")
        return

    EM = {'order': 2, 'design': 3, 'date': 4, 'machine': 5, 'material': 6,
          'structure': 7, 'lam_pass': 8, 'input': 9, 'input_size': 10,
          'input_kgs': 11, 'output_kgs': 12, 'output_mtrs': 13,
          'input_rmc': 14, 'input_val': 15,
          'wastage_kgs': 16, 'wastage_val': 17, 'output_rmc': 20}

    if ctx.jobtrack_df is None or ctx.jobtrack_df.empty:
        ctx._log("  No Jobtrack data for Embossing")
        return

    df = ctx.jobtrack_df
    proc_col = None
    for c in df.columns:
        if str(c).strip().lower() == 'process':
            proc_col = c
            break

    emb_df = df[df[proc_col].astype(str).str.strip().str.upper() == 'EMBOSSING'].copy()
    ctx._log(f"  Embossing rows from Jobtrack: {len(emb_df)}")

    def fc(name):
        for c in df.columns:
            if str(c).strip() == name:
                return c
        return None

    filled = 0
    em_cache = {}
    row_idx = 7

    for _, jt in emb_df.iterrows():
        order = _ss(jt.get(fc('Order No'), ''))
        if not order: continue
        ou = order.upper()

        input_kgs = _sf(jt.get(fc('1st Input  Qty'), 0)) + _sf(jt.get(fc('Balance Qty'), 0))
        output_kgs = _sf(jt.get(fc('Net Wt. (Kgs-Output)'), 0))

        rate = ctx.slit_rate_cache.get(ou, 0)
        if rate == 0:
            for lp in ['L1', 'L2', 'L3']:
                rate = ctx.pivot_lam_rates.get(ou + lp, 0)
                if rate > 0: break

        if rate > 0:
            filled += 1

        input_val = rate * input_kgs
        wastage_kgs = input_kgs - output_kgs
        wastage_val = wastage_kgs * rate
        output_rmc = input_val / output_kgs if output_kgs > 0 else 0

        ws.cell(row=row_idx, column=EM['order'], value=order)
        ws.cell(row=row_idx, column=EM['design'], value=_ss(jt.get(fc('Design Name'), '')))
        ws.cell(row=row_idx, column=EM['material'], value=_ss(jt.get(fc('Application'), '')))
        ws.cell(row=row_idx, column=EM['structure'], value=_ss(jt.get(fc('Mat Structure'), '')))
        ws.cell(row=row_idx, column=EM['input_kgs'], value=input_kgs)
        ws.cell(row=row_idx, column=EM['output_kgs'], value=output_kgs)
        ws.cell(row=row_idx, column=EM['input_rmc'], value=rate)
        ws.cell(row=row_idx, column=EM['input_val'], value=input_val)
        ws.cell(row=row_idx, column=EM['wastage_kgs'], value=wastage_kgs)
        ws.cell(row=row_idx, column=EM['wastage_val'], value=wastage_val)
        ws.cell(row=row_idx, column=EM['output_rmc'], value=output_rmc)

        if ou not in em_cache:
            em_cache[ou] = {'wastage_kgs': 0, 'wastage_val': 0}
        em_cache[ou]['wastage_kgs'] += wastage_kgs
        em_cache[ou]['wastage_val'] += wastage_val
        row_idx += 1

    total = row_idx - 7
    ctx.embossing_by_order = em_cache
    ctx._log(f"  Embossing filled: {filled}/{total} rows")


# ═══════════════════════════════════════════════════════════════
# PIVOT_LAM RATES
# ═══════════════════════════════════════════════════════════════
def build_pivot_lam_rates(ctx: RMCContext) -> None:
    """Build the Pivot_Lam Rates sheet from Lam cache."""
    ctx._log("Building Pivot_Lam Rates sheet...")
    ws = ctx.wb['Pivot_Lam Rates'] if 'Pivot_Lam Rates' in ctx.wb.sheetnames else None
    if not ws:
        ctx._log("  Pivot_Lam Rates sheet not found")
        return

    ws.cell(row=3, column=1, value='Order No')
    ws.cell(row=3, column=2, value='Lam Process')
    ws.cell(row=3, column=3, value='Sum of Output\nKgs')
    ws.cell(row=3, column=4, value='Sum of Total Input Val.')
    ws.cell(row=3, column=6, value='Avg Rate')

    row = 4
    for key, rate in sorted(ctx.pivot_lam_rates.items()):
        if len(key) > 2 and key[-2] == 'L' and key[-1] in '123':
            order = key[:-2]
            lam_pass = key[-2:]
        else:
            order = key
            lam_pass = ""

        lam_data = ctx.lam_by_order.get(order, {})
        output_kgs = lam_data.get('output_kgs', 0)
        total_val = lam_data.get('total_input_val', 0)

        ws.cell(row=row, column=1, value=order)
        ws.cell(row=row, column=2, value=lam_pass)
        ws.cell(row=row, column=3, value=output_kgs)
        ws.cell(row=row, column=4, value=total_val)
        ws.cell(row=row, column=6, value=rate)
        ws.cell(row=row, column=8, value=order + lam_pass)
        ws.cell(row=row, column=9, value=rate)
        row += 1

    ctx._log(f"  Pivot_Lam Rates: {row - 4} entries written")


# ═══════════════════════════════════════════════════════════════
# FG (Finished Goods)
# ═══════════════════════════════════════════════════════════════
def fill_fg(ctx: RMCContext) -> None:
    """Fill FG sheet from Jobtrack unique orders."""
    ctx._log("Filling FG sheet...")
    ws = ctx.wb['FG'] if 'FG' in ctx.wb.sheetnames else None
    if not ws:
        ctx._error("  FG sheet not found!")
        return

    FG = {'order': 1, 'output_kgs': 2, 'hci_waste': 6, 'final_fg': 7}

    if ctx.jobtrack_df is None or ctx.jobtrack_df.empty:
        ctx._log("  No Jobtrack data for FG")
        return

    df = ctx.jobtrack_df

    def fc(name):
        for c in df.columns:
            if str(c).strip() == name:
                return c
        return None

    # Aggregate output kgs by order
    order_col = fc('Order No')
    output_col = fc('Net Wt. (Kgs-Output)')
    if not order_col:
        ctx._error("  Order No column not found")
        return

    # Group by order, sum output kgs
    orders = {}
    for _, row in df.iterrows():
        o = _ss(row.get(order_col, ''))
        if not o:
            continue
        ou = o.upper()
        if ou not in orders:
            orders[ou] = {'order': o, 'output_kgs': 0}
        orders[ou]['output_kgs'] += _sf(row.get(output_col, 0))

    row_idx = 4
    fg_cache = {}
    for ou, data in sorted(orders.items()):
        output_kgs = data['output_kgs']
        hci_waste = 0
        if ou in getattr(ctx, 'hci_rew_by_order', {}):
            hci_waste = ctx.hci_rew_by_order[ou].get('wastage', 0)

        final_fg = output_kgs - hci_waste

        ws.cell(row=row_idx, column=FG['order'], value=data['order'])
        ws.cell(row=row_idx, column=FG['output_kgs'], value=output_kgs)
        ws.cell(row=row_idx, column=FG['hci_waste'], value=hci_waste)
        ws.cell(row=row_idx, column=FG['final_fg'], value=final_fg)

        fg_cache[ou] = {'output_kgs': output_kgs, 'hci_waste': hci_waste, 'final_fg': final_fg}
        row_idx += 1

    ctx.fg_by_order = fg_cache
    ctx._log(f"  FG filled: {len(fg_cache)} rows, {len(fg_cache)} orders")


# ═══════════════════════════════════════════════════════════════
# CLS_WIP (Closing WIP Rate Cascade)
# ═══════════════════════════════════════════════════════════════
def fill_cls_wip(ctx: RMCContext) -> None:
    """Fill CLS_WIP sheet with computed rates."""
    ctx._log("Filling CLS_WIP sheet...")
    ws = ctx.wb['CLS_WIP'] if 'CLS_WIP' in ctx.wb.sheetnames else None
    if not ws:
        ctx._error("  CLS_WIP sheet not found!")
        return

    from engine.base_rmc.wip_keys import compute_wip_composite_key

    CW = {'key': 1, 'wo': 2, 'design': 3, 'mat': 4, 'process': 5,
          'substrate': 6, 'lam_pass': 7, 'qty': 8, 'rate': 9, 'value': 10}

    filled = 0
    total = 0
    cls_cache = {}

    for r in range(6, ws.max_row + 1):
        order = _ss(ws.cell(row=r, column=CW['wo']).value)
        if not order: continue
        total += 1
        ou = order.upper()

        process = _ss(ws.cell(row=r, column=CW['process']).value)
        lam_pass = _ss(ws.cell(row=r, column=CW['lam_pass']).value)
        qty = _sf(ws.cell(row=r, column=CW['qty']).value)

        key = compute_wip_composite_key(ou, process, lam_pass)
        ws.cell(row=r, column=CW['key'], value=key)

        rate = 0.0

        if key.endswith('Pm'):
            rate = ctx.print_rate_cache.get(ou, 0)
            if rate == 0 and key in ctx.opn_wip_by_key:
                rate = ctx.opn_wip_by_key[key].get('rate', 0)

        elif key.endswith('Lg') or ('Lm' in key and key[-2] == 'L'):
            if lam_pass:
                rate = ctx.pivot_lam_rates.get(ou + lam_pass, 0)
            if rate == 0:
                for lp in ['L1', 'L2', 'L3']:
                    rate = ctx.pivot_lam_rates.get(ou + lp, 0)
                    if rate > 0: break
            if rate == 0 and key in ctx.opn_wip_by_key:
                rate = ctx.opn_wip_by_key[key].get('rate', 0)

        elif key.endswith('Wh') or key.endswith('pg'):
            rate = ctx.bp_rate_cache.get(ou, 0)
            if rate == 0:
                rate = ctx.slit_rate_cache.get(ou, 0)
            if rate == 0 and key in ctx.opn_wip_by_key:
                rate = ctx.opn_wip_by_key[key].get('rate', 0)

        else:
            if key in ctx.opn_wip_by_key:
                rate = ctx.opn_wip_by_key[key].get('rate', 0)

        if rate > 0:
            ws.cell(row=r, column=CW['rate'], value=rate)
            ws.cell(row=r, column=CW['value'], value=rate * qty)
            filled += 1

        cls_cache[key] = {'order': ou, 'qty': qty, 'rate': rate, 'value': rate * qty}

    ctx.cls_wip_by_key = cls_cache
    ctx._log(f"  CLS_WIP filled: {filled}/{total} rows, {len(cls_cache)} keys")
