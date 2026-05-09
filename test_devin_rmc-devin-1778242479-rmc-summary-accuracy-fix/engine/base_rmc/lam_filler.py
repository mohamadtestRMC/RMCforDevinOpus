"""
Lam Filler — Fills the Lamination sheet FROM Jobtrack data.

Lam sheet columns (from filled reference, 1-indexed):
  B=Order No, C=Design, D=Date, E=M/c, F=Material, G=Structure,
  H=Lam Process, I=Next Dept, J=Sleeve size,
  K=Ptd Mat, L=Size, M=Mic, N=Lam Mat, O=Size, P=Mic,
  Q=Fresh Mat, R=Size, S=Mic, T=Fresh Mat2, U=Size, V=Mic,
  W=ADH NAME, X=Hard Name, Y=Adh GSM,
  Z=Ptd Mat Qty(26), AA=Rate(27), AB=Ptd Mat Value(28),
  AC=Lam Mat Qty(29), AD=Rate(30), AE=Lam Input Value(31),
  AF=1st Fresh Qty(32), AG=Rate(33), AH=1st Fresh Value(34),
  AI=2nd Fresh Qty(35), AJ=Rate(36), AK=2nd Fresh Value(37),
  AL=Adh Qty(38), AM=Adh Solids(39), AN=Adh rate(40), AO=Adh Value(41),
  AP=Hard Qty(42), AQ=Hard Solids(43), AR=Hard Rate(44), AS=Hard Value(45),
  AT=Adh+Hard Solids(46), AU=Adh+Hard Calc(47),
  AV=Solv Qty(48), AW=Sol Rate(49), AX=Solv Value(50),
  AY=Fresh Mat Qty(51), AZ=Fresh Mat Value(52),
  BA=Adh+Hard Solids Qty(53), BB=Adh+Hard+Solv Val(54),
  BC=Total Input Qty(55), BD=Total Input Val(56),
  BE=Output Kgs(57), BF=Per Kg RMC(58), BG=Output Mtrs(59), BH=Prod Sq Mtr(60)
"""
from __future__ import annotations
import logging
import pandas as pd
from engine.base_rmc.context import RMCContext

logger = logging.getLogger(__name__)

# Lam sheet columns (1-indexed) matching filled reference
C = {
    'order_no': 2, 'design': 3, 'date': 4, 'machine': 5,
    'material': 6, 'structure': 7, 'lam_process': 8, 'next_dept': 9,
    'sleeve_size': 10,
    'ptd_mat': 11, 'ptd_size': 12, 'ptd_mic': 13,
    'lam_mat': 14, 'lam_size': 15, 'lam_mic': 16,
    'fresh1_mat': 17, 'fresh1_size': 18, 'fresh1_mic': 19,
    'fresh2_mat': 20, 'fresh2_size': 21, 'fresh2_mic': 22,
    'adh_name': 23, 'hard_name': 24, 'adh_gsm': 25,
    'ptd_qty': 26, 'ptd_rate': 27, 'ptd_value': 28,
    'lam_qty': 29, 'lam_rate': 30, 'lam_value': 31,
    'fresh1_qty': 32, 'fresh1_rate': 33, 'fresh1_value': 34,
    'fresh2_qty': 35, 'fresh2_rate': 36, 'fresh2_value': 37,
    'adh_qty': 38, 'adh_solids': 39, 'adh_rate': 40, 'adh_value': 41,
    'hard_qty': 42, 'hard_solids': 43, 'hard_rate': 44, 'hard_value': 45,
    'adh_hard_solids': 46, 'adh_hard_calc': 47,
    'solv_qty': 48, 'solv_rate': 49, 'solv_value': 50,
    'fresh_mat_qty': 51, 'fresh_mat_value': 52,
    'adh_hard_solids_qty': 53, 'adh_hard_solv_val': 54,
    'total_input_qty': 55, 'total_input_val': 56,
    'output_kgs': 57, 'per_kg_rmc': 58, 'output_mtrs': 59, 'prod_sqm': 60,
}

DATA_START = 7
SUBTOTAL_ROW = 5


def _sf(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def _ss(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return str(val).strip()


def fill_lam(ctx: RMCContext) -> None:
    """Fill the Lam sheet by writing LAM rows from enriched Jobtrack."""
    ctx._log("Filling Lam sheet...")

    ws = ctx.wb['Lam'] if 'Lam' in ctx.wb.sheetnames else None
    if ws is None:
        ctx._error("  Lam sheet not found!")
        return

    if ctx.jobtrack_df is None or ctx.jobtrack_df.empty:
        ctx._error("  No Jobtrack data for Lam")
        return

    df = ctx.jobtrack_df
    proc_col = None
    for c in df.columns:
        if str(c).strip().lower() == 'process':
            proc_col = c
            break

    lam_df = df[df[proc_col].astype(str).str.strip().str.upper() == 'LAM'].copy()
    ctx._log(f"  Lam rows from Jobtrack: {len(lam_df)}")

    if lam_df.empty:
        return

    def fc(name):
        for c in df.columns:
            if str(c).strip() == name:
                return c
        return None

    # Chemical rates
    solvent_rate = ctx.solvent_rate

    filled_ptd = 0
    filled_adh = 0
    filled_hard = 0
    filled_solv = 0
    lam_cache = {}
    order_cache = {}
    row_idx = DATA_START

    for _, jt in lam_df.iterrows():
        order_no = _ss(jt.get(fc('Order No'), ''))
        if not order_no:
            continue

        order_upper = order_no.upper()
        design = _ss(jt.get(fc('Design Name'), ''))
        date_val = jt.get(fc('Date'), None)
        machine = _ss(jt.get(fc('Machine'), ''))
        material = _ss(jt.get(fc('Application'), ''))
        lam_process = _ss(jt.get(fc('LAM PROCESS'), ''))
        next_dept = _ss(jt.get(fc('Next Dept.'), ''))
        sleeve_size = _sf(jt.get(fc('sleeve size'), 0))

        # Material inputs from Jobtrack
        ptd_mat = _ss(jt.get(fc('1st Input Ptd Mat'), ''))
        ptd_size = _sf(jt.get(fc('1st Input  Size (MM).1'), 0))
        ptd_mic = _sf(jt.get(fc('1st Input Mic.1'), 0))
        lam_mat = _ss(jt.get(fc('Laminated-Mat.'), ''))
        lam_size = _sf(jt.get(fc('Size'), 0))
        lam_mic = _sf(jt.get(fc('Mic.'), 0))
        fresh1_mat = _ss(jt.get(fc('1st Fresh Material'), ''))
        fresh1_size = _sf(jt.get(fc('Size.1'), 0))
        fresh1_mic = _sf(jt.get(fc('Mic..1'), 0))
        fresh2_mat = _ss(jt.get(fc('2nd Fresh Material'), ''))
        fresh2_size = _sf(jt.get(fc('Size.2'), 0))
        fresh2_mic = _sf(jt.get(fc('Mic..2'), 0))
        adh_name = _ss(jt.get(fc('ADH NAME'), ''))
        hard_name = _ss(jt.get(fc('HARDNER'), ''))

        # Quantities (simple additive formulas are evaluatable)
        ptd_qty = _sf(jt.get(fc('Total 1st Ptd-Mat Input Qty'), 0))
        if ptd_qty == 0:
            ptd_qty = _sf(jt.get(fc('1st Input  Ptd Qty'), 0)) + _sf(jt.get(fc('Balance Qty.1'), 0))
        lam_qty = _sf(jt.get(fc('Total Lam-Input Qty'), 0))
        if lam_qty == 0:
            lam_qty = _sf(jt.get(fc('Lam-Mat. Qty'), 0)) + _sf(jt.get(fc('Balance Qty.2'), 0))
        fresh1_qty = _sf(jt.get(fc('Total 1st Fresh Material Qty'), 0))
        if fresh1_qty == 0:
            fresh1_qty = _sf(jt.get(fc('1st Fresh Material Qty'), 0)) + _sf(jt.get(fc('Balance Qty.3'), 0))
        fresh2_qty = _sf(jt.get(fc('Total 2nd Fresh Material Qty'), 0))
        if fresh2_qty == 0:
            fresh2_qty = _sf(jt.get(fc('2nd Fresh Material Qty'), 0)) + _sf(jt.get(fc('Balance Qty.4'), 0))
        adh_qty = _sf(jt.get(fc('ADH KGS'), 0))
        adh_solids = _sf(jt.get(fc('Adh Solids'), 0))
        hard_qty = _sf(jt.get(fc('HARDNER KG'), 0))
        hard_solids = _sf(jt.get(fc('Hard Solids'), 0))
        output_kgs = _sf(jt.get(fc('Net Wt. (Kgs-Output)'), 0))
        output_mtrs = _sf(jt.get(fc('Output (Meters)'), 0))

        # ── Rates from enriched Jobtrack ──
        # Ptd rate: from Print rate cache or OPN_WIP
        ptd_rate = 0.0
        pm_key = order_upper + "Pm"
        if pm_key in ctx.opn_wip_by_key:
            ptd_rate = ctx.opn_wip_by_key[pm_key].get('rate', 0)
        elif order_upper in ctx.print_rate_cache:
            ptd_rate = ctx.print_rate_cache[order_upper]
        if ptd_rate > 0:
            filled_ptd += 1

        # Fresh1 rate from Jobtrack enrichment
        fresh1_rate = _sf(jt.get(fc('Rate.1'), 0))
        # Fresh2 rate from Jobtrack enrichment
        fresh2_rate = _sf(jt.get(fc('Rate.2'), 0))

        # Adhesive rate from Jobtrack enrichment
        adh_rate = _sf(jt.get(fc('Rate.3'), 0))
        if adh_rate > 0:
            filled_adh += 1

        # Hardener rate from Jobtrack enrichment
        hard_rate = _sf(jt.get(fc('Rate.4'), 0))
        if hard_rate > 0:
            filled_hard += 1

        # Solvent
        solv_qty = _sf(jt.get(fc('LAM SOL (E/A)'), 0))
        solv_rate = _sf(jt.get(fc('Rate.5'), 0))
        if solv_rate == 0 and solvent_rate > 0:
            solv_rate = solvent_rate
        if solv_qty > 0 and solv_rate > 0:
            filled_solv += 1

        # ── Compute values ──
        ptd_value = ptd_qty * ptd_rate
        lam_value = 0  # Lam mat usually has 0 qty for most rows
        fresh1_value = fresh1_qty * fresh1_rate
        fresh2_value = fresh2_qty * fresh2_rate
        adh_value = adh_qty * adh_rate
        hard_value = hard_qty * hard_rate
        solv_value = solv_qty * solv_rate

        fresh_mat_qty = fresh1_qty + fresh2_qty
        fresh_mat_value = fresh1_value + fresh2_value
        adh_hard_solids = adh_solids + hard_solids
        adh_hard_solv_val = adh_value + hard_value + solv_value
        total_input_qty = ptd_qty + lam_qty + fresh1_qty + fresh2_qty + adh_solids + hard_solids + solv_qty
        total_input_val = ptd_value + lam_value + fresh1_value + fresh2_value + adh_value + hard_value + solv_value
        per_kg_rmc = total_input_val / output_kgs if output_kgs > 0 else 0

        # ── Write to Lam sheet ──
        ws.cell(row=row_idx, column=C['order_no'], value=order_no)
        ws.cell(row=row_idx, column=C['design'], value=design)
        if date_val is not None and not (isinstance(date_val, float) and pd.isna(date_val)):
            ws.cell(row=row_idx, column=C['date'], value=date_val)
        ws.cell(row=row_idx, column=C['machine'], value=machine)
        ws.cell(row=row_idx, column=C['material'], value=material)
        ws.cell(row=row_idx, column=C['lam_process'], value=lam_process)
        ws.cell(row=row_idx, column=C['next_dept'], value=next_dept)
        if sleeve_size > 0:
            ws.cell(row=row_idx, column=C['sleeve_size'], value=sleeve_size)

        # Material names
        ws.cell(row=row_idx, column=C['ptd_mat'], value=ptd_mat)
        ws.cell(row=row_idx, column=C['fresh1_mat'], value=fresh1_mat)
        ws.cell(row=row_idx, column=C['fresh2_mat'], value=fresh2_mat)
        ws.cell(row=row_idx, column=C['adh_name'], value=adh_name)
        ws.cell(row=row_idx, column=C['hard_name'], value=hard_name)

        # Quantities, rates, values
        ws.cell(row=row_idx, column=C['ptd_qty'], value=ptd_qty)
        ws.cell(row=row_idx, column=C['ptd_rate'], value=ptd_rate)
        ws.cell(row=row_idx, column=C['ptd_value'], value=ptd_value)
        ws.cell(row=row_idx, column=C['lam_qty'], value=lam_qty)
        ws.cell(row=row_idx, column=C['fresh1_qty'], value=fresh1_qty)
        ws.cell(row=row_idx, column=C['fresh1_rate'], value=fresh1_rate)
        ws.cell(row=row_idx, column=C['fresh1_value'], value=fresh1_value)
        ws.cell(row=row_idx, column=C['fresh2_qty'], value=fresh2_qty)
        ws.cell(row=row_idx, column=C['fresh2_rate'], value=fresh2_rate)
        ws.cell(row=row_idx, column=C['fresh2_value'], value=fresh2_value)
        ws.cell(row=row_idx, column=C['adh_qty'], value=adh_qty)
        ws.cell(row=row_idx, column=C['adh_solids'], value=adh_solids)
        ws.cell(row=row_idx, column=C['adh_rate'], value=adh_rate)
        ws.cell(row=row_idx, column=C['adh_value'], value=adh_value)
        ws.cell(row=row_idx, column=C['hard_qty'], value=hard_qty)
        ws.cell(row=row_idx, column=C['hard_solids'], value=hard_solids)
        ws.cell(row=row_idx, column=C['hard_rate'], value=hard_rate)
        ws.cell(row=row_idx, column=C['hard_value'], value=hard_value)
        ws.cell(row=row_idx, column=C['adh_hard_solids'], value=adh_hard_solids)
        ws.cell(row=row_idx, column=C['solv_qty'], value=solv_qty)
        ws.cell(row=row_idx, column=C['solv_rate'], value=solv_rate)
        ws.cell(row=row_idx, column=C['solv_value'], value=solv_value)
        ws.cell(row=row_idx, column=C['fresh_mat_qty'], value=fresh_mat_qty)
        ws.cell(row=row_idx, column=C['fresh_mat_value'], value=fresh_mat_value)
        ws.cell(row=row_idx, column=C['adh_hard_solids_qty'], value=adh_hard_solids)
        ws.cell(row=row_idx, column=C['adh_hard_solv_val'], value=adh_hard_solv_val)
        ws.cell(row=row_idx, column=C['total_input_qty'], value=total_input_qty)
        ws.cell(row=row_idx, column=C['total_input_val'], value=total_input_val)
        ws.cell(row=row_idx, column=C['output_kgs'], value=output_kgs)
        ws.cell(row=row_idx, column=C['per_kg_rmc'], value=per_kg_rmc)
        ws.cell(row=row_idx, column=C['output_mtrs'], value=output_mtrs)

        # ── Cache for Pivot_Lam Rates ──
        lam_pass_key = lam_process if lam_process else 'L1'
        pivot_key = (order_upper, lam_pass_key)
        if pivot_key not in lam_cache:
            lam_cache[pivot_key] = {'output_kgs': 0, 'total_input_val': 0}
        lam_cache[pivot_key]['output_kgs'] += output_kgs
        lam_cache[pivot_key]['total_input_val'] += total_input_val

        # ── Cache for RMC Summary (aggregated by order) ──
        if order_upper not in order_cache:
            order_cache[order_upper] = {
                'output_kgs': 0, 'total_input_val': 0, 'total_input_qty': 0,
                'ptd_qty': 0, 'ptd_value': 0,
                'fresh_mat_qty': 0, 'fresh_mat_value': 0,
                'adh_hard_solids': 0, 'adh_hard_solv_val': 0,
                'wastage_kgs': 0, 'wastage_val': 0,
            }
        oc = order_cache[order_upper]
        oc['output_kgs'] += output_kgs
        oc['total_input_val'] += total_input_val
        oc['total_input_qty'] += total_input_qty
        oc['ptd_qty'] += ptd_qty
        oc['ptd_value'] += ptd_value
        oc['fresh_mat_qty'] += fresh_mat_qty
        oc['fresh_mat_value'] += fresh_mat_value
        oc['adh_hard_solids'] += adh_hard_solids
        oc['adh_hard_solv_val'] += adh_hard_solv_val
        wastage_kgs_row = total_input_qty - output_kgs
        wastage_val_row = wastage_kgs_row * per_kg_rmc if per_kg_rmc > 0 else 0
        oc['wastage_kgs'] += wastage_kgs_row
        oc['wastage_val'] += wastage_val_row

        row_idx += 1

    total_rows = row_idx - DATA_START

    # Build Pivot_Lam Rates
    for (order, lpass), data in lam_cache.items():
        key = order + lpass
        if data['output_kgs'] > 0:
            ctx.pivot_lam_rates[key] = data['total_input_val'] / data['output_kgs']
        ctx.lam_rate_cache[(order, lpass)] = ctx.pivot_lam_rates.get(key, 0)

    ctx.lam_by_order = order_cache

    ctx._log(f"  Lam filled: {total_rows} rows")
    ctx._log(f"    Ptd rates: {filled_ptd}, Adh rates: {filled_adh}, "
             f"Hard rates: {filled_hard}, Solv rates: {filled_solv}")
    ctx._log(f"    Pivot_Lam_Rates: {len(ctx.pivot_lam_rates)} entries")
    ctx._log(f"    Solvent rate: {solvent_rate:.4f}")
