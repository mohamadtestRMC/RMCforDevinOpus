"""
RMC Summary Filler — Reads DIRECTLY from filled process sheets using SUMIF.

Instead of using Python caches (which have wrong values), this reads from
the actual filled process sheets in ctx.wb — exactly like Excel SUMIF does.

Formula reference (from filled workbook):
  I  = SUMIF(OPN_WIP!B, order, OPN_WIP!H)
  J  = SUMIF(Print!B, order, Print!G)       # G = Film Input Kgs
  K  = SUMIF(Lam!B, order, Lam!AY)          # AY=51
  L  = SUMIFS(Slit!K, Slit!B=order, Slit!F=material)
  M  = SUMIF(Print!B, order, Print!H)       # H = Dry Ink Kgs
  N  = SUMIF(Lam!B, order, Lam!BA)          # BA=53
  O  = SUMIF(B&P!B, order, B&P!Y) + SUMIF(S&V!B, order, S&V!AJ)
  P  = SUMIF(CLS_WIP!B, order, CLS_WIP!H)
  Q  = SUMIF(OPN_WIP!B, order, OPN_WIP!J)
  R  = SUMIF(Print!B, order, Print!J)       # J = Film Value
  S  = SUMIF(Lam!B, order, Lam!AZ)          # AZ=52
  T  = SUMIFS(Slit!O, Slit!B=order, Slit!F=material)
  U  = SUMIF(Print!B, order, Print!K)       # K = Ink Value
  V  = SUMIF(Lam!B, order, Lam!BB)          # BB=54
  W  = SUMIF(B&P!B, order, B&P!Z) + SUMIF(S&V!B, order, S&V!AK)
  X  = SUMIF(CLS_WIP!B, order, CLS_WIP!J)
  Y  = VLOOKUP(order, FG!A:G, 7, 0)
  Z  = Q+R+S+T+U+V+W-X
  AA = Z/Y
"""
from __future__ import annotations
import logging
from engine.base_rmc.context import RMCContext

logger = logging.getLogger(__name__)
DATA_START = 7


def _sf(v):
    if v is None: return 0.0
    try: return float(v)
    except: return 0.0


def _build_sheet_index(ws, order_col, value_cols, start_row=5, max_row=None):
    """Build {ORDER: {col: sum_value}} from a worksheet, like SUMIF."""
    idx = {}
    mr = max_row or ws.max_row
    for r in range(start_row, mr + 1):
        order = ws.cell(row=r, column=order_col).value
        if not order: continue
        ou = str(order).strip().upper()
        if ou not in idx:
            idx[ou] = {c: 0.0 for c in value_cols}
        for c in value_cols:
            idx[ou][c] = idx[ou].get(c, 0.0) + _sf(ws.cell(row=r, column=c).value)
    return idx


def _build_slit_index(ws, order_col=2, mat_col=6, start_row=5):
    """Build Slit index keyed by (ORDER, MATERIAL) for SUMIFS."""
    idx = {}  # {ORDER: {col: value}}
    idx_mat = {}  # {(ORDER, MAT): {col: value}}
    for r in range(start_row, ws.max_row + 1):
        order = ws.cell(row=r, column=order_col).value
        if not order: continue
        ou = str(order).strip().upper()
        mat = str(ws.cell(row=r, column=mat_col).value or '').strip().upper()
        # Simple SUMIF by order (no material filter)
        if ou not in idx:
            idx[ou] = {}
        for c in [11, 15, 16, 17]:  # K=11, O=15, P=16, Q=17
            idx[ou][c] = idx[ou].get(c, 0.0) + _sf(ws.cell(row=r, column=c).value)
        # SUMIFS by (order, material)
        key = (ou, mat)
        if key not in idx_mat:
            idx_mat[key] = {}
        for c in [11, 15]:
            idx_mat[key][c] = idx_mat[key].get(c, 0.0) + _sf(ws.cell(row=r, column=c).value)
    return idx, idx_mat


def _build_fg_index(ws, start_row=3):
    """Build FG lookup: VLOOKUP(order, FG!A:G, 7, 0)."""
    idx = {}
    for r in range(start_row, ws.max_row + 1):
        order = ws.cell(row=r, column=1).value
        if not order: continue
        ou = str(order).strip().upper()
        if ou not in idx:  # VLOOKUP takes first match
            fg_val = _sf(ws.cell(row=r, column=7).value)
            idx[ou] = fg_val
    return idx


def fill_rmc_summary(ctx: RMCContext) -> None:
    """Fill RMC Summary by reading directly from filled process sheets."""
    ctx._log("Filling RMC Summary (sheet-based SUMIF)...")

    sn = ctx.wb.sheetnames
    ws = ctx.wb['RMC summary'] if 'RMC summary' in sn else (ctx.wb['RMC Summary'] if 'RMC Summary' in sn else None)
    if not ws:
        ctx._error("  RMC summary sheet not found!")
        return

    wb = ctx.wb
    sn = wb.sheetnames

    # ── Build indexes from each process sheet ──
    ctx._log("  Building sheet indexes...")

    # Print: B=order, G=Film Input Kgs, H=Dry Ink, J=Film Value, K=Ink Value
    print_idx = {}
    if 'Print' in sn:
        print_idx = _build_sheet_index(wb['Print'], 2, [7, 8, 10, 11, 17, 18], start_row=7)
    ctx._log(f"    Print: {len(print_idx)} orders")

    # Lam: B=order, AY=51, AZ=52, BA=53, BB=54, BJ=62, BL=64
    lam_idx = {}
    if 'Lam' in sn:
        lam_idx = _build_sheet_index(wb['Lam'], 2, [51, 52, 53, 54, 62, 64], start_row=5)
    ctx._log(f"    Lam: {len(lam_idx)} orders")

    # BFL: B=order, P=16 wastage_kgs, Q=17 wastage_val
    bfl_idx = {}
    if 'BFL' in sn:
        bfl_idx = _build_sheet_index(wb['BFL'], 2, [16, 17], start_row=5)
    ctx._log(f"    BFL: {len(bfl_idx)} orders")

    # Slit: B=order, K=11 input_kgs, O=15 input_val, P=16 waste_kgs, Q=17 waste_val
    slit_idx = {}
    slit_mat_idx = {}
    if 'Slit' in sn:
        slit_idx, slit_mat_idx = _build_slit_index(wb['Slit'])
    ctx._log(f"    Slit: {len(slit_idx)} orders")

    # Bag&Pouch: B=order, P=16 waste_kgs, S=19 waste_val, Y=25 input_kgs, Z=26 total_val
    bp_idx = {}
    if 'Bag&Pouch' in sn:
        bp_idx = _build_sheet_index(wb['Bag&Pouch'], 2, [16, 19, 25, 26], start_row=5)
    ctx._log(f"    Bag&Pouch: {len(bp_idx)} orders")

    # Spout&Valve: B=order, AC=29 waste_kgs, AE=31 waste_val, AJ=36 input_kgs, AK=37 total_val
    sv_idx = {}
    if 'Spout&Valve' in sn:
        sv_idx = _build_sheet_index(wb['Spout&Valve'], 2, [29, 31, 36, 37], start_row=5)
    ctx._log(f"    Spout&Valve: {len(sv_idx)} orders")

    # HCI Rew: B=order, I=9 waste_kgs, K=11 waste_val
    hci_idx = {}
    if 'HCI Rew' in sn:
        hci_idx = _build_sheet_index(wb['HCI Rew'], 2, [9, 11], start_row=5)
    ctx._log(f"    HCI Rew: {len(hci_idx)} orders")

    # PTR Rew: B=order, AE=31 waste_kgs, AF=32 waste_val
    ptr_idx = {}
    if 'PTR Rew' in sn:
        ptr_idx = _build_sheet_index(wb['PTR Rew'], 2, [31, 32], start_row=5)
    ctx._log(f"    PTR Rew: {len(ptr_idx)} orders")

    # OPN_WIP: B=order, H=8 qty, J=10 value
    opn_idx = {}
    if 'OPN_WIP' in sn:
        opn_idx = _build_sheet_index(wb['OPN_WIP'], 2, [8, 10], start_row=5)
    ctx._log(f"    OPN_WIP: {len(opn_idx)} orders")

    # CLS_WIP: B=order, H=8 qty, J=10 value
    cls_idx = {}
    if 'CLS_WIP' in sn:
        cls_idx = _build_sheet_index(wb['CLS_WIP'], 2, [8, 10], start_row=5)
    ctx._log(f"    CLS_WIP: {len(cls_idx)} orders")

    # FG: A=order, G=7 final_fg (VLOOKUP col 7)
    fg_idx = {}
    if 'FG' in sn:
        fg_idx = _build_fg_index(wb['FG'])
    ctx._log(f"    FG: {len(fg_idx)} orders")

    # ── Collect ALL unique orders ──
    all_orders = set()
    for idx in [print_idx, lam_idx, bfl_idx, slit_idx, bp_idx, sv_idx,
                hci_idx, ptr_idx, opn_idx, cls_idx, fg_idx]:
        all_orders.update(idx.keys())
    if ctx.order_list:
        for o in ctx.order_list:
            all_orders.add(str(o).strip().upper())
    all_orders = sorted(all_orders)
    ctx._log(f"  Total unique orders: {len(all_orders)}")

    # ── Get order metadata from Jobtrack ──
    meta_cache = {}
    if ctx.jobtrack_df is not None and not ctx.jobtrack_df.empty:
        df = ctx.jobtrack_df
        oc = None
        for c in df.columns:
            if 'order' in str(c).lower():
                oc = c; break
        if oc:
            for _, row in df.iterrows():
                o = str(row.get(oc, '')).strip().upper()
                if o and o not in meta_cache:
                    m = {}
                    for c in df.columns:
                        cl = str(c).lower()
                        if 'design' in cl: m['design'] = str(row[c]) if row[c] else ''
                        elif 'customer' in cl: m['customer'] = str(row[c]) if row[c] else ''
                        elif cl in ('material', 'sales code'): m['material'] = str(row[c]) if row[c] else ''
                        elif 'structure' in cl: m['structure'] = str(row[c]) if row[c] else ''
                    meta_cache[o] = m

    # ── Write RMC Summary rows ──
    filled = 0
    row = DATA_START

    for order in all_orders:
        ou = order
        meta = meta_cache.get(ou, {})
        pr = print_idx.get(ou, {})
        lm = lam_idx.get(ou, {})
        bf = bfl_idx.get(ou, {})
        sl = slit_idx.get(ou, {})
        bp = bp_idx.get(ou, {})
        sv = sv_idx.get(ou, {})
        hc = hci_idx.get(ou, {})
        pt = ptr_idx.get(ou, {})
        op = opn_idx.get(ou, {})
        cl = cls_idx.get(ou, {})
        fg = fg_idx.get(ou, 0)

        # Order info (B-H)
        ws.cell(row=row, column=2, value=order)
        ws.cell(row=row, column=3, value=meta.get('design') or None)
        ws.cell(row=row, column=4, value=meta.get('customer') or None)
        ws.cell(row=row, column=5, value=meta.get('material') or None)
        ws.cell(row=row, column=6, value=meta.get('structure') or None)

        # QUANTITIES (I-P)
        I = op.get(8, 0)           # OPN_WIP qty
        J = pr.get(7, 0)          # Print Film Input Kgs
        K = lm.get(51, 0)         # Lam Fresh Mat Qty (AY)
        L = sl.get(11, 0)         # Slit Input Kgs
        M = pr.get(8, 0)          # Print Dry Ink Kgs
        N = lm.get(53, 0)         # Lam Adh+Hard Solids (BA)
        O = bp.get(25, 0) + sv.get(36, 0)  # B&P Y + S&V AJ
        P = cl.get(8, 0)          # CLS_WIP qty

        for c, v in [(9,I),(10,J),(11,K),(12,L),(13,M),(14,N),(15,O),(16,P)]:
            ws.cell(row=row, column=c, value=v if v else None)

        # VALUES (Q-X)
        Q = op.get(10, 0)         # OPN_WIP value
        R = pr.get(10, 0)         # Print Film Value (J)
        S = lm.get(52, 0)         # Lam Fresh Mat Value (AZ)
        T = sl.get(15, 0)         # Slit Input Value
        U = pr.get(11, 0)         # Print Ink Value (K)
        V = lm.get(54, 0)         # Lam Adh+Hard+Sol Value (BB)
        W = bp.get(26, 0) + sv.get(37, 0)  # B&P Z + S&V AK
        X = cl.get(10, 0)         # CLS_WIP value

        for c, v in [(17,Q),(18,R),(19,S),(20,T),(21,U),(22,V),(23,W),(24,X)]:
            ws.cell(row=row, column=c, value=v if v else None)

        # COMPUTED (Y-AA)
        Y = fg                     # FG Output
        Z = Q + R + S + T + U + V + W - X  # Total Cost
        AA = Z / Y if Y > 0 else 0  # RMC/Kg

        ws.cell(row=row, column=25, value=Y if Y else None)
        ws.cell(row=row, column=26, value=Z if Z else None)
        ws.cell(row=row, column=27, value=AA if AA else None)

        # WASTAGE KGS (AJ-AS)
        w_bfl = bf.get(16, 0)
        w_pr = pr.get(17, 0)
        w_lm = lm.get(64, 0)      # BL=64
        w_sl = sl.get(16, 0)
        w_bp = bp.get(16, 0)
        w_sv = sv.get(29, 0)       # AC=29
        w_hc = hc.get(9, 0)
        w_pt = pt.get(31, 0)       # AE=31
        w_total = w_pr + w_lm + w_sl + w_bp + w_sv + w_hc + w_pt

        for c, v in [(36,w_bfl),(38,w_pr),(39,w_lm),(40,w_sl),(41,w_bp),
                      (42,w_sv),(43,w_hc),(44,w_pt),(45,w_total)]:
            ws.cell(row=row, column=c, value=v if v else None)

        # CURRENT WASTAGE VALUE (BN-BU)
        wv_bfl = bf.get(17, 0)
        wv_pr = pr.get(18, 0)
        wv_lm = lm.get(62, 0)     # BJ=62
        wv_sl = sl.get(17, 0)
        wv_bp = bp.get(19, 0)     # S=19
        wv_sv = sv.get(31, 0)     # AE=31
        wv_hc = hc.get(11, 0)
        wv_pt = pt.get(32, 0)     # AF=32

        for c, v in [(66,wv_bfl),(67,wv_pr),(68,wv_lm),(69,wv_sl),
                      (70,wv_bp),(71,wv_sv),(72,wv_hc),(73,wv_pt)]:
            ws.cell(row=row, column=c, value=v if v else None)

        # COMBINED WASTAGE (AU-BC) = prev(BF-BM) + current(BN-BU)
        for i, (prev_c, cur_v) in enumerate([(58,wv_bfl),(59,wv_pr),(60,wv_lm),
            (61,wv_sl),(62,wv_bp),(63,wv_sv),(64,wv_hc),(65,wv_pt)]):
            prev_v = _sf(ws.cell(row=row, column=prev_c).value)
            ws.cell(row=row, column=47+i, value=(prev_v + cur_v) or None)

        grand_wv = sum(wv_bfl+wv_pr+wv_lm+wv_sl+wv_bp+wv_sv+wv_hc+wv_pt
                       for _ in [1])
        ws.cell(row=row, column=55, value=grand_wv or None)

        # MISC
        ws.cell(row=row, column=76, value=Z if Z else None)
        ws.cell(row=row, column=77, value=(I+J+K+L+M+N+O-P) or None)

        if Z > 0: filled += 1
        row += 1

    ctx._log(f"  RMC Summary: {row-DATA_START} rows, {filled} with cost data")
