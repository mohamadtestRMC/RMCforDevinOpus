"""
Per-Cell Audit Panel — Trace any RMC Summary cell back to its source.
"""
from __future__ import annotations
import logging
from typing import Dict, List, Any

logger = logging.getLogger(__name__)

RMC_COLUMN_MAP = {
    'OPN WIP Qty': (9, 'OPN_WIP', 2, 8, 'SUMIF(OPN_WIP!B,order,OPN_WIP!H)'),
    'Print Film Kgs': (10, 'Print', 2, 10, 'SUMIF(Print!B,order,Print!J)'),
    'Lam Fresh Qty': (11, 'Lam', 2, 51, 'SUMIF(Lam!B,order,Lam!AY)'),
    'Slit Input Kgs': (12, 'Slit', 2, 11, 'SUMIFS(Slit!K,Slit!B=order)'),
    'Print Ink Kgs': (13, 'Print', 2, 11, 'SUMIF(Print!B,order,Print!K)'),
    'Lam Adh+Hard': (14, 'Lam', 2, 53, 'SUMIF(Lam!B,order,Lam!BA)'),
    'CLS WIP Qty': (16, 'CLS_WIP', 2, 8, 'SUMIF(CLS_WIP!B,order,CLS_WIP!H)'),
    'OPN WIP Value': (17, 'OPN_WIP', 2, 10, 'SUMIF(OPN_WIP!B,order,OPN_WIP!J)'),
    'Print Film Value': (18, 'Print', 2, 13, 'SUMIF(Print!B,order,Print!M)'),
    'Lam Fresh Value': (19, 'Lam', 2, 52, 'SUMIF(Lam!B,order,Lam!AZ)'),
    'Slit Value': (20, 'Slit', 2, 15, 'SUMIFS(Slit!O,Slit!B=order)'),
    'Ink Value': (21, 'Print', 2, 14, 'SUMIF(Print!B,order,Print!N)'),
    'Lam Chem Value': (22, 'Lam', 2, 54, 'SUMIF(Lam!B,order,Lam!BB)'),
    'CLS WIP Value': (24, 'CLS_WIP', 2, 10, 'SUMIF(CLS_WIP!B,order,CLS_WIP!J)'),
    'FG Output': (25, 'FG', 1, 7, 'VLOOKUP(order,FG!A:G,7,0)'),
    'Total Cost': (26, None, None, None, 'Q+R+S+T+U+V+W-X'),
    'RMC/Kg': (27, None, None, None, 'Total Cost / FG Output'),
}

PROCESS_FILTERS = {
    'Print': 'PRINTING', 'Lam': 'LAM', 'Slit': 'SLITTING',
    'Bag&Pouch': 'POUCHING', 'Spout&Valve': 'SPOUT & VALVE',
    'BFL': 'BFL', 'HCI Rew': 'REWINDING', 'PTR Rew': 'REWINDING',
    'OPN_WIP': 'Opening Balance', 'CLS_WIP': 'Closing Balance',
    'FG': 'All Processes',
}

def _sf(v):
    try: return float(v) if v is not None else 0.0
    except: return 0.0

def _get_rmc_ws(wb):
    for n in ['RMC summary', 'RMC Summary']:
        if n in wb.sheetnames: return wb[n]
    return None

def trace_rmc_cell(wb, order: str, column_name: str) -> Dict:
    if column_name not in RMC_COLUMN_MAP:
        return {'error': f'Unknown column: {column_name}'}
    col_num, src, ocol, vcol, formula = RMC_COLUMN_MAP[column_name]
    ou = order.strip().upper()
    result = {'column': column_name, 'rmc_col': col_num, 'formula': formula,
              'source': src or 'Computed', 'filter': PROCESS_FILTERS.get(src, 'N/A'),
              'value': 0.0, 'rows': []}
    if src is None:
        ws = _get_rmc_ws(wb)
        if ws:
            for r in range(7, ws.max_row + 1):
                o = ws.cell(row=r, column=2).value
                if o and str(o).strip().upper() == ou:
                    result['value'] = _sf(ws.cell(row=r, column=col_num).value)
                    break
        return result
    if src not in wb.sheetnames:
        result['error'] = f'{src} sheet not found'
        return result
    ws = wb[src]
    start = 7 if src == 'Print' else (3 if src == 'FG' else 5)
    total = 0.0
    for r in range(start, ws.max_row + 1):
        o = ws.cell(row=r, column=ocol).value
        if not o or str(o).strip().upper() != ou: continue
        val = _sf(ws.cell(row=r, column=vcol).value)
        total += val
        info = {'row': r, 'value': val}
        if src == 'Print':
            info['material'] = str(ws.cell(row=r, column=9).value or '')
        elif src == 'Lam':
            info['lam_process'] = str(ws.cell(row=r, column=8).value or '')
        elif src in ('OPN_WIP', 'CLS_WIP'):
            info['process'] = str(ws.cell(row=r, column=5).value or '')
            info['rate'] = _sf(ws.cell(row=r, column=9).value)
        result['rows'].append(info)
    result['value'] = total
    return result

def trace_all_columns(wb, order: str) -> Dict:
    return {k: trace_rmc_cell(wb, order, k) for k in RMC_COLUMN_MAP}

def get_rmc_orders(wb) -> List[str]:
    ws = _get_rmc_ws(wb)
    if not ws: return []
    return [str(ws.cell(row=r, column=2).value).strip()
            for r in range(7, ws.max_row + 1)
            if ws.cell(row=r, column=2).value]

def get_columns() -> List[str]:
    return list(RMC_COLUMN_MAP.keys())
