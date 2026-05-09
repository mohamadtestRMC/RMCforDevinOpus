"""
Manual-Check Flag List — Automated anomaly detection for RMC pipeline output.

Detects: negative output, output > inputs, RMC outliers, zero rates,
fallback paths, missing FG, WIP imbalance, excessive wastage.
"""
from __future__ import annotations
import logging
import pandas as pd
import numpy as np
from typing import List, Dict, Any

logger = logging.getLogger(__name__)

# Flag severity levels
SEVERITY_HIGH = "🔴 HIGH"
SEVERITY_MEDIUM = "🟡 MEDIUM"
SEVERITY_LOW = "🟢 LOW"


def check_negative_output(ws, sheet_name: str, order_col: int, output_col: int,
                          start_row: int = 7) -> List[Dict]:
    """Flag rows where output_kgs < 0."""
    flags = []
    for r in range(start_row, ws.max_row + 1):
        order = ws.cell(row=r, column=order_col).value
        if not order:
            continue
        output = ws.cell(row=r, column=output_col).value
        try:
            val = float(output) if output is not None else 0
        except (ValueError, TypeError):
            continue
        if val < 0:
            flags.append({
                'severity': SEVERITY_HIGH,
                'category': 'Negative Output',
                'sheet': sheet_name,
                'order': str(order).strip(),
                'row': r,
                'cell': f"Row {r}",
                'value': round(val, 2),
                'action': f'Output is {val:.2f} Kg (negative). Verify Jobtrack input data.',
            })
    return flags


def check_output_exceeds_input(ws, sheet_name: str, order_col: int,
                                input_col: int, output_col: int,
                                start_row: int = 7, tolerance: float = 1.05) -> List[Dict]:
    """Flag rows where output > input * tolerance (5% margin)."""
    flags = []
    for r in range(start_row, ws.max_row + 1):
        order = ws.cell(row=r, column=order_col).value
        if not order:
            continue
        try:
            inp = float(ws.cell(row=r, column=input_col).value or 0)
            out = float(ws.cell(row=r, column=output_col).value or 0)
        except (ValueError, TypeError):
            continue
        if inp > 0 and out > inp * tolerance:
            flags.append({
                'severity': SEVERITY_HIGH,
                'category': 'Output > Input',
                'sheet': sheet_name,
                'order': str(order).strip(),
                'row': r,
                'cell': f"Row {r}",
                'value': f"In={inp:.1f}, Out={out:.1f}",
                'action': f'Output ({out:.1f}) exceeds input ({inp:.1f}) by {(out/inp-1)*100:.1f}%. Check for data entry errors.',
            })
    return flags


def check_rmc_outliers(rmc_summary_ws, start_row: int = 7) -> List[Dict]:
    """Flag orders where RMC/Kg is an outlier (>3x or <1/3 of median)."""
    flags = []
    rmc_values = []
    rmc_rows = []

    for r in range(start_row, rmc_summary_ws.max_row + 1):
        order = rmc_summary_ws.cell(row=r, column=2).value
        if not order:
            continue
        rmc_kg = rmc_summary_ws.cell(row=r, column=27).value  # AA = RMC/Kg
        try:
            val = float(rmc_kg) if rmc_kg is not None else 0
        except (ValueError, TypeError):
            continue
        if val > 0:
            rmc_values.append(val)
            rmc_rows.append((r, str(order).strip(), val))

    if len(rmc_values) < 3:
        return flags

    median_rmc = float(np.median(rmc_values))
    iqr = float(np.percentile(rmc_values, 75) - np.percentile(rmc_values, 25))
    upper = median_rmc + 2 * max(iqr, median_rmc * 0.5)
    lower = max(0.01, median_rmc - 2 * max(iqr, median_rmc * 0.5))

    for r, order, val in rmc_rows:
        if val > upper:
            flags.append({
                'severity': SEVERITY_MEDIUM,
                'category': 'RMC Outlier (High)',
                'sheet': 'RMC Summary',
                'order': order,
                'row': r,
                'cell': f"AA{r}",
                'value': f"AED {val:.2f}/Kg (median: {median_rmc:.2f})",
                'action': f'RMC/Kg is {val/median_rmc:.1f}x above median. Review material rates and input quantities.',
            })
        elif val < lower and val > 0:
            flags.append({
                'severity': SEVERITY_MEDIUM,
                'category': 'RMC Outlier (Low)',
                'sheet': 'RMC Summary',
                'order': order,
                'row': r,
                'cell': f"AA{r}",
                'value': f"AED {val:.2f}/Kg (median: {median_rmc:.2f})",
                'action': f'RMC/Kg is unusually low ({val/median_rmc:.1%} of median). May indicate missing cost components.',
            })
    return flags


def check_zero_rates(ws, sheet_name: str, order_col: int, qty_col: int,
                     rate_col: int, rate_name: str, start_row: int = 7) -> List[Dict]:
    """Flag rows where qty > 0 but rate = 0."""
    flags = []
    for r in range(start_row, ws.max_row + 1):
        order = ws.cell(row=r, column=order_col).value
        if not order:
            continue
        try:
            qty = float(ws.cell(row=r, column=qty_col).value or 0)
            rate = float(ws.cell(row=r, column=rate_col).value or 0)
        except (ValueError, TypeError):
            continue
        if qty > 0 and rate == 0:
            flags.append({
                'severity': SEVERITY_MEDIUM,
                'category': 'Zero Rate',
                'sheet': sheet_name,
                'order': str(order).strip(),
                'row': r,
                'cell': f"Row {r}, {rate_name}",
                'value': f"Qty={qty:.1f}, Rate=0",
                'action': f'{rate_name} is 0 but quantity is {qty:.1f}. MRR lookup may have failed.',
            })
    return flags


def check_missing_fg(rmc_summary_ws, start_row: int = 7) -> List[Dict]:
    """Flag orders in RMC Summary with cost data but no FG output."""
    flags = []
    for r in range(start_row, rmc_summary_ws.max_row + 1):
        order = rmc_summary_ws.cell(row=r, column=2).value
        if not order:
            continue
        try:
            total_cost = float(rmc_summary_ws.cell(row=r, column=26).value or 0)  # Z
            fg_output = float(rmc_summary_ws.cell(row=r, column=25).value or 0)   # Y
        except (ValueError, TypeError):
            continue
        if total_cost > 0 and fg_output == 0:
            flags.append({
                'severity': SEVERITY_HIGH,
                'category': 'Missing FG Output',
                'sheet': 'RMC Summary',
                'order': str(order).strip(),
                'row': r,
                'cell': f"Y{r}",
                'value': f"Cost=AED {total_cost:,.0f}, FG=0",
                'action': 'Order has cost data but no FG output. RMC/Kg cannot be computed.',
            })
    return flags


def check_wip_imbalance(rmc_summary_ws, start_row: int = 7) -> List[Dict]:
    """Flag orders with extreme WIP imbalance (CLS >> OPN or vice versa)."""
    flags = []
    for r in range(start_row, rmc_summary_ws.max_row + 1):
        order = rmc_summary_ws.cell(row=r, column=2).value
        if not order:
            continue
        try:
            opn_val = float(rmc_summary_ws.cell(row=r, column=17).value or 0)  # Q
            cls_val = float(rmc_summary_ws.cell(row=r, column=24).value or 0)  # X
        except (ValueError, TypeError):
            continue
        if opn_val > 0 and cls_val > opn_val * 3:
            flags.append({
                'severity': SEVERITY_MEDIUM,
                'category': 'WIP Imbalance',
                'sheet': 'RMC Summary',
                'order': str(order).strip(),
                'row': r,
                'cell': f"Row {r}",
                'value': f"OPN={opn_val:,.0f}, CLS={cls_val:,.0f}",
                'action': f'CLS_WIP ({cls_val:,.0f}) is {cls_val/opn_val:.1f}x OPN_WIP ({opn_val:,.0f}). Verify WIP quantities.',
            })
        elif cls_val > 0 and opn_val > cls_val * 5:
            flags.append({
                'severity': SEVERITY_MEDIUM,
                'category': 'WIP Imbalance',
                'sheet': 'RMC Summary',
                'order': str(order).strip(),
                'row': r,
                'cell': f"Row {r}",
                'value': f"OPN={opn_val:,.0f}, CLS={cls_val:,.0f}",
                'action': f'OPN_WIP ({opn_val:,.0f}) is {opn_val/cls_val:.1f}x CLS_WIP ({cls_val:,.0f}). Large drawdown — verify.',
            })
    return flags


def check_excessive_wastage(ws, sheet_name: str, order_col: int,
                            input_col: int, output_col: int,
                            start_row: int = 7, threshold: float = 0.15) -> List[Dict]:
    """Flag rows where wastage exceeds threshold (default 15%)."""
    flags = []
    for r in range(start_row, ws.max_row + 1):
        order = ws.cell(row=r, column=order_col).value
        if not order:
            continue
        try:
            inp = float(ws.cell(row=r, column=input_col).value or 0)
            out = float(ws.cell(row=r, column=output_col).value or 0)
        except (ValueError, TypeError):
            continue
        if inp > 0 and out >= 0:
            waste_pct = (inp - out) / inp
            if waste_pct > threshold:
                flags.append({
                    'severity': SEVERITY_MEDIUM,
                    'category': 'Excessive Wastage',
                    'sheet': sheet_name,
                    'order': str(order).strip(),
                    'row': r,
                    'cell': f"Row {r}",
                    'value': f"{waste_pct*100:.1f}% (threshold: {threshold*100:.0f}%)",
                    'action': f'Wastage is {waste_pct*100:.1f}%, exceeding the {threshold*100:.0f}% threshold. Investigate root cause.',
                })
    return flags


def check_fallback_paths(log: list) -> List[Dict]:
    """Flag orders that used fallback rate paths from pipeline log."""
    flags = []
    seen = set()
    for entry in log:
        entry_lower = entry.lower() if isinstance(entry, str) else ''
        if 'fallback' in entry_lower or 'month avg' in entry_lower or 'no rate' in entry_lower:
            # Extract order if present
            order = ''
            for part in entry.split():
                if part.startswith('WO=') or part.startswith('order='):
                    order = part.split('=', 1)[1].strip(',')
            key = (order, entry[:80])
            if key in seen:
                continue
            seen.add(key)
            flags.append({
                'severity': SEVERITY_MEDIUM,
                'category': 'Fallback Path',
                'sheet': 'Pipeline',
                'order': order or 'N/A',
                'row': 0,
                'cell': 'Log',
                'value': entry[:120],
                'action': 'Rate was computed using fallback logic (not exact MRR match). Verify accuracy.',
            })
    return flags


def run_all_checks(wb, log: list = None) -> List[Dict]:
    """Run all checks against the filled workbook and return consolidated flag list."""
    all_flags = []
    sn = wb.sheetnames

    # Print sheet checks
    if 'Print' in sn:
        ws = wb['Print']
        all_flags.extend(check_negative_output(ws, 'Print', 2, 16, 7))
        all_flags.extend(check_output_exceeds_input(ws, 'Print', 2, 12, 16, 7))
        all_flags.extend(check_zero_rates(ws, 'Print', 2, 12, 13, 'Film Value', 7))
        all_flags.extend(check_excessive_wastage(ws, 'Print', 2, 12, 16, 7))

    # Lam sheet checks
    if 'Lam' in sn:
        ws = wb['Lam']
        all_flags.extend(check_negative_output(ws, 'Lam', 2, 57, 7))
        all_flags.extend(check_output_exceeds_input(ws, 'Lam', 2, 55, 57, 7))
        all_flags.extend(check_zero_rates(ws, 'Lam', 2, 26, 27, 'Ptd Rate', 7))

    # Slit sheet checks
    if 'Slit' in sn:
        ws = wb['Slit']
        all_flags.extend(check_negative_output(ws, 'Slit', 2, 12, 5))
        all_flags.extend(check_excessive_wastage(ws, 'Slit', 2, 11, 12, 5))

    # RMC Summary checks
    rmc_name = 'RMC summary' if 'RMC summary' in sn else ('RMC Summary' if 'RMC Summary' in sn else None)
    if rmc_name:
        ws_rmc = wb[rmc_name]
        all_flags.extend(check_rmc_outliers(ws_rmc))
        all_flags.extend(check_missing_fg(ws_rmc))
        all_flags.extend(check_wip_imbalance(ws_rmc))

    # Pipeline log checks
    if log:
        all_flags.extend(check_fallback_paths(log))

    # Sort: HIGH first, then MEDIUM, then LOW
    severity_order = {SEVERITY_HIGH: 0, SEVERITY_MEDIUM: 1, SEVERITY_LOW: 2}
    all_flags.sort(key=lambda f: severity_order.get(f['severity'], 9))

    logger.info(f"Flag checker: {len(all_flags)} flags detected "
                f"({sum(1 for f in all_flags if f['severity']==SEVERITY_HIGH)} HIGH, "
                f"{sum(1 for f in all_flags if f['severity']==SEVERITY_MEDIUM)} MEDIUM)")

    return all_flags
