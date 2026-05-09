"""
Row Explainer — builds structured explanations for each filled row.
Reuses the results_log produced by fill_jobtrack, no recomputation.
"""
import openpyxl
import io
import logging
from engine.fill_jobtrack import COLS, DATA_START_ROW, HEADER_ROW

logger = logging.getLogger(__name__)

# Columns to display in the raw Excel snapshot
RAW_COLS = {
    'UID': COLS['UID'],
    'Date': COLS['Date'],
    'Process': COLS['Process'],
    'Order No': COLS['Order_No'],
    'Input Name': COLS['Input_Name'],
    'Input Size': COLS['Input_Size'],
    'Input Mic': COLS['Input_Mic'],
    'Input Qty': COLS['Input_Qty'],
    'Bal Qty': COLS['Balance_Qty'],
    'Total Input': COLS['Total_1st_Input'],
    'Film MR#': COLS['Film_MR'],
    'Film Rate': COLS['Film_Rate'],
    'Film Value': COLS['Film_Value'],
    'Fresh1 Name': COLS['Fresh1_Name'],
    'Fresh1 Size': COLS['Fresh1_Size'],
    'Fresh1 Mic': COLS['Fresh1_Mic'],
    'Fresh1 Qty': COLS['Fresh1_Qty'],
    'Fresh1 Bal': COLS['Fresh1_Balance'],
    'Total Fresh1': COLS['Total_Fresh1'],
    'Fresh1 MR#': COLS['Fresh1_MR'],
    'Fresh1 Rate': COLS['Fresh1_Rate'],
    'Fresh1 Value': COLS['Fresh1_Value'],
    'Fresh2 Name': COLS['Fresh2_Name'],
    'Fresh2 Size': COLS['Fresh2_Size'],
    'Fresh2 Mic': COLS['Fresh2_Mic'],
    'Fresh2 Qty': COLS['Fresh2_Qty'],
    'Fresh2 Bal': COLS['Fresh2_Balance'],
    'Total Fresh2': COLS['Total_Fresh2'],
    'Fresh2 MR#': COLS['Fresh2_MR'],
    'Fresh2 Rate': COLS['Fresh2_Rate'],
    'Fresh2 Value': COLS['Fresh2_Value'],
    'Adh Name': COLS['Adh_Name'],
    'Adh Kgs': COLS['Adh_Kgs'],
    'Adh Rate': COLS['Adh_Rate'],
    'Adh Value': COLS['Adh_Value'],
    'Hard Kgs': COLS['Hard_Kgs'],
    'Hard Rate': COLS['Hard_Rate'],
    'Hard Value': COLS['Hard_Value'],
    'Sol Qty': COLS['Sol_Qty'],
    'Sol Rate': COLS['Sol_Rate'],
    'Sol Value': COLS['Sol_Value'],
}


def get_raw_excel_row(output_data: bytes, row_num: int) -> dict:
    """Extract raw cell values for a specific row from the filled workbook."""
    wb = openpyxl.load_workbook(io.BytesIO(output_data), data_only=False)
    ws = wb.active
    raw = {}
    for label, col_idx in RAW_COLS.items():
        val = ws.cell(row=row_num, column=col_idx).value
        if val is not None:
            raw[label] = val
    wb.close()
    return raw


def _detect_issues(log_entries: list, raw_row: dict) -> list:
    """Detect potential issues / risk flags for a row based on its log entries."""
    issues = []
    details_combined = " ".join(e.get('detail', '') for e in log_entries)
    statuses = [e.get('status', '') for e in log_entries]

    # 1. MRR mismatch or extra MRRs
    for e in log_entries:
        mr_val = e.get('detail', '')
        if '/' in str(raw_row.get('Film MR#', '')) or \
           '/' in str(raw_row.get('Fresh1 MR#', '')) or \
           '/' in str(raw_row.get('Fresh2 MR#', '')):
            issues.append({
                'id': 1,
                'title': 'Multiple MRRs detected',
                'detail': 'Engine found multiple MRRs for this material. '
                          'Qty-weighted average was used for rate calculation.',
                'severity': 'MEDIUM'
            })
            break

    # 2. INH (in-house) material
    if raw_row.get('Film MR#') == 'INH':
        issues.append({
            'id': 2,
            'title': 'In-house material (INH)',
            'detail': 'WPE/WLDPE material detected. Rate is from PR material-level '
                      'lookup (not MRR-specific).',
            'severity': 'MEDIUM'
        })

    # 3. Outlier rate adjustment
    if 'outlier' in details_combined.lower() or 'month avg' in details_combined.lower():
        issues.append({
            'id': 3,
            'title': 'Outlier rate adjustment applied',
            'detail': 'MRR-specific rate deviated >50% from current month average. '
                      'Month-level material rate was used instead.',
            'severity': 'LOW'
        })

    # 4. Fallback rate
    if 'fallback' in details_combined.lower() or \
       any('no rate' in s.lower() for s in [e.get('detail', '') for e in log_entries]):
        issues.append({
            'id': 4,
            'title': 'Fallback rate used',
            'detail': 'MRR-specific rate not found in PR. Material-level average rate '
                      'for the reporting month was used.',
            'severity': 'LOW'
        })

    # 5. Dominant MRR filtering
    for e in log_entries:
        detail = e.get('detail', '')
        if 'MR#=' in detail and '/' in detail:
            issues.append({
                'id': 5,
                'title': 'Dominant MRR logic applied',
                'detail': 'Only MRRs contributing ≥10% of total issue qty were used '
                          'for rate calculation.',
                'severity': 'MEDIUM'
            })
            break

    # 6. Warnings
    if any('WARN' in s for s in statuses):
        warn_details = [e['detail'] for e in log_entries if 'WARN' in e.get('status', '')]
        issues.append({
            'id': 6,
            'title': 'Warning flag',
            'detail': '; '.join(warn_details),
            'severity': 'LOW'
        })

    # 7. Missing data
    if any('MISS' in s for s in statuses):
        issues.append({
            'id': 7,
            'title': 'Missing MRR data',
            'detail': 'No matching MRR found in Stores for this material/order combination.',
            'severity': 'LOW'
        })

    return issues


def _compute_confidence(log_entries: list, issues: list) -> dict:
    """Compute confidence score for a row."""
    statuses = [e.get('status', '') for e in log_entries]
    issue_ids = set(i['id'] for i in issues)

    # HIGH: exact match, single MRR, no adjustments
    if not issues or (len(issues) == 0):
        return {'level': 'HIGH', 'color': '#059669',
                'reason': 'Exact match, single MRR, no adjustments'}

    # LOW: fallback rate, missing data, or adjustments applied
    if issue_ids & {3, 4, 7}:
        reasons = [i['title'] for i in issues if i['id'] in {3, 4, 7}]
        return {'level': 'LOW', 'color': '#DC2626',
                'reason': '; '.join(reasons)}

    # LOW: any WARN or MISS
    if any('WARN' in s for s in statuses) or any('MISS' in s for s in statuses):
        reasons = [i['title'] for i in issues if i['severity'] == 'LOW']
        return {'level': 'LOW', 'color': '#DC2626',
                'reason': '; '.join(reasons) if reasons else 'Warning or missing data detected'}

    # MEDIUM: multi-MRR weighted average
    if issue_ids & {1, 2, 5}:
        reasons = [i['title'] for i in issues if i['id'] in {1, 2, 5}]
        return {'level': 'MEDIUM', 'color': '#D97706',
                'reason': '; '.join(reasons)}

    return {'level': 'HIGH', 'color': '#059669',
            'reason': 'No issues detected'}


def _build_rate_breakdown(log_entries: list, raw_row: dict) -> list:
    """Build a rate breakdown from the log entries."""
    breakdown = []

    for e in log_entries:
        entry_type = e.get('type', '')
        detail = e.get('detail', '')
        status = e.get('status', '')

        if 'Rate=' in detail or 'Rate:' in detail or '@' in detail:
            # Parse rate info
            parts = detail.split(',')
            info = {
                'type': entry_type,
                'status': status,
                'detail': detail,
                'parts': {}
            }
            for p in parts:
                p = p.strip()
                if '=' in p:
                    k, v = p.split('=', 1)
                    info['parts'][k.strip()] = v.strip()
                elif '@' in p:
                    # Format like "MB655@11.72"
                    info['parts']['material_rate'] = p.strip()
            breakdown.append(info)
        elif detail:
            breakdown.append({
                'type': entry_type,
                'status': status,
                'detail': detail,
                'parts': {}
            })

    return breakdown


def explain_row(output_data: bytes, results_log: list, row_num: int) -> dict:
    """
    Build a complete explanation for a single Jobtrack row.

    Args:
        output_data: The filled workbook bytes
        results_log: The full results_log from fill_jobtrack
        row_num: The Excel row number (1-based)

    Returns:
        dict with keys: raw_row, issues, confidence, rate_breakdown, log_entries
    """
    # Get raw Excel data
    raw_row = get_raw_excel_row(output_data, row_num)

    # Get log entries for this row
    log_entries = [e for e in results_log if e.get('row') == row_num]

    # Detect issues
    issues = _detect_issues(log_entries, raw_row)

    # Compute confidence
    confidence = _compute_confidence(log_entries, issues)

    # Build rate breakdown
    rate_breakdown = _build_rate_breakdown(log_entries, raw_row)

    return {
        'row_num': row_num,
        'uid': raw_row.get('UID', ''),
        'process': raw_row.get('Process', ''),
        'order_no': raw_row.get('Order No', ''),
        'raw_row': raw_row,
        'log_entries': log_entries,
        'issues': issues,
        'confidence': confidence,
        'rate_breakdown': rate_breakdown,
    }


def build_all_explanations(output_data: bytes, results_log: list) -> dict:
    """
    Pre-build explanations for ALL rows that have log entries.
    Returns {row_num: explanation_dict}.
    Cached in session_state to avoid recomputation.
    """
    rows_with_logs = set(e['row'] for e in results_log if 'row' in e)
    explanations = {}
    for row_num in sorted(rows_with_logs):
        explanations[row_num] = explain_row(output_data, results_log, row_num)
    return explanations
