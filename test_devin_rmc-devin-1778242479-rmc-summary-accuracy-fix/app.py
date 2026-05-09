"""
IPP Production Automation Platform
Tab 1: Jobtrack MRR Auto-Fill (Debug / Fallback)
Tab 2: RMC Report Generator (Main Flow)
"""
import streamlit as st
import pandas as pd
import io
import logging
import base64
from datetime import datetime

from engine.fill_jobtrack import fill_jobtrack, get_filled_data_for_dashboard, COLS, DATA_START_ROW
from engine.explainer import build_all_explanations, explain_row
from engine.pdf_report import generate_row_pdf
from rmc_tab import render_rmc_sidebar, render_rmc_tab
from dashboard.charts import (
    create_cost_breakdown_pie, create_process_cost_bar,
    create_material_usage_chart, create_daily_cost_trend,
    create_wastage_chart, create_top_orders_chart,
    create_coverage_gauge, compute_dashboard_stats,
    create_rate_comparison_chart, create_cost_per_kg_chart,
    create_chemical_rates_summary,
    generate_ai_insights, generate_recommendations,
    create_material_cost_waterfall, create_efficiency_scatter,
    create_cost_pareto, create_process_efficiency_radar,
    COLORS
)

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

st.set_page_config(
    page_title="IPP Production Automation",
    page_icon="assets/ipp-logo.png",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
    
    .stApp {
        background: #F8FAFC;
        font-family: 'Inter', sans-serif;
    }
    header[data-testid="stHeader"] {
        background: rgba(248, 250, 252, 0.9);
        backdrop-filter: blur(12px);
    }
    section[data-testid="stSidebar"] {
        background: #FFFFFF;
        border-right: 1px solid #E2E8F0;
    }
    section[data-testid="stSidebar"] .stMarkdown h1,
    section[data-testid="stSidebar"] .stMarkdown h2,
    section[data-testid="stSidebar"] .stMarkdown h3 {
        color: #1E293B;
    }
    section[data-testid="stSidebar"] .stMarkdown p,
    section[data-testid="stSidebar"] .stMarkdown li {
        color: #64748B;
    }
    
    .kpi-row { display: flex; gap: 14px; flex-wrap: wrap; margin-bottom: 20px; }
    .kpi-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 14px; padding: 18px 22px; flex: 1; min-width: 170px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); transition: all 0.3s ease;
    }
    .kpi-card:hover {
        border-color: #CBD5E1;
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
    }
    .kpi-label {
        font-size: 11px; font-weight: 600; color: #64748B;
        text-transform: uppercase; letter-spacing: 1px; margin-bottom: 4px;
    }
    .kpi-value {
        font-size: 26px; font-weight: 700;
        background: linear-gradient(135deg, #4F46E5, #DB2777);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        line-height: 1.2;
    }
    .kpi-value.green {
        background: linear-gradient(135deg, #059669, #0D9488);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }
    .kpi-value.amber {
        background: linear-gradient(135deg, #D97706, #EA580C);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }
    .kpi-value.blue {
        background: linear-gradient(135deg, #2563EB, #4F46E5);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    }
    .kpi-sub { font-size: 11px; color: #475569; margin-top: 3px; }
    
    .section-header {
        font-size: 18px; font-weight: 700; color: #1E293B;
        margin: 28px 0 12px 0; padding-bottom: 8px;
        border-bottom: 2px solid #E2E8F0;
    }
    .chart-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 14px; padding: 16px; margin-bottom: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }
    .chart-hint {
        font-size: 11px; color: #64748B; margin-top: -8px; margin-bottom: 8px;
        font-style: italic; padding: 0 4px;
    }
    
    .stButton > button {
        background: linear-gradient(135deg, #4F46E5, #7C3AED) !important;
        color: white !important; border: none !important;
        border-radius: 10px !important; padding: 12px 32px !important;
        font-weight: 600 !important; font-size: 15px !important;
    }
    .stButton > button:hover {
        transform: translateY(-1px) !important;
        box-shadow: 0 6px 24px rgba(79, 70, 229, 0.4) !important;
    }
    .stDownloadButton > button {
        background: linear-gradient(135deg, #059669, #047857) !important;
        color: white !important; border: none !important;
        border-radius: 10px !important; padding: 12px 32px !important;
        font-weight: 600 !important; font-size: 15px !important;
    }
    .stProgress > div > div > div {
        background: linear-gradient(90deg, #4F46E5, #DB2777) !important;
    }
    .stTabs [data-baseweb="tab-list"] { gap: 8px; }
    .stTabs [data-baseweb="tab"] {
        background: #F1F5F9; border-radius: 8px;
        color: #64748B; border: 1px solid transparent;
    }
    .stTabs [aria-selected="true"] {
        background: #FFFFFF !important;
        color: #4F46E5 !important; border-color: #E2E8F0 !important;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
    }
    .hero-title {
        font-size: 36px; font-weight: 800;
        background: linear-gradient(135deg, #4F46E5 0%, #DB2777 50%, #7C3AED 100%);
        -webkit-background-clip: text; -webkit-text-fill-color: transparent;
        margin-bottom: 6px; line-height: 1.3; text-align: center;
    }
    .hero-sub { font-size: 15px; color: #64748B; margin-bottom: 32px; text-align: center; max-width: 600px; margin-left: auto; margin-right: auto; }
    
    .engine-rule {
        background: #F8FAFC; border: 1px solid #E2E8F0;
        border-radius: 10px; padding: 14px 18px; margin: 8px 0; font-size: 13px; color: #475569;
    }
    .engine-rule b { color: #059669; }

    /* Row Explorer styles */
    .explain-panel {
        background: #FFFFFF; border: 1px solid #E2E8F0;
        border-radius: 14px; padding: 20px 24px; margin: 12px 0;
        box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);
    }
    .explain-section {
        margin: 16px 0; padding: 14px 18px;
        background: #F8FAFC; border-radius: 10px;
        border-left: 4px solid #4F46E5;
    }
    .explain-section h4 {
        margin: 0 0 10px 0; font-size: 14px;
        color: #1E293B; font-weight: 700;
    }
    .confidence-badge {
        display: inline-block; padding: 6px 18px;
        border-radius: 20px; font-weight: 700;
        font-size: 14px; letter-spacing: 1px;
    }
    .confidence-HIGH { background: #D1FAE5; color: #065F46; }
    .confidence-MEDIUM { background: #FEF3C7; color: #92400E; }
    .confidence-LOW { background: #FEE2E2; color: #991B1B; }
    .issue-item {
        padding: 8px 12px; margin: 6px 0;
        background: #FFFFFF; border: 1px solid #E2E8F0;
        border-radius: 8px; font-size: 13px;
    }
    .issue-severity {
        display: inline-block; padding: 2px 8px;
        border-radius: 4px; font-size: 10px;
        font-weight: 600; text-transform: uppercase;
        margin-left: 8px;
    }
    .sev-LOW { background: #FEE2E2; color: #991B1B; }
    .sev-MEDIUM { background: #FEF3C7; color: #92400E; }
    .sev-HIGH { background: #D1FAE5; color: #065F46; }
    .raw-data-table {
        width: 100%; border-collapse: collapse;
        font-size: 12px; margin: 8px 0;
    }
    .raw-data-table td {
        padding: 5px 10px; border-bottom: 1px solid #F1F5F9;
    }
    .raw-data-table td:first-child {
        font-weight: 600; color: #64748B;
        width: 140px; white-space: nowrap;
    }
    .raw-data-table td:last-child { color: #1E293B; }
</style>
""", unsafe_allow_html=True)


def render_kpi_row(cards: list):
    """Render KPI cards. cards = [(label, value, sub, color_class)]"""
    html = '<div class="kpi-row">'
    for label, value, sub, cls in cards:
        html += f'''
        <div class="kpi-card">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value {cls}">{value}</div>
            <div class="kpi-sub">{sub}</div>
        </div>'''
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)


def render_results_log(results_log: list):
    """Render the processing results log."""
    if not results_log:
        return
    log_df = pd.DataFrame(results_log)
    status_counts = log_df['status'].value_counts().to_dict()
    cols = st.columns(4)
    cols[0].metric("Success", sum(v for k, v in status_counts.items() if 'OK' in k))
    cols[1].metric("Warnings", sum(v for k, v in status_counts.items() if 'WARN' in k))
    cols[2].metric("Missing", sum(v for k, v in status_counts.items() if 'MISS' in k))
    cols[3].metric("Skipped", sum(v for k, v in status_counts.items() if 'SKIP' in k))
    with st.expander("Detailed Processing Log", expanded=False):
        status_filter = st.multiselect(
            "Filter by status", options=log_df['status'].unique().tolist(),
            default=log_df['status'].unique().tolist(),
        )
        filtered = log_df[log_df['status'].isin(status_filter)]
        st.dataframe(filtered, use_container_width=True, height=400)


# ─────────────────────────────────────────────────────────────────────
# ROW EXPLORER
# ─────────────────────────────────────────────────────────────────────
def _render_row_explorer(output_data: bytes, results_log: list, fill_stats: dict):
    """Render the Row Explorer tab with explain + PDF export."""
    st.markdown('<div class="section-header">Row Explorer — Explain Every Number</div>',
                unsafe_allow_html=True)
    st.caption("Select any processed row to see the full calculation breakdown, "
               "risk report, and confidence score. Export as PDF for auditing.")

    # Build row selector from results_log
    rows_with_data = {}
    for entry in results_log:
        r = entry.get('row')
        if r and r not in rows_with_data:
            rows_with_data[r] = {
                'uid': entry.get('uid', ''),
                'type': entry.get('type', ''),
                'status': entry.get('status', ''),
            }

    if not rows_with_data:
        st.info("No processed rows to explore. Run the engine first.")
        return

    # Pre-compute confidence for all rows (lightweight — from log entries only)
    from engine.explainer import _detect_issues, _compute_confidence
    row_confidence = {}
    for r in rows_with_data:
        log_entries = [e for e in results_log if e.get('row') == r]
        issues = _detect_issues(log_entries, {})
        conf = _compute_confidence(log_entries, issues)
        row_confidence[r] = conf['level']

    # ── Filters Row ──
    conf_colors = {'HIGH': '#059669', 'MEDIUM': '#D97706', 'LOW': '#DC2626'}
    col_filter, col_select = st.columns([1, 3])

    with col_filter:
        available_levels = sorted(set(row_confidence.values()),
                                   key=['HIGH', 'MEDIUM', 'LOW'].index)
        # Build labels with counts
        level_counts = {}
        for lv in available_levels:
            level_counts[lv] = sum(1 for v in row_confidence.values() if v == lv)

        selected_levels = st.multiselect(
            "Filter by Confidence",
            options=available_levels,
            default=available_levels,
            key="confidence_filter",
            help="Filter rows by confidence level to focus on problematic entries"
        )

    # Filter rows by selected confidence levels
    if not selected_levels:
        selected_levels = available_levels

    filtered_rows = {r: info for r, info in rows_with_data.items()
                     if row_confidence.get(r) in selected_levels}

    if not filtered_rows:
        st.warning("No rows match the selected confidence filter.")
        return

    # Show confidence summary badges
    badge_html = '<div style="display:flex; gap:10px; margin:8px 0 16px 0;">'
    for lv in ['HIGH', 'MEDIUM', 'LOW']:
        cnt = level_counts.get(lv, 0)
        if cnt > 0:
            active = lv in selected_levels
            opacity = '1' if active else '0.3'
            badge_html += f'''<div style="background:{conf_colors[lv]}15; border:1px solid {conf_colors[lv]};
                border-radius:8px; padding:6px 14px; opacity:{opacity};">
                <span style="color:{conf_colors[lv]}; font-weight:700; font-size:13px;">{lv}</span>
                <span style="color:#64748B; font-size:12px; margin-left:6px;">{cnt} rows</span>
            </div>'''
    badge_html += '</div>'
    st.markdown(badge_html, unsafe_allow_html=True)

    # Build selectbox options from filtered rows
    row_options = []
    for r in sorted(filtered_rows.keys()):
        info = filtered_rows[r]
        conf_lv = row_confidence[r]
        conf_icon = {'HIGH': '🟢', 'MEDIUM': '🟡', 'LOW': '🔴'}.get(conf_lv, '⚪')
        label = f"{conf_icon} Row {r}  |  {info['uid']}  |  {info['type']}  [{conf_lv}]"
        row_options.append((r, label))

    with col_select:
        selected_label = st.selectbox(
            "Select a row to explain",
            options=[label for _, label in row_options],
            index=0,
            key="row_explorer_select"
        )

    # Find the selected row number
    selected_row = None
    for r, label in row_options:
        if label == selected_label:
            selected_row = r
            break

    if selected_row is None:
        return

    # Get or build explanation (cached in session state)
    cache_key = f"explain_{selected_row}"
    if cache_key not in st.session_state:
        st.session_state[cache_key] = explain_row(output_data, results_log, selected_row)
    expl = st.session_state[cache_key]

    # ── Confidence Badge ──
    conf = expl['confidence']
    st.markdown(f'''
    <div style="margin: 16px 0; display: flex; align-items: center; gap: 16px;">
        <span class="confidence-badge confidence-{conf['level']}">{conf['level']} CONFIDENCE</span>
        <span style="color: #64748B; font-size: 13px;">{conf['reason']}</span>
    </div>
    ''', unsafe_allow_html=True)

    # ── Three columns: Summary | Issues | Export ──
    c1, c2, c3 = st.columns([2, 2, 1])

    with c1:
        st.markdown('<div class="explain-section"><h4>📋 Row Summary</h4>', unsafe_allow_html=True)
        raw = expl['raw_row']
        process = str(expl.get('process', '')).upper().strip()

        summary_items = [
            ('UID', expl.get('uid', '')),
            ('Process', expl.get('process', '')),
            ('Order No', expl.get('order_no', '')),
        ]
        if process == 'PRINTING':
            summary_items.extend([
                ('Material', raw.get('Input Name', '')),
                ('Size', raw.get('Input Size', '')),
                ('Micron', raw.get('Input Mic', '')),
                ('Total Qty', raw.get('Total Input', '')),
                ('Film MR#', raw.get('Film MR#', '')),
                ('Film Rate', f"{raw.get('Film Rate', 0):.4f}" if raw.get('Film Rate') else '—'),
                ('Film Value', f"AED {raw.get('Film Value', 0):,.2f}" if raw.get('Film Value') else '—'),
            ])
        elif process == 'LAM':
            if raw.get('Fresh1 Name'):
                summary_items.extend([
                    ('1st Fresh', raw.get('Fresh1 Name', '')),
                    ('Fresh1 MR#', raw.get('Fresh1 MR#', '')),
                    ('Fresh1 Rate', f"{raw.get('Fresh1 Rate', 0):.4f}" if raw.get('Fresh1 Rate') else '—'),
                ])
            if raw.get('Fresh2 Name'):
                summary_items.extend([
                    ('2nd Fresh', raw.get('Fresh2 Name', '')),
                    ('Fresh2 MR#', raw.get('Fresh2 MR#', '')),
                    ('Fresh2 Rate', f"{raw.get('Fresh2 Rate', 0):.4f}" if raw.get('Fresh2 Rate') else '—'),
                ])
            if raw.get('Adh Name'):
                summary_items.extend([
                    ('Adhesive', raw.get('Adh Name', '')),
                    ('Adh Rate', f"{raw.get('Adh Rate', 0):.4f}" if raw.get('Adh Rate') else '—'),
                    ('Hard Rate', f"{raw.get('Hard Rate', 0):.4f}" if raw.get('Hard Rate') else '—'),
                    ('Sol Rate', f"{raw.get('Sol Rate', 0):.4f}" if raw.get('Sol Rate') else '—'),
                ])

        table_html = '<table class="raw-data-table">'
        for k, v in summary_items:
            table_html += f'<tr><td>{k}</td><td>{v}</td></tr>'
        table_html += '</table>'
        st.markdown(table_html, unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with c2:
        # ── Issues ──
        issues = expl.get('issues', [])
        st.markdown('<div class="explain-section"><h4>⚠️ Mismatch / Risk Report</h4>',
                    unsafe_allow_html=True)
        if issues:
            for issue in issues:
                sev = issue.get('severity', 'MEDIUM')
                st.markdown(f'''
                <div class="issue-item">
                    <strong>{issue['title']}</strong>
                    <span class="issue-severity sev-{sev}">{sev}</span>
                    <div style="color: #64748B; font-size: 12px; margin-top: 4px;">
                        {issue.get('detail', '')}
                    </div>
                </div>
                ''', unsafe_allow_html=True)
        else:
            st.markdown('''
            <div class="issue-item" style="border-color: #D1FAE5; background: #F0FDF4;">
                <strong style="color: #065F46;">✓ No issues detected</strong>
                <div style="color: #047857; font-size: 12px; margin-top: 4px;">
                    All values are direct matches with no adjustments.
                </div>
            </div>
            ''', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with c3:
        # ── PDF Export ──
        st.markdown('<div class="explain-section"><h4>📄 Export</h4>', unsafe_allow_html=True)
        try:
            pdf_bytes = generate_row_pdf(expl)
            uid_safe = str(expl.get('uid', 'row')).replace('-', '_')
            st.download_button(
                label="📥 Export Row Report (PDF)",
                data=pdf_bytes,
                file_name=f"Row_Report_{uid_safe}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key=f"pdf_export_{selected_row}"
            )
        except Exception as e:
            st.error(f"PDF generation error: {e}")
            logger.exception("PDF export failed")
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Rate Breakdown (full width) ──
    st.markdown('<div class="explain-section"><h4>📊 Rate Breakdown</h4>',
                unsafe_allow_html=True)
    breakdown = expl.get('rate_breakdown', [])
    if breakdown:
        bd_df = pd.DataFrame([
            {
                'Type': e.get('type', ''),
                'Status': e.get('status', ''),
                'Detail': e.get('detail', ''),
            }
            for e in breakdown
        ])
        st.dataframe(bd_df, use_container_width=True, hide_index=True)
    else:
        st.info("No rate calculation entries for this row.")
    st.markdown('</div>', unsafe_allow_html=True)

    # ── Raw Excel Data (full width, expandable) ──
    with st.expander("📑 Full Excel Source Row Data", expanded=False):
        raw = expl['raw_row']
        if raw:
            raw_df = pd.DataFrame([
                {'Column': k, 'Value': v}
                for k, v in raw.items()
                if v is not None
            ])
            st.dataframe(raw_df, use_container_width=True, hide_index=True)
        else:
            st.warning("No raw data available for this row.")

    # ── Engine Log Entries (expandable) ──
    with st.expander("🔍 Engine Processing Log Entries", expanded=False):
        entries = expl.get('log_entries', [])
        if entries:
            log_df = pd.DataFrame(entries)
            st.dataframe(log_df, use_container_width=True, hide_index=True)
        else:
            st.info("No engine log entries for this row.")

def _render_chart(fig, key, hint=None):
    """Safely render a chart — skip if fig is None (no data)."""
    if fig is None:
        return False
    st.markdown('<div class="chart-card">', unsafe_allow_html=True)
    if hint:
        st.markdown(f'<div class="chart-hint">{hint}</div>', unsafe_allow_html=True)
    st.plotly_chart(fig, use_container_width=True, key=key)
    st.markdown('</div>', unsafe_allow_html=True)
    return True


# ─────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────
def main():
    # ── Global Sidebar ──
    with st.sidebar:
        col1, col2, col3 = st.columns([1, 4, 1])
        with col2:
            try:
                st.image("assets/ipp-logo.png", width="stretch")
            except Exception:
                pass
        st.markdown("""
        <div style="text-align:center; margin-bottom:20px; margin-top: 10px;">
            <div style="font-size:17px; font-weight:700; color:#1E293B;">IPP Platform</div>
            <div style="font-size:11px; color:#64748B;">Production Automation</div>
        </div>
        """, unsafe_allow_html=True)
        st.markdown("---")

        # RMC Generator file uploads
        rmc_files = render_rmc_sidebar()

        st.markdown("---")
        st.markdown("### 🖨️ Jobtrack Fill (Debug)")
        st.caption("Upload files below for standalone Jobtrack MRR fill.")

        jt_file = st.file_uploader("Jobtrack (Without MRR)", type=['xlsx'], key="jt_upload",
                                   help="The empty Jobtrack template to fill")
        stores_file = st.file_uploader("Stores Recordings", type=['xlsx'], key="stores_upload",
                                       help="Stores material receipt recordings")
        pr_file = st.file_uploader("Purchase Register", type=['xlsx'], key="pr_upload",
                                   help="Purchase register with material rates")

        st.markdown("<div style='font-size:12px; color:#94A3B8; margin: 8px 0 4px;'>Optional: Supplier Rate Files</div>", unsafe_allow_html=True)
        granules_file = st.file_uploader("Granules Recipe (Current Month)", type=['xlsx'], key="granules_upload",
                                         help="For Bandera/CYM supplier rates (by WO#)")
        prev_granules_file = st.file_uploader("Granules Recipe (Previous Month)", type=['xlsx'], key="prev_granules_upload",
                                              help="Fallback: used when WO# not found in current month")
        megapack_file = st.file_uploader("MEGA PACK", type=['xlsx'], key="megapack_upload",
                                         help="For Mega Pack supplier rates (TPE/WPE)")

        st.markdown("---")
        all_uploaded = all([jt_file, stores_file, pr_file])
        process_btn = st.button("Process & Fill", disabled=not all_uploaded, use_container_width=True)

        if not all_uploaded:
            missing = []
            if not jt_file: missing.append("Jobtrack")
            if not stores_file: missing.append("Stores")
            if not pr_file: missing.append("Purchase Register")
            st.warning(f"Missing: {', '.join(missing)}")

        st.markdown("---")

        # Engine rules
        with st.expander("Engine Rules (No Special Cases)", expanded=False):
            st.markdown("""
            <div class="engine-rule">
                <b>Rule 1: Film/Fresh Rate</b><br>
                Qty-weighted avg = SUM(PR_rate x Stores_qty) / SUM(Stores_qty)
            </div>
            <div class="engine-rule">
                <b>Rule 2: Chemical Rates</b><br>
                Total Amount / Total Qty from PR entries in the <b>same month</b>
            </div>
            <div class="engine-rule">
                <b>Rule 3: MRR Discovery</b><br>
                Stores: Order + Process + Material + Mic (no width filter)
            </div>
            <div class="engine-rule">
                <b>Rule 4: PR Validation</b><br>
                Only MRRs present in Purchase Register are used for rates
            </div>
            """, unsafe_allow_html=True)

        st.markdown("""
        <div style="font-size:11px; color:#475569; text-align:center; margin-top:16px;">
            <div>IPP Production Platform</div>
        </div>
        """, unsafe_allow_html=True)

    # ── Main Content: Two Tabs ──
    main_tab1, main_tab2 = st.tabs(["📊 RMC Generator (Main)", "🖨️ Jobtrack Fill (Debug)"])

    with main_tab2:
        # ── Jobtrack Fill Tab ──
        st.markdown("""
        <div class="hero-title">Jobtrack MRR Auto-Fill</div>
        <div class="hero-sub">Debug / Fallback: Upload your monthly files to auto-fill MR#, Rate, and Value.</div>
        """, unsafe_allow_html=True)

        # -- Process --
        if process_btn and all_uploaded:
            st.markdown('<div class="section-header">Processing</div>', unsafe_allow_html=True)
            progress_bar = st.progress(0)
            status_text = st.empty()

            def progress_cb(pct, msg):
                progress_bar.progress(min(pct, 100))
                status_text.text(msg)

            try:
                jt_bytes = io.BytesIO(jt_file.read())
                stores_bytes = io.BytesIO(stores_file.read())
                pr_bytes = io.BytesIO(pr_file.read())
                granules_bytes = io.BytesIO(granules_file.read()) if granules_file else None
                prev_granules_bytes = io.BytesIO(prev_granules_file.read()) if prev_granules_file else None
                megapack_bytes = io.BytesIO(megapack_file.read()) if megapack_file else None
                output_stream, results_log, fill_stats = fill_jobtrack(
                    jt_bytes, stores_bytes, pr_bytes, progress_callback=progress_cb,
                    granules_file=granules_bytes, megapack_file=megapack_bytes,
                    prev_granules_file=prev_granules_bytes
                )
                # Store the raw bytes (not the stream) to avoid seek/position bugs
                output_stream.seek(0)
                st.session_state['output_data'] = output_stream.read()
                st.session_state['results_log'] = results_log
                st.session_state['fill_stats'] = fill_stats
                st.session_state['processed'] = True
                # Pre-parse dashboard DataFrame once (avoid re-parsing per tab)
                st.session_state['filled_df'] = get_filled_data_for_dashboard(
                    io.BytesIO(st.session_state['output_data'])
                )
                status_text.text("Processing complete!")
                st.success(f"Successfully processed! {fill_stats['film_filled']} film, "
                           f"{fill_stats['fresh1_filled']} fresh-1, {fill_stats['fresh2_filled']} fresh-2, "
                           f"{fill_stats['adh_filled']} adhesive, {fill_stats['hard_filled']} hardener, "
                           f"{fill_stats['sol_filled']} solvent rows filled.")
            except Exception as e:
                st.error(f"Error during processing: {str(e)}")
                logger.exception("Processing failed")
                st.session_state['processed'] = False

        # -- Results --
        if st.session_state.get('processed'):
            output_data = st.session_state['output_data']
            results_log = st.session_state['results_log']
            fill_stats = st.session_state['fill_stats']
            filled_df = st.session_state['filled_df']

            st.markdown("---")
            now = datetime.now().strftime("%Y%m%d_%H%M")
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                st.download_button(
                    label="Download Filled Jobtrack",
                    data=output_data,
                    file_name=f"Jobtrack_Filled_MRR_{now}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            # ── Manual Review Alert ──
            warn_rows = [e for e in results_log if e.get('status') in ('WARN', 'MISS', 'ERROR')]
            if warn_rows:
                st.markdown("---")
                with st.expander(f"⚠️ **Manual Review Required** — {len(warn_rows)} items need attention", expanded=True):
                    st.markdown("""
                    <div style="background: linear-gradient(135deg, #1a1a2e 0%, #2d1b3d 100%); 
                         border-radius: 12px; padding: 16px; margin-bottom: 12px;
                         border-left: 4px solid #FF6B6B;">
                        <p style="color: #FF6B6B; font-weight: 600; margin-bottom: 8px;">
                            🔴 These rows are highlighted RED in the downloaded Excel file</p>
                        <p style="color: #94A3B8; font-size: 13px; margin: 0;">
                            <b>Option 1:</b> Download the Excel, fill the highlighted rows manually, then use as final output<br>
                            <b>Option 2:</b> Accept the engine's calculated values (weighted average across all matching MRRs)<br>
                            <b>Option 3:</b> Provide missing data files (e.g., previous month Granules) and re-process
                        </p>
                    </div>
                    """, unsafe_allow_html=True)
                
                    # Group by row
                    from collections import defaultdict
                    row_issues = defaultdict(list)
                    for e in warn_rows:
                        row_issues[e.get('row', '?')].append(e)
                
                    for row_num in sorted(row_issues.keys()):
                        issues = row_issues[row_num]
                        uid = issues[0].get('uid', '')
                        types = ', '.join(e.get('type', '') for e in issues)
                        details = ' | '.join(e.get('detail', '') for e in issues)
                        status = issues[0].get('status', 'WARN')
                    
                        icon = "🔴" if status == "MISS" else "🟡"
                        st.markdown(f"""
                        <div style="background: #1E293B; border-radius: 8px; padding: 10px 14px; 
                             margin: 4px 0; border-left: 3px solid {'#FF6B6B' if status == 'MISS' else '#FBBF24'};">
                            <span style="color: #E2E8F0; font-weight: 600;">
                                {icon} Row {row_num}</span>
                            <span style="color: #64748B; margin-left: 12px;">{uid}</span>
                            <span style="color: #94A3B8; margin-left: 12px;">[{types}]</span>
                            <br><span style="color: #78716C; font-size: 12px;">{details}</span>
                        </div>
                        """, unsafe_allow_html=True)

            st.markdown("---")
            tab1, tab2, tab3, tab4, tab5 = st.tabs([
                "CEO Overview", "Cost Intelligence",
                "Production Analytics", "Processing Log", "Row Explorer"
            ])

            with tab1:
                try:
                    ds = compute_dashboard_stats(filled_df, fill_stats)

                    # ── Executive KPIs ──
                    st.markdown('<div class="section-header">Executive Summary</div>', unsafe_allow_html=True)
                    render_kpi_row([
                        ("Total Material Cost", f"AED {ds.get('total_material_cost',0):,.0f}",
                         f"{ds.get('total_rows',0)} production rows", ""),
                        ("Net Output", f"{ds.get('total_output_kg',0):,.0f} Kg",
                         "Total production weight", "green"),
                        ("Cost / Kg Output", f"AED {ds.get('cost_per_kg',0):.2f}",
                         "Material efficiency metric", "blue"),
                        ("Unique Orders", f"{ds.get('unique_orders',0)}",
                         "Distinct production orders", "amber"),
                        ("Wastage Rate", f"{ds.get('waste_pct',0):.1f}%",
                         f"{ds.get('total_waste_kg',0):,.0f} Kg total waste", "amber"),
                    ])

                    # ── AI Insights Panel ──
                    st.markdown('<div class="section-header">Business Insights</div>', unsafe_allow_html=True)
                    insights = generate_ai_insights(filled_df, ds)
                    if insights:
                        cols_per_row = 3
                        for i in range(0, len(insights), cols_per_row):
                            cols = st.columns(cols_per_row)
                            for j, col in enumerate(cols):
                                idx = i + j
                                if idx < len(insights):
                                    ins = insights[idx]
                                    type_colors = {
                                        'cost': '#6366F1', 'efficiency': '#3B82F6',
                                        'alert': '#EF4444', 'success': '#10B981',
                                        'order': '#F59E0B', 'process': '#8B5CF6',
                                        'automation': '#14B8A6',
                                    }
                                    border_color = type_colors.get(ins['type'], '#6366F1')
                                    with col:
                                        st.markdown(f'''
                                        <div style="background:#FFFFFF; border:1px solid #E2E8F0;
                                             border-left:4px solid {border_color}; border-radius:12px;
                                             padding:16px 20px; height:100%; min-height:120px;
                                             box-shadow: 0 2px 8px rgba(0,0,0,0.04);">
                                            <div style="font-size:24px; margin-bottom:6px;">{ins['icon']}</div>
                                            <div style="font-size:14px; font-weight:700; color:#1E293B; margin-bottom:6px;">
                                                {ins['title']}</div>
                                            <div style="font-size:12px; color:#64748B; line-height:1.5;">
                                                {ins['detail']}</div>
                                        </div>
                                        ''', unsafe_allow_html=True)

                    # ── Recommendations Panel ──
                    st.markdown('<div class="section-header">Action Required</div>', unsafe_allow_html=True)
                    recs = generate_recommendations(filled_df, ds)
                    if recs:
                        for rec in recs:
                            pri_colors = {'HIGH': '#EF4444', 'MEDIUM': '#F59E0B', 'LOW': '#10B981'}
                            pri_bg = {'HIGH': '#FEF2F2', 'MEDIUM': '#FFFBEB', 'LOW': '#F0FDF4'}
                            pri = rec['priority']
                            st.markdown(f'''
                            <div style="background:{pri_bg.get(pri,'#F8FAFC')}; border:1px solid #E2E8F0;
                                 border-radius:12px; padding:16px 20px; margin-bottom:10px;
                                 box-shadow: 0 2px 8px rgba(0,0,0,0.04);">
                                <div style="display:flex; align-items:center; gap:12px; margin-bottom:8px;">
                                    <span style="background:{pri_colors.get(pri,'#6366F1')}; color:white;
                                         padding:3px 10px; border-radius:6px; font-size:10px;
                                         font-weight:700; letter-spacing:1px;">{pri}</span>
                                    <span style="background:#E0E7FF; color:#4338CA; padding:3px 10px;
                                         border-radius:6px; font-size:10px; font-weight:600;">{rec['category']}</span>
                                </div>
                                <div style="font-size:14px; font-weight:700; color:#1E293B; margin-bottom:4px;">
                                    {rec['action']}</div>
                                <div style="font-size:12px; color:#64748B;">{rec['detail']}</div>
                                <div style="font-size:11px; color:#059669; margin-top:6px; font-weight:600;">
                                    Impact: {rec['impact']}</div>
                            </div>
                            ''', unsafe_allow_html=True)
                    else:
                        st.success("No critical action items — all metrics within expected ranges.")

                    # ── Process Detail for CEO ──
                    st.markdown('<div class="section-header">Process Overview</div>', unsafe_allow_html=True)
                    p_rows = ds.get('printing_rows', 0)
                    l_rows = ds.get('lam_rows', 0)
                    film_cost = ds.get('total_film_cost', 0)
                    lam_cost = sum(ds.get(k, 0) for k in ['total_fresh1_cost', 'total_fresh2_cost',
                                                            'total_adh_cost', 'total_hard_cost', 'total_sol_cost'])
                    total = ds.get('total_material_cost', 0)
                    film_pct = film_cost / total * 100 if total > 0 else 0
                    lam_pct = lam_cost / total * 100 if total > 0 else 0

                    col1, col2 = st.columns(2)
                    with col1:
                        st.markdown(f'''
                        <div style="background:linear-gradient(135deg, #EEF2FF 0%, #E0E7FF 100%);
                             border-radius:16px; padding:24px; border:1px solid #C7D2FE;">
                            <div style="display:flex; align-items:center; gap:12px; margin-bottom:16px;">
                                <div style="background:#6366F1; color:white; width:40px; height:40px;
                                     border-radius:10px; display:flex; align-items:center; justify-content:center;
                                     font-size:20px;">🖨</div>
                                <div>
                                    <div style="font-size:18px; font-weight:700; color:#312E81;">PRINTING</div>
                                    <div style="font-size:12px; color:#6366F1;">{p_rows} production rows</div>
                                </div>
                            </div>
                            <div style="font-size:28px; font-weight:800; color:#312E81;">AED {film_cost:,.0f}</div>
                            <div style="font-size:12px; color:#4F46E5; margin-top:4px;">{film_pct:.0f}% of total spend</div>
                            <div style="background:#C7D2FE; height:6px; border-radius:3px; margin-top:12px;">
                                <div style="background:#6366F1; height:6px; border-radius:3px; width:{min(film_pct, 100):.0f}%;"></div>
                            </div>
                            <div style="margin-top:14px; font-size:12px; color:#64748B; line-height:1.6;">
                                <b>What it covers:</b> Film material (PET, BOPP, NYLON, etc.) used as the primary input in printing jobs.
                                The rate is a qty-weighted average from Purchase Register matched via Stores MRR numbers.
                            </div>
                        </div>
                        ''', unsafe_allow_html=True)
                    with col2:
                        st.markdown(f'''
                        <div style="background:linear-gradient(135deg, #FDF2F8 0%, #FCE7F3 100%);
                             border-radius:16px; padding:24px; border:1px solid #F9A8D4;">
                            <div style="display:flex; align-items:center; gap:12px; margin-bottom:16px;">
                                <div style="background:#EC4899; color:white; width:40px; height:40px;
                                     border-radius:10px; display:flex; align-items:center; justify-content:center;
                                     font-size:20px;">🔬</div>
                                <div>
                                    <div style="font-size:18px; font-weight:700; color:#831843;">LAMINATION</div>
                                    <div style="font-size:12px; color:#EC4899;">{l_rows} production rows</div>
                                </div>
                            </div>
                            <div style="font-size:28px; font-weight:800; color:#831843;">AED {lam_cost:,.0f}</div>
                            <div style="font-size:12px; color:#DB2777; margin-top:4px;">{lam_pct:.0f}% of total spend</div>
                            <div style="background:#F9A8D4; height:6px; border-radius:3px; margin-top:12px;">
                                <div style="background:#EC4899; height:6px; border-radius:3px; width:{min(lam_pct, 100):.0f}%;"></div>
                            </div>
                            <div style="margin-top:14px; font-size:12px; color:#64748B; line-height:1.6;">
                                <b>What it covers:</b> Fresh materials (1st & 2nd layer), Adhesive, Hardener, and Solvent
                                used in lamination. Includes 5 sub-costs calculated from PR and Stores data.
                            </div>
                        </div>
                        ''', unsafe_allow_html=True)

                    # ── Charts (only show if data exists) ──
                    col1, col2 = st.columns(2)
                    with col1:
                        _render_chart(create_daily_cost_trend(filled_df), "ceo_trend",
                                      "Daily cost with 3-day moving average.")
                    with col2:
                        _render_chart(create_coverage_gauge(ds), "ceo_gauge",
                                      "Auto-fill coverage. Green zone (>80%) = production-ready.")

                except Exception as e:
                    st.error(f"Dashboard error: {str(e)}")
                    logger.exception("CEO Overview failed")

            with tab2:
                try:
                    ds = compute_dashboard_stats(filled_df, fill_stats)

                    # ── Process KPIs ──
                    st.markdown('<div class="section-header">Cost Breakdown by Category</div>', unsafe_allow_html=True)
                    film_cost = ds.get('total_film_cost', 0)
                    render_kpi_row([
                        ("Printing (Film)", f"AED {film_cost:,.0f}",
                         f"{ds.get('film_filled',0)} rows filled", ""),
                        ("1st Fresh Material", f"AED {ds.get('total_fresh1_cost',0):,.0f}",
                         f"{ds.get('fresh1_filled',0)} rows", ""),
                        ("2nd Fresh Material", f"AED {ds.get('total_fresh2_cost',0):,.0f}",
                         f"{ds.get('fresh2_filled',0)} rows", ""),
                        ("Adhesive", f"AED {ds.get('total_adh_cost',0):,.0f}",
                         f"Rate: {ds.get('adh_rate',0):.3f}/Kg", "green"),
                        ("Hardener", f"AED {ds.get('total_hard_cost',0):,.0f}",
                         f"Rate: {ds.get('hard_rate',0):.3f}/Kg", "green"),
                        ("Solvent (E/A)", f"AED {ds.get('total_sol_cost',0):,.0f}",
                         f"Rate: {ds.get('sol_rate',0):.3f}/Kg", "green"),
                    ])

                    # ── Waterfall + Donut ──
                    st.markdown('<div class="section-header">Cost Analysis</div>', unsafe_allow_html=True)
                    col1, col2 = st.columns(2)
                    with col1:
                        _render_chart(create_material_cost_waterfall(ds), "waterfall",
                                      "How each material category builds up to total cost.")
                    with col2:
                        _render_chart(create_cost_breakdown_pie(ds), "ci_pie",
                                      "Each slice = share of total material spend.")

                    # ── Pareto + Process Bar ──
                    col1, col2 = st.columns(2)
                    with col1:
                        _render_chart(create_cost_pareto(filled_df), "pareto",
                                      "Which 20% of orders drive 80% of cost.")
                    with col2:
                        _render_chart(create_process_cost_bar(filled_df), "ci_bar",
                                      "Stacked cost by process.")

                    # ── Top Orders + Chemical Rates ──
                    col1, col2 = st.columns([2, 1])
                    with col1:
                        _render_chart(create_top_orders_chart(filled_df), "ci_top",
                                      "Longest bar = highest cost order.")
                    with col2:
                        _render_chart(create_chemical_rates_summary(ds), "ci_chem",
                                      "Current month chemical procurement rates.")

                except Exception as e:
                    st.error(f"Cost Intelligence error: {str(e)}")
                    logger.exception("Cost Intelligence failed")

            with tab3:
                try:
                    ds = compute_dashboard_stats(filled_df, fill_stats)

                    st.markdown('<div class="section-header">Production Efficiency</div>', unsafe_allow_html=True)
                    col1, col2 = st.columns(2)
                    with col1:
                        _render_chart(create_efficiency_scatter(filled_df), "scatter",
                                      "Each bubble = 1 order. Green = efficient, Red = expensive.")
                    with col2:
                        _render_chart(create_process_efficiency_radar(filled_df, ds), "radar",
                                      "Printing vs LAM across 5 dimensions.")

                    st.markdown('<div class="section-header">Rate & Material Analysis</div>', unsafe_allow_html=True)
                    col1, col2 = st.columns(2)
                    with col1:
                        _render_chart(create_cost_per_kg_chart(filled_df), "pa_cpk",
                                      "Higher bar = more expensive per Kg output.")
                    with col2:
                        _render_chart(create_rate_comparison_chart(filled_df), "pa_rates",
                                      "Wider box = more rate variation. Dots = outliers.")

                    st.markdown('<div class="section-header">Materials & Wastage</div>', unsafe_allow_html=True)
                    col1, col2 = st.columns(2)
                    with col1:
                        _render_chart(create_material_usage_chart(filled_df), "pa_mats",
                                      "Most-consumed materials for procurement planning.")
                    with col2:
                        _render_chart(create_wastage_chart(filled_df), "pa_waste",
                                      "Waste per process. Compare to 5% target.")

                except Exception as e:
                    st.error(f"Analytics error: {str(e)}")
                    logger.exception("Production Analytics failed")

            with tab4:
                st.markdown('<div class="section-header">Processing Results</div>', unsafe_allow_html=True)
                render_results_log(results_log)

            with tab5:
                _render_row_explorer(output_data, results_log, fill_stats)

        else:
            # Welcome state
            st.markdown("---")
            col1, col2, col3 = st.columns(3)
            steps = [
                ("1", "Upload Jobtrack", "Upload the monthly Jobtrack file (without MRR data)"),
                ("2", "Upload Stores & PR", "Upload Stores Recordings and Purchase Register"),
                ("3", "Process & Download", "Click Process to auto-fill and download the result"),
            ]
            for col, (num, title, desc) in zip([col1, col2, col3], steps):
                with col:
                    st.markdown(f"""
                    <div class="kpi-card">
                        <div class="kpi-label">Step {num}</div>
                        <div style="color:#1E293B; font-size:15px; font-weight:600; margin-bottom:8px;">{title}</div>
                        <div style="color:#94A3B8; font-size:13px;">{desc}</div>
                    </div>
                    """, unsafe_allow_html=True)

            st.markdown("")
            st.info("Upload your 3 files in the sidebar to get started!")

            # Engine transparency
            st.markdown('<div class="section-header">How The Engine Works</div>', unsafe_allow_html=True)
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("""
                <div class="engine-rule">
                    <b>Film / Fresh Material Rate</b><br>
                    1. Find all MRRs from Stores: Order + Process + Material + Mic<br>
                    2. Validate each MRR exists in Purchase Register<br>
                    3. Rate = SUM(PR_rate x Stores_qty) / SUM(Stores_qty)<br>
                    4. Value = Total Input Qty x Rate
                </div>
                <div class="engine-rule">
                    <b>Chemical Rates (Adhesive, Hardener, Solvent)</b><br>
                    1. Filter Purchase Register by the reporting month<br>
                    2. Rate = Total Amount / Total Quantity for that month<br>
                    3. Value = Kgs Used x Rate
                </div>
                """, unsafe_allow_html=True)
            with col2:
                st.markdown("""
                <div class="engine-rule">
                    <b>MRR Discovery (Stores Recordings)</b><br>
                    - Match: Order No + Process + Material Name + Mic<br>
                    - No width/size filter (all rolls of same material included)<br>
                    - Only MRRs with positive issue qty are counted
                </div>
                <div class="engine-rule">
                    <b>PR Validation</b><br>
                    - Each MRR from Stores is checked against Purchase Register<br>
                    - MRRs not in PR are excluded from rate calculation<br>
                    - Ensures only real, verified receipts affect cost
                </div>
                """, unsafe_allow_html=True)

    with main_tab1:
        render_rmc_tab(rmc_files)


if __name__ == "__main__":
    main()
