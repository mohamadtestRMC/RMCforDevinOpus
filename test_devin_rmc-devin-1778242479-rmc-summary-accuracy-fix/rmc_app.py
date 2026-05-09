"""
RMC App — Enhanced Streamlit application with Executive Dashboard,
Audit Panel, Flag Checker, and Cross-Month Validation tabs.
Premium dark theme with glassmorphism design.
"""
import io
import sys
import logging
from pathlib import Path
from datetime import datetime

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

logger = logging.getLogger(__name__)
_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(_ROOT))

COLORS = ["#6366F1","#8B5CF6","#3B82F6","#10B981","#F59E0B",
          "#EF4444","#06B6D4","#EC4899","#F97316","#84CC16"]

# ── Premium CSS ──
CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800&display=swap');
:root {
    --bg-primary: #0F172A; --bg-card: rgba(30,41,59,0.7);
    --border: rgba(99,102,241,0.2); --text-primary: #F1F5F9;
    --text-secondary: #94A3B8; --accent: #6366F1;
    --accent2: #8B5CF6; --success: #10B981;
    --warning: #F59E0B; --danger: #EF4444;
}
.stApp { font-family: 'Inter', sans-serif; }
.kpi-grid { display:grid; grid-template-columns:repeat(auto-fit,minmax(180px,1fr));
    gap:16px; margin:20px 0; }
.kpi-card { background:var(--bg-card); border:1px solid var(--border);
    border-radius:16px; padding:20px; text-align:center;
    backdrop-filter:blur(12px); transition:all 0.3s ease; }
.kpi-card:hover { border-color:var(--accent); transform:translateY(-2px);
    box-shadow:0 8px 32px rgba(99,102,241,0.15); }
.kpi-label { font-size:12px; color:var(--text-secondary);
    text-transform:uppercase; letter-spacing:1px; margin-bottom:8px; }
.kpi-value { font-size:28px; font-weight:800;
    background:linear-gradient(135deg,#6366F1,#8B5CF6);
    -webkit-background-clip:text; -webkit-text-fill-color:transparent; }
.kpi-value.green { background:linear-gradient(135deg,#10B981,#34D399);
    -webkit-background-clip:text; -webkit-text-fill-color:transparent; }
.kpi-value.red { background:linear-gradient(135deg,#EF4444,#F87171);
    -webkit-background-clip:text; -webkit-text-fill-color:transparent; }
.kpi-value.blue { background:linear-gradient(135deg,#3B82F6,#60A5FA);
    -webkit-background-clip:text; -webkit-text-fill-color:transparent; }
.kpi-sub { font-size:11px; color:var(--text-secondary); margin-top:4px; }
.hero-title { font-size:32px; font-weight:800; text-align:center;
    background:linear-gradient(135deg,#6366F1 0%,#EC4899 50%,#8B5CF6 100%);
    -webkit-background-clip:text; -webkit-text-fill-color:transparent; margin-bottom:4px; }
.hero-sub { font-size:14px; color:#94A3B8; text-align:center; margin-bottom:24px; }
.flag-high { color:#EF4444; font-weight:700; }
.flag-medium { color:#F59E0B; font-weight:600; }
.flag-low { color:#10B981; }
.insight-card { background:var(--bg-card); border:1px solid var(--border);
    border-radius:12px; padding:16px; margin:8px 0;
    backdrop-filter:blur(12px); }
.rec-card { background:var(--bg-card); border-left:4px solid var(--accent);
    border-radius:0 12px 12px 0; padding:16px; margin:8px 0; }
.rec-card.high { border-left-color:#EF4444; }
.rec-card.medium { border-left-color:#F59E0B; }
.section-header { font-size:18px; font-weight:700; margin:24px 0 12px;
    padding-bottom:8px; border-bottom:1px solid var(--border); }
</style>
"""


def _sf(v):
    try:
        import math
        f = float(v) if v is not None else 0.0
        return 0.0 if math.isnan(f) or math.isinf(f) else f
    except: return 0.0


def main():
    st.set_page_config(page_title="RMC Pipeline", page_icon="🏭",
                       layout="wide", initial_sidebar_state="expanded")
    st.markdown(CUSTOM_CSS, unsafe_allow_html=True)
    st.markdown('<div class="hero-title">🏭 RMC Pipeline Engine</div>', unsafe_allow_html=True)
    st.markdown('<div class="hero-sub">Manufacturing Raw Material Cost Analysis & Reporting</div>',
                unsafe_allow_html=True)

    # Sidebar
    with st.sidebar:
        st.markdown("### 📂 Source Files")
        files = _render_sidebar()

    # Main tabs
    if st.session_state.get("rmc_complete"):
        tabs = st.tabs(["🚀 Pipeline", "📊 Dashboard", "🔍 Audit",
                        "⚠️ Flags", "📆 Cross-Month"])
        with tabs[0]: _render_pipeline_tab(files)
        with tabs[1]: _render_dashboard_tab()
        with tabs[2]: _render_audit_tab()
        with tabs[3]: _render_flags_tab()
        with tabs[4]: _render_cross_month_tab()
    else:
        _render_pipeline_tab(files)


def _render_sidebar():
    jt = st.file_uploader("Job Track (Raw)", type=['xlsx'], key="rmc_jt")
    stores = st.file_uploader("Stores Recordings", type=['xlsx'], key="rmc_stores")
    pr = st.file_uploader("Purchase Register", type=['xlsx'], key="rmc_pr")
    template = st.file_uploader("Base RMC Template", type=['xlsx'], key="rmc_template")
    st.markdown("**Optional:**")
    gran = st.file_uploader("Granules Recipe", type=['xlsx'], key="rmc_gran")
    pgran = st.file_uploader("Prev Granules", type=['xlsx'], key="rmc_pgran")
    mega = st.file_uploader("MEGA PACK Rates", type=['xlsx'], key="rmc_mega")
    ink = st.file_uploader("Ink Consumption", type=['xlsx'], key="rmc_ink")
    comp = st.file_uploader("Components", type=['xlsx'], key="rmc_comp")
    opn = st.file_uploader("Opening WIP", type=['xlsx'], key="rmc_opnwip")
    cls = st.file_uploader("Closing WIP", type=['xlsx'], key="rmc_clswip")
    ref = st.file_uploader("Filled Reference", type=['xlsx'], key="rmc_ref")
    st.markdown("**Cross-Month:**")
    prev = st.file_uploader("Previous Month RMC", type=['xlsx'], key="rmc_prev_month")
    return {"jt":jt,"stores":stores,"pr":pr,"template":template,"granules":gran,
            "prev_granules":pgran,"megapack":mega,"ink":ink,"components":comp,
            "opn_wip":opn,"cls_wip":cls,"filled_ref":ref,"prev_month":prev}


def _render_pipeline_tab(files):
    required = ["jt","stores","pr"]
    missing = [k for k in required if files.get(k) is None]

    # File status badges
    labels = {"jt":"Job Track","stores":"Stores","pr":"Purchase Register",
              "template":"Template","granules":"Granules","megapack":"MEGA PACK",
              "ink":"Ink","components":"Components","opn_wip":"OPN WIP",
              "cls_wip":"CLS WIP","filled_ref":"Reference"}
    req_set = {"jt","stores","pr"}
    html = '<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:16px;">'
    for k,label in labels.items():
        ok = files.get(k) is not None
        if ok: bg,color,icon = "#ECFDF5","#059669","✅"
        elif k in req_set: bg,color,icon = "#FEF2F2","#DC2626","❌"
        else: bg,color,icon = "#F8FAFC","#94A3B8","○"
        html += f'<span style="background:{bg};color:{color};padding:4px 10px;border-radius:6px;font-size:11px;font-weight:600">{icon} {label}</span>'
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)

    if missing:
        st.warning(f"Required files missing: **{', '.join(missing)}**")
        return

    c1,c2,c3 = st.columns([1,2,1])
    with c2:
        run = st.button("🚀 Generate RMC Report", type="primary",
                        use_container_width=True, key="rmc_run")

    # Download buttons
    if st.session_state.get("rmc_output_bytes"):
        c1,c2,c3 = st.columns([1,2,1])
        now = datetime.now().strftime("%Y%m%d_%H%M")
        with c2:
            st.download_button("📥 Download RMC Report",
                data=st.session_state["rmc_output_bytes"],
                file_name=f"RMC_Output_{now}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, key="dl_rmc")
        if st.session_state.get("rmc_filled_template"):
            with c2:
                st.download_button("📥 Download Filled Base RMC",
                    data=st.session_state["rmc_filled_template"],
                    file_name=f"Base_RMC_Filled_{now}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True, key="dl_tmpl")

    if run:
        _run_pipeline(files)

    if st.session_state.get("rmc_complete"):
        _render_pipeline_results()


def _run_pipeline(files):
    bar = st.progress(0)
    status = st.empty()
    logs = []
    def cb(pct,msg):
        bar.progress(min(pct/100,1.0))
        status.markdown(f"**{msg}**")
        logs.append(f"[{pct:3d}%] {msg}")
    try:
        from engine.rmc_pipeline import UnifiedRMCPipeline
        def _b(f): return io.BytesIO(f.read()) if f else None
        pipe = UnifiedRMCPipeline()
        result = pipe.run(
            jt_file=_b(files["jt"]), stores_file=_b(files["stores"]),
            pr_file=_b(files["pr"]),
            granules_file=_b(files.get("granules")),
            megapack_file=_b(files.get("megapack")),
            prev_granules_file=_b(files.get("prev_granules")),
            base_rmc_template=_b(files.get("template")),
            opening_wip=_b(files.get("opn_wip")),
            closing_wip=_b(files.get("cls_wip")),
            ink_consumption=_b(files.get("ink")),
            components=_b(files.get("components")),
            filled_rmc_reference=_b(files.get("filled_ref")),
            progress_cb=cb,
        )
        for k,v in result.items():
            st.session_state[f"rmc_{k}"] = v
        st.session_state["rmc_output_bytes"] = result.get("output_bytes")
        st.session_state["rmc_filled_template"] = result.get("filled_template_bytes")
        st.session_state["rmc_rows"] = result.get("rmc_rows", [])
        st.session_state["rmc_metrics"] = result.get("metrics", {})
        st.session_state["rmc_log"] = result.get("log", [])
        st.session_state["rmc_complete"] = True
        bar.progress(1.0)
        status.markdown("**✅ Pipeline complete!**")
        m = result.get("metrics",{})
        acc = m.get("accuracy_pct",-1)
        msg = f"Generated {len(result.get('rmc_rows',[]))} orders"
        if acc >= 0: msg += f" | Accuracy: {acc:.1f}%"
        msg += f" | Time: {m.get('elapsed_seconds',0):.1f}s"
        st.success(msg)
    except Exception as e:
        st.error(f"Pipeline error: {e}")
        import traceback
        st.code(traceback.format_exc())


def _render_pipeline_results():
    m = st.session_state.get("rmc_metrics",{})
    acc = m.get("accuracy_pct",-1)
    orders = m.get("rmc_summary_orders", len(st.session_state.get("rmc_rows",[])))
    exact = m.get("exact_matches",0)
    total_chk = m.get("total_checks",0)
    mismatch = m.get("mismatches_gt1",0)
    elapsed = m.get("elapsed_seconds",0)

    acc_cls = "green" if acc>=99.5 else ("" if acc>=95 else "red")
    acc_lbl = f"{acc:.1f}%" if acc>=0 else "N/A"

    st.markdown(f"""<div class="kpi-grid">
        <div class="kpi-card"><div class="kpi-label">Orders</div>
            <div class="kpi-value blue">{orders:,}</div></div>
        <div class="kpi-card"><div class="kpi-label">Accuracy</div>
            <div class="kpi-value {acc_cls}">{acc_lbl}</div></div>
        <div class="kpi-card"><div class="kpi-label">Exact Matches</div>
            <div class="kpi-value green">{exact:,}/{total_chk:,}</div></div>
        <div class="kpi-card"><div class="kpi-label">Mismatches (&gt;1)</div>
            <div class="kpi-value {'red' if mismatch>0 else 'green'}">{mismatch}</div></div>
        <div class="kpi-card"><div class="kpi-label">Time</div>
            <div class="kpi-value blue">{elapsed:.1f}s</div></div>
    </div>""", unsafe_allow_html=True)

    log = st.session_state.get("rmc_log",[])
    if log:
        with st.expander("📝 Pipeline Log", expanded=False):
            st.code("\n".join(log[-30:]), language="text")


def _render_dashboard_tab():
    """Executive Dashboard with KPIs and charts."""
    from dashboard.executive_dashboard import (
        compute_kpis, compute_cost_breakdown, compute_top_orders,
        compute_rmc_distribution, generate_insights, generate_recommendations
    )
    from dashboard.cross_month import load_rmc_summary

    rmc_bytes = st.session_state.get("rmc_filled_template") or st.session_state.get("rmc_output_bytes")
    if not rmc_bytes:
        st.info("Run the pipeline first to see dashboard.")
        return

    df = load_rmc_summary(rmc_bytes, "Current")
    if df.empty:
        st.warning("No RMC Summary data found.")
        return

    kpis = compute_kpis(df)
    tc = kpis.get('total_cost',0)
    to = kpis.get('total_output',0)

    # KPI cards
    st.markdown(f"""<div class="kpi-grid">
        <div class="kpi-card"><div class="kpi-label">Total Material Cost</div>
            <div class="kpi-value">AED {tc:,.0f}</div>
            <div class="kpi-sub">{kpis['order_count']} orders</div></div>
        <div class="kpi-card"><div class="kpi-label">Total Output</div>
            <div class="kpi-value blue">{to:,.0f} Kg</div></div>
        <div class="kpi-card"><div class="kpi-label">Avg RMC/Kg</div>
            <div class="kpi-value green">AED {kpis['avg_rmc']:,.2f}</div></div>
        <div class="kpi-card"><div class="kpi-label">Coverage</div>
            <div class="kpi-value {'green' if kpis['coverage_pct']>95 else 'red'}">{kpis['coverage_pct']:.1f}%</div></div>
        <div class="kpi-card"><div class="kpi-label">Wastage</div>
            <div class="kpi-value {'red' if kpis['waste_pct']>8 else 'green'}">{kpis['waste_pct']:.1f}%</div></div>
    </div>""", unsafe_allow_html=True)

    # Charts
    c1,c2 = st.columns(2)
    with c1:
        st.markdown('<div class="section-header">Cost Breakdown</div>', unsafe_allow_html=True)
        bd = compute_cost_breakdown(df)
        if bd:
            fig = px.pie(names=list(bd.keys()), values=list(bd.values()),
                         color_discrete_sequence=COLORS, hole=0.45)
            fig.update_traces(textinfo="percent+label", textfont_size=10)
            fig.update_layout(margin=dict(t=10,b=10,l=10,r=10), height=360,
                              showlegend=False, font=dict(family="Inter"),
                              paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True, key="dash_pie")

    with c2:
        st.markdown('<div class="section-header">Top 15 Orders by Cost</div>', unsafe_allow_html=True)
        top = compute_top_orders(df)
        if not top.empty:
            fig = px.bar(top.sort_values('Total_Cost'), x='Total_Cost', y='Order',
                         orientation='h', color_discrete_sequence=[COLORS[0]])
            fig.update_layout(margin=dict(t=10,b=30,l=10,r=10), height=360,
                              font=dict(family="Inter"),
                              paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig, use_container_width=True, key="dash_top")

    # RMC Distribution
    st.markdown('<div class="section-header">RMC/Kg Distribution</div>', unsafe_allow_html=True)
    dist = compute_rmc_distribution(df)
    if dist:
        fig = px.bar(x=list(dist.keys()), y=list(dist.values()),
                     color_discrete_sequence=[COLORS[2]],
                     labels={'x':'RMC/Kg Range','y':'Order Count'})
        fig.update_layout(margin=dict(t=10,b=30,l=10,r=10), height=280,
                          font=dict(family="Inter"),
                          paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
        st.plotly_chart(fig, use_container_width=True, key="dash_dist")

    # Insights
    st.markdown('<div class="section-header">💡 AI Insights</div>', unsafe_allow_html=True)
    for ins in generate_insights(df, kpis):
        st.markdown(f'<div class="insight-card">{ins}</div>', unsafe_allow_html=True)

    # Recommendations
    st.markdown('<div class="section-header">📋 Recommendations</div>', unsafe_allow_html=True)
    for rec in generate_recommendations(df, kpis):
        cls = rec['priority'].lower()
        st.markdown(f'<div class="rec-card {cls}">'
                    f'<b>{rec["icon"]} [{rec["priority"]}] {rec["title"]}</b><br>'
                    f'{rec["detail"]}</div>', unsafe_allow_html=True)


def _render_audit_tab():
    """Per-cell audit panel."""
    from dashboard.audit_panel import trace_rmc_cell, get_rmc_orders, get_columns
    import openpyxl

    rmc_bytes = st.session_state.get("rmc_filled_template")
    if not rmc_bytes:
        st.info("Run pipeline with Base RMC Template to enable audit.")
        return

    st.markdown('<div class="section-header">🔍 Per-Cell Audit Trail</div>', unsafe_allow_html=True)
    st.caption("Select an order and column to trace the value back to its source.")

    wb = openpyxl.load_workbook(io.BytesIO(rmc_bytes), data_only=True)
    orders = get_rmc_orders(wb)
    columns = get_columns()

    if not orders:
        st.warning("No orders found in RMC Summary.")
        wb.close()
        return

    c1,c2 = st.columns(2)
    with c1:
        order = st.selectbox("Order:", orders, key="audit_order")
    with c2:
        col = st.selectbox("Column:", columns, key="audit_col")

    if order and col:
        trace = trace_rmc_cell(wb, order, col)
        st.markdown(f"""<div class="kpi-grid">
            <div class="kpi-card"><div class="kpi-label">Value</div>
                <div class="kpi-value">{trace['value']:,.4f}</div></div>
            <div class="kpi-card"><div class="kpi-label">Source</div>
                <div class="kpi-value blue" style="font-size:16px">{trace['source']}</div></div>
            <div class="kpi-card"><div class="kpi-label">Formula</div>
                <div class="kpi-value green" style="font-size:12px">{trace['formula']}</div></div>
        </div>""", unsafe_allow_html=True)

        if trace.get('rows'):
            st.markdown(f"**Contributing Rows** ({len(trace['rows'])} rows):")
            rdf = pd.DataFrame(trace['rows'])
            st.dataframe(rdf, use_container_width=True, hide_index=True)
        elif trace.get('error'):
            st.warning(trace['error'])

    wb.close()


def _render_flags_tab():
    """Manual-check flag list."""
    import openpyxl
    from dashboard.flag_checker import run_all_checks

    rmc_bytes = st.session_state.get("rmc_filled_template")
    if not rmc_bytes:
        st.info("Run pipeline with Base RMC Template to enable flag checking.")
        return

    st.markdown('<div class="section-header">⚠️ Manual-Check Flag List</div>', unsafe_allow_html=True)

    wb = openpyxl.load_workbook(io.BytesIO(rmc_bytes), data_only=True)
    log = st.session_state.get("rmc_log", [])
    flags = run_all_checks(wb, log)
    wb.close()

    if not flags:
        st.success("🎉 No anomalies detected! All checks passed.")
        return

    # Summary
    high = sum(1 for f in flags if 'HIGH' in f['severity'])
    med = sum(1 for f in flags if 'MEDIUM' in f['severity'])
    st.markdown(f"""<div class="kpi-grid">
        <div class="kpi-card"><div class="kpi-label">Total Flags</div>
            <div class="kpi-value">{len(flags)}</div></div>
        <div class="kpi-card"><div class="kpi-label">High Severity</div>
            <div class="kpi-value red">{high}</div></div>
        <div class="kpi-card"><div class="kpi-label">Medium Severity</div>
            <div class="kpi-value" style="background:linear-gradient(135deg,#F59E0B,#FBBF24);-webkit-background-clip:text;-webkit-text-fill-color:transparent">{med}</div></div>
    </div>""", unsafe_allow_html=True)

    # Filter
    cats = sorted(set(f['category'] for f in flags))
    sel_cat = st.multiselect("Filter by category:", cats, default=cats, key="flag_cats")
    filtered = [f for f in flags if f['category'] in sel_cat]

    fdf = pd.DataFrame(filtered)
    cols_show = ['severity','category','sheet','order','value','action']
    cols_show = [c for c in cols_show if c in fdf.columns]
    st.dataframe(fdf[cols_show], use_container_width=True, height=400, hide_index=True)


def _render_cross_month_tab():
    """Cross-month validation."""
    from dashboard.cross_month import load_rmc_summary, compare_months

    st.markdown('<div class="section-header">📆 Cross-Month Validation</div>', unsafe_allow_html=True)

    curr_bytes = st.session_state.get("rmc_filled_template") or st.session_state.get("rmc_output_bytes")
    prev_file = st.session_state.get("rmc_prev_month")

    if not curr_bytes:
        st.info("Run the pipeline first.")
        return
    if not prev_file:
        st.warning("Upload **Previous Month RMC** in the sidebar to enable cross-month comparison.")
        return

    prev_bytes = prev_file.read() if hasattr(prev_file, 'read') else prev_file

    with st.spinner("Loading and comparing..."):
        df_prev = load_rmc_summary(prev_bytes, "Prev")
        df_curr = load_rmc_summary(curr_bytes, "Curr")
        comp = compare_months(df_prev, df_curr, "Prev", "Curr")

    # Summary KPIs
    sp = comp.get('summary_p',{})
    sc = comp.get('summary_c',{})
    score = comp.get('consistency_score', 0)

    st.markdown(f"""<div class="kpi-grid">
        <div class="kpi-card"><div class="kpi-label">Common Orders</div>
            <div class="kpi-value">{len(comp['common_orders'])}</div></div>
        <div class="kpi-card"><div class="kpi-label">New Orders</div>
            <div class="kpi-value blue">{len(comp['new_orders'])}</div></div>
        <div class="kpi-card"><div class="kpi-label">Dropped Orders</div>
            <div class="kpi-value red">{len(comp['dropped_orders'])}</div></div>
        <div class="kpi-card"><div class="kpi-label">Consistency</div>
            <div class="kpi-value {'green' if score>90 else 'red'}">{score:.1f}%</div></div>
    </div>""", unsafe_allow_html=True)

    # Rate Drift
    drift = comp.get('rate_drift', [])
    if drift:
        st.markdown("### Rate Drift")
        st.dataframe(pd.DataFrame(drift), use_container_width=True, hide_index=True)

    # Deltas
    deltas = comp.get('deltas', [])
    if deltas:
        st.markdown("### Order-Level Deltas")
        ddf = pd.DataFrame(deltas)
        show_cols = ['Order'] + [c for c in ddf.columns if 'delta' in c or 'pct' in c]
        show_cols = [c for c in show_cols if c in ddf.columns]
        st.dataframe(ddf[show_cols], use_container_width=True, height=400, hide_index=True)


if __name__ == "__main__":
    main()
