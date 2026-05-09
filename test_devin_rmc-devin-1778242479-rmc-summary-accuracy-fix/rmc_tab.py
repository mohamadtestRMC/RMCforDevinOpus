"""
RMC Generator Tab — renders inside the main app.py as Tab 2.
Unified pipeline: Jobtrack Fill → RMC Generation in one click.
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

# Ensure imports work
_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(_ROOT))
sys.path.insert(0, str(_ROOT / "rmc_from_scratch"))

CHART_COLORS = [
    "#4F46E5", "#7C3AED", "#2563EB", "#059669", "#D97706",
    "#DC2626", "#0891B2", "#DB2777", "#EA580C", "#65A30D",
]


def _safe_float(v):
    try:
        f = float(v) if v is not None else 0.0
        import math
        return 0.0 if math.isnan(f) or math.isinf(f) else f
    except (ValueError, TypeError):
        return 0.0


def render_rmc_sidebar():
    """Render the RMC Generator sidebar file uploads. Returns dict of files."""
    st.markdown("### 📂 Source Files")
    st.caption("Upload all monthly source files for RMC generation.")

    st.markdown("**Required Files:**")
    jt = st.file_uploader("Job Track (Raw)", type=['xlsx'], key="rmc_jt",
                          help="Monthly Jobtrack without MRR data")
    stores = st.file_uploader("Stores Recordings", type=['xlsx'], key="rmc_stores",
                              help="Material receipt recordings")
    pr = st.file_uploader("Purchase Register", type=['xlsx'], key="rmc_pr",
                          help="Purchase register with material rates")
    template = st.file_uploader("Base RMC Template", type=['xlsx'], key="rmc_template",
                                help="Unfilled Base RMC workbook (contains WIP/FG)")

    st.markdown("**Optional — Supplier Rates:**")
    granules = st.file_uploader("Granules Recipe (Current)", type=['xlsx'], key="rmc_gran")
    prev_gran = st.file_uploader("Granules Recipe (Previous)", type=['xlsx'], key="rmc_pgran")
    mega = st.file_uploader("MEGA PACK Rates", type=['xlsx'], key="rmc_mega")

    st.markdown("**Optional — RMC Inputs:**")
    ink = st.file_uploader("Ink Consumption", type=['xlsx'], key="rmc_ink")
    components = st.file_uploader("Components Consumption", type=['xlsx'], key="rmc_comp")
    opn_wip = st.file_uploader("Opening WIP Stock", type=['xlsx'], key="rmc_opnwip")
    cls_wip = st.file_uploader("Closing WIP Stock", type=['xlsx'], key="rmc_clswip")

    st.markdown("**Optional — Validation:**")
    filled_ref = st.file_uploader("Filled RMC Reference", type=['xlsx'], key="rmc_ref",
                                  help="For accuracy validation against known-good output")

    return {
        "jt": jt, "stores": stores, "pr": pr, "template": template,
        "granules": granules, "prev_granules": prev_gran, "megapack": mega,
        "ink": ink, "components": components,
        "opn_wip": opn_wip, "cls_wip": cls_wip,
        "filled_ref": filled_ref,
    }


def render_rmc_tab(files: dict):
    """Render the main RMC Generator tab content."""

    # Check required files
    required = ["jt", "stores", "pr"]
    missing = [k for k in required if files.get(k) is None]

    # Hero
    st.markdown("""
    <div style="text-align:center; margin-bottom:24px;">
        <div style="font-size:28px; font-weight:800;
             background:linear-gradient(135deg,#4F46E5 0%,#DB2777 50%,#7C3AED 100%);
             -webkit-background-clip:text; -webkit-text-fill-color:transparent;">
            RMC Report Generator</div>
        <div style="font-size:14px; color:#64748B; margin-top:4px;">
            One-click: Fill Jobtrack → Build Process Sheets → Compute RMC Summary</div>
    </div>
    """, unsafe_allow_html=True)

    # Status cards
    _render_file_status(files)

    if missing:
        st.warning(f"Required files missing: **{', '.join(missing)}**. Upload them in the sidebar.")
        return

    # Generate button
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        run_btn = st.button("🚀 Generate RMC Report", type="primary",
                            use_container_width=True, key="rmc_run_btn")

    # Download button if result exists
    if st.session_state.get("rmc_output_bytes"):
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            now = datetime.now().strftime("%Y%m%d_%H%M")
            st.download_button("📥 Download RMC Report (.xlsx)",
                               data=st.session_state["rmc_output_bytes"],
                               file_name=f"RMC_Output_{now}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               use_container_width=True, key="rmc_dl_btn")
        # Filled Base RMC template (full 27-sheet workbook)
        if st.session_state.get("rmc_filled_template"):
            with col2:
                st.download_button("📥 Download Filled Base RMC Template (.xlsx)",
                                   data=st.session_state["rmc_filled_template"],
                                   file_name=f"Base_RMC_Filled_{now}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   help="Full Base RMC workbook (all sheets) populated with computed data \u2014 same layout as your manual file.",
                                   use_container_width=True, key="rmc_tmpl_dl_btn")
        # Also offer filled Jobtrack download
        if st.session_state.get("rmc_filled_jt"):
            with col2:
                st.download_button("📥 Download Filled Jobtrack (.xlsx)",
                                   data=st.session_state["rmc_filled_jt"],
                                   file_name=f"Jobtrack_Filled_{now}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True, key="rmc_jt_dl")

    # Run pipeline
    if run_btn:
        _run_pipeline(files)

    # Show results
    if st.session_state.get("rmc_complete"):
        _render_results()


def _render_file_status(files):
    """Show file upload status as badges."""
    labels = {
        "jt": "Job Track", "stores": "Stores", "pr": "Purchase Register",
        "template": "RMC Template", "granules": "Granules", "prev_granules": "Prev Granules",
        "megapack": "MEGA PACK", "ink": "Ink Consumption", "components": "Components",
        "opn_wip": "Opening WIP", "cls_wip": "Closing WIP", "filled_ref": "Filled Reference",
    }
    required = {"jt", "stores", "pr"}
    html = '<div style="display:flex;flex-wrap:wrap;gap:6px;margin-bottom:16px;">'
    for k, label in labels.items():
        ok = files.get(k) is not None
        req = k in required
        if ok:
            bg, color, icon = "#ECFDF5", "#059669", "✅"
        elif req:
            bg, color, icon = "#FEF2F2", "#DC2626", "❌"
        else:
            bg, color, icon = "#F8FAFC", "#94A3B8", "○"
        html += f'<span style="background:{bg};color:{color};padding:4px 10px;border-radius:6px;font-size:11px;font-weight:600;">{icon} {label}</span>'
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)


def _run_pipeline(files):
    """Execute the unified pipeline with progress bar."""
    progress_bar = st.progress(0)
    status_text = st.empty()
    detail_text = st.empty()
    log_area = st.empty()
    logs = []

    def progress_cb(pct, msg):
        progress_bar.progress(min(pct / 100.0, 1.0))
        status_text.markdown(f"**{msg}**")
        logs.append(f"[{pct:3d}%] {msg}")
        log_area.code("\n".join(logs[-8:]), language="text")

    try:
        from engine.rmc_pipeline import UnifiedRMCPipeline

        # Prepare file bytes
        def _bytes(f):
            if f is None:
                return None
            return io.BytesIO(f.read())

        pipeline = UnifiedRMCPipeline()
        result = pipeline.run(
            jt_file=_bytes(files["jt"]),
            stores_file=_bytes(files["stores"]),
            pr_file=_bytes(files["pr"]),
            granules_file=_bytes(files.get("granules")),
            megapack_file=_bytes(files.get("megapack")),
            prev_granules_file=_bytes(files.get("prev_granules")),
            base_rmc_template=_bytes(files.get("template")),
            opening_wip=_bytes(files.get("opn_wip")),
            closing_wip=_bytes(files.get("cls_wip")),
            ink_consumption=_bytes(files.get("ink")),
            components=_bytes(files.get("components")),
            filled_rmc_reference=_bytes(files.get("filled_ref")),
            progress_cb=progress_cb,
        )

        # Store results in session state
        st.session_state["rmc_output_bytes"] = result.get("output_bytes")
        st.session_state["rmc_filled_template"] = result.get("filled_template_bytes")
        st.session_state["rmc_filled_jt"] = result.get("filled_jt_bytes")
        st.session_state["rmc_rows"] = result.get("rmc_rows", [])
        st.session_state["rmc_ref_rows"] = result.get("rmc_ref_rows", [])
        st.session_state["rmc_metrics"] = result.get("metrics", {})
        st.session_state["rmc_log"] = result.get("log", [])
        st.session_state["rmc_jt_stats"] = result.get("jt_stats", {})
        st.session_state["rmc_investigation"] = result.get("investigation", [])
        st.session_state["rmc_idx"] = result.get("idx", {})
        st.session_state["rmc_offsets"] = result.get("offsets", {})
        st.session_state["rmc_transfers"] = result.get("transfer_orders", set())
        st.session_state["rmc_other_film"] = result.get("other_film_orders", set())
        st.session_state["rmc_complete"] = True

        progress_bar.progress(1.0)
        status_text.markdown("**✅ Pipeline complete!**")

        m = result.get("metrics", {})
        acc = m.get("accuracy_pct", -1)
        if acc >= 0:
            st.success(f"Generated {len(result.get('rmc_rows', []))} orders | "
                       f"Accuracy: {acc:.1f}% | Time: {m.get('elapsed_seconds', 0):.1f}s")
        else:
            st.success(f"Generated {len(result.get('rmc_rows', []))} orders | "
                       f"Time: {m.get('elapsed_seconds', 0):.1f}s (no reference for validation)")

    except Exception as e:
        st.error(f"Pipeline error: {e}")
        import traceback
        st.code(traceback.format_exc())
        st.session_state["rmc_complete"] = False


def _render_results():
    """Render the RMC results dashboard."""
    rmc_rows = st.session_state.get("rmc_rows", [])
    metrics = st.session_state.get("rmc_metrics", {})

    if not rmc_rows:
        st.info("No RMC data to display.")
        return

    from rmc_engine.rmc_compute import RMC_COL_ORDER, TEXT_COLS

    # Build DataFrame
    df = pd.DataFrame(rmc_rows)
    cols_to_show = [c for c in RMC_COL_ORDER if c in df.columns]
    df_display = df[cols_to_show].copy()
    num_cols = [c for c in cols_to_show if c not in TEXT_COLS]
    for c in num_cols:
        df_display[c] = pd.to_numeric(df_display[c], errors="coerce").fillna(0)

    st.markdown("---")

    # KPI row
    acc = metrics.get("accuracy_pct", -1)
    orders_count = metrics.get("rmc_summary_orders", len(rmc_rows))
    elapsed = metrics.get("elapsed_seconds", 0)
    exact = metrics.get("exact_matches", 0)
    total_checks = metrics.get("total_checks", 0)
    mismatches = metrics.get("mismatches_gt1", 0)

    acc_class = "green" if acc >= 99.5 else ("amber" if acc >= 95 else "red")
    if acc < 0:
        acc_class = "blue"
        acc_label = "N/A"
    else:
        acc_label = f"{acc:.1f}%"

    kpi_html = f"""
    <div class="kpi-row">
        <div class="kpi-card"><div class="kpi-label">Orders</div>
            <div class="kpi-value blue">{orders_count:,}</div></div>
        <div class="kpi-card"><div class="kpi-label">Accuracy</div>
            <div class="kpi-value {acc_class}">{acc_label}</div></div>
        <div class="kpi-card"><div class="kpi-label">Exact Matches</div>
            <div class="kpi-value green">{exact:,} / {total_checks:,}</div></div>
        <div class="kpi-card"><div class="kpi-label">Mismatches (&gt;1 AED)</div>
            <div class="kpi-value {'red' if mismatches > 0 else 'green'}">{mismatches}</div></div>
        <div class="kpi-card"><div class="kpi-label">Time</div>
            <div class="kpi-value blue">{elapsed:.1f}s</div></div>
    </div>
    """
    st.markdown(kpi_html, unsafe_allow_html=True)

    # ── Download buttons row ──
    now = datetime.now().strftime("%Y%m%d_%H%M")
    dl_cols = st.columns(3)
    with dl_cols[0]:
        if st.session_state.get("rmc_output_bytes"):
            st.download_button("📥 RMC Report (.xlsx)",
                data=st.session_state["rmc_output_bytes"],
                file_name=f"RMC_Output_{now}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, key="dl_rmc_out")
    with dl_cols[1]:
        if st.session_state.get("rmc_filled_template"):
            st.download_button("📥 Filled Base RMC (.xlsx)",
                data=st.session_state["rmc_filled_template"],
                file_name=f"Base_RMC_Filled_{now}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, key="dl_rmc_tmpl")
    with dl_cols[2]:
        if st.session_state.get("rmc_filled_jt"):
            st.download_button("📥 Filled Jobtrack (.xlsx)",
                data=st.session_state["rmc_filled_jt"],
                file_name=f"Jobtrack_Filled_{now}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True, key="dl_rmc_jt")

    st.markdown("---")

    # Tabs — include new Audit, Flags, Cross-Month
    tab_dash, tab_summary, tab_mismatch, tab_investigate, tab_audit, tab_flags, tab_log = st.tabs([
        "📈 Dashboard", "📊 Summary Table", "⚠️ Mismatches",
        "🔍 Investigate", "🔍 Audit Trail", "⚠️ Flag Check", "📝 Log",
    ])

    with tab_dash:
        _render_dashboard(df_display, num_cols)

    with tab_summary:
        st.markdown("### RMC Summary — All Orders")
        fmt = {c: "{:,.2f}" for c in num_cols}
        st.dataframe(df_display.style.format(fmt, na_rep=""),
                     use_container_width=True, height=500)
        csv = df_display.to_csv(index=False)
        st.download_button("📥 Download CSV", csv, "rmc_summary.csv", "text/csv")

    with tab_mismatch:
        _render_mismatches(metrics)

    with tab_investigate:
        _render_investigate(rmc_rows)

    with tab_audit:
        _render_audit_panel()

    with tab_flags:
        _render_flag_checker()

    with tab_log:
        log = st.session_state.get("rmc_log", [])
        st.code("\n".join(log[-50:]), language="text")


def _render_dashboard(df, num_cols):
    """CEO dashboard with charts."""
    if df.empty:
        st.info("No data for dashboard.")
        return

    total_cost = df["Total Cost"].sum() if "Total Cost" in df.columns else 0
    total_output = df["Prod / Output (Kg)"].sum() if "Prod / Output (Kg)" in df.columns else 0
    avg_rmc = total_cost / total_output if total_output > 0 else 0

    # Value columns
    val_cols = {
        "Opening WIP Value (AED)": "Opening WIP",
        "Printing Film Value (AED)": "Printing Film",
        "Lam Fresh Mat Value (AED)": "Lam Fresh Mat",
        "Other Film Value (AED)": "Other Film",
        "Ink & Sol Value (AED)": "Ink & Solvent",
        "Adh+ Hard +Sol Value (AED)": "Adh+Hard+Solv",
        "Zipper + PE strip +Valve Value (AED)": "Zipper/Valve",
    }

    st.markdown(f"""
    <div class="kpi-row">
        <div class="kpi-card"><div class="kpi-label">Total Production Cost</div>
            <div class="kpi-value">AED {total_cost:,.0f}</div>
            <div class="kpi-sub">{len(df):,} orders</div></div>
        <div class="kpi-card"><div class="kpi-label">Total Output</div>
            <div class="kpi-value blue">{total_output:,.0f} Kg</div></div>
        <div class="kpi-card"><div class="kpi-label">Avg RMC/Kg</div>
            <div class="kpi-value green">AED {avg_rmc:,.2f}</div></div>
    </div>
    """, unsafe_allow_html=True)

    # Charts
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### Cost Breakdown")
        cats = {label: df[col].sum() for col, label in val_cols.items()
                if col in df.columns and df[col].sum() > 0}
        if cats:
            fig = px.pie(names=list(cats.keys()), values=list(cats.values()),
                         color_discrete_sequence=CHART_COLORS, hole=0.45)
            fig.update_traces(textinfo="percent+label", textfont_size=10)
            fig.update_layout(margin=dict(t=10, b=10, l=10, r=10), height=360,
                              showlegend=False, font=dict(family="Inter"))
            st.plotly_chart(fig, use_container_width=True, key="rmc_pie")

    with col2:
        st.markdown("#### Top 15 Orders by Cost")
        if "Total Cost" in df.columns and "Order No" in df.columns:
            top = df.nlargest(15, "Total Cost")[["Order No", "Total Cost"]].sort_values("Total Cost")
            fig = px.bar(top, x="Total Cost", y="Order No", orientation="h",
                         color_discrete_sequence=[CHART_COLORS[0]])
            fig.update_layout(margin=dict(t=10, b=30, l=10, r=10), height=360,
                              font=dict(family="Inter"))
            st.plotly_chart(fig, use_container_width=True, key="rmc_top")

    # Wastage
    waste_cols = {
        "BFL Wastage Qty": "BFL", "Print Wastage Qty": "Print",
        "Lam Wastage Qty": "Lam", "Slit Wastage Qty": "Slit",
        "B&P Wastage Qty": "Bag&Pouch", "HCI Wastage Qty": "HCI Rew",
        "PTR Wastage Qty": "PTR Rew",
    }
    waste_data = [{
        "Process": label, "Wastage (Kg)": df[col].sum()
    } for col, label in waste_cols.items() if col in df.columns and df[col].sum() > 0]

    if waste_data:
        st.markdown("#### Wastage by Process")
        wdf = pd.DataFrame(waste_data).sort_values("Wastage (Kg)", ascending=False)
        fig = px.bar(wdf, x="Process", y="Wastage (Kg)", color="Process",
                     color_discrete_sequence=CHART_COLORS)
        fig.update_layout(margin=dict(t=10, b=30, l=10, r=10), height=300,
                          showlegend=False, font=dict(family="Inter"))
        st.plotly_chart(fig, use_container_width=True, key="rmc_waste")


def _render_mismatches(metrics):
    """Show mismatches tab."""
    top_mm = metrics.get("top_mismatches", [])
    if top_mm:
        st.markdown(f"### {len(top_mm)} Mismatches (> 1 AED difference)")
        mdf = pd.DataFrame(top_mm).rename(columns={
            "order": "Order", "col": "Column",
            "computed": "Computed", "reference": "Reference", "diff": "Difference",
        })
        st.dataframe(mdf.style.format({
            "Computed": "{:,.4f}", "Reference": "{:,.4f}", "Difference": "{:,.4f}"
        }), use_container_width=True, height=400)

        # Summary by column
        st.markdown("### Mismatches by Column")
        col_summary = mdf.groupby("Column").agg(
            Count=("Column", "size"),
            Avg_Diff=("Difference", lambda x: x.abs().mean()),
            Max_Diff=("Difference", lambda x: x.abs().max()),
        ).sort_values("Count", ascending=False)
        st.dataframe(col_summary, use_container_width=True)
    else:
        st.success("🎉 All values match the reference! 100% accuracy.")


def _render_investigate(rmc_rows):
    """Per-order investigation."""
    st.markdown("### Investigate a Specific Order")
    st.caption("Select an order to see all contributing values and trace computation.")

    order_list = [r.get("Order No", "") for r in rmc_rows if r.get("Order No")]
    if not order_list:
        st.info("No orders to investigate.")
        return

    selected = st.selectbox("Select Order:", order_list, key="rmc_investigate_order")
    if not selected:
        return

    row_data = next((r for r in rmc_rows if r.get("Order No") == selected), None)
    if not row_data:
        return

    st.markdown(f"**Order:** `{selected}` | "
                f"**Remarks:** `{row_data.get('Remarks', '')}` | "
                f"**Material:** `{row_data.get('Material', '')}` | "
                f"**Customer:** `{row_data.get('Customer Name', '')}`")

    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Total Cost", f"AED {_safe_float(row_data.get('Total Cost')):,.2f}")
    with c2:
        st.metric("Output (Kg)", f"{_safe_float(row_data.get('Prod / Output (Kg)')):,.2f}")
    with c3:
        st.metric("RMC/Kg", f"AED {_safe_float(row_data.get('Prod RMC / Kg')):,.2f}")

    # Cost breakdown chart
    costs = {
        "Opening WIP": _safe_float(row_data.get("Opening WIP Value (AED)")),
        "Print Film": _safe_float(row_data.get("Printing Film Value (AED)")),
        "Lam Fresh": _safe_float(row_data.get("Lam Fresh Mat Value (AED)")),
        "Other Film": _safe_float(row_data.get("Other Film Value (AED)")),
        "Ink&Sol": _safe_float(row_data.get("Ink & Sol Value (AED)")),
        "Adh+Hard": _safe_float(row_data.get("Adh+ Hard +Sol Value (AED)")),
        "Zipper/Valve": _safe_float(row_data.get("Zipper + PE strip +Valve Value (AED)")),
        "CLS WIP (-)": -_safe_float(row_data.get("Closing WIP Value (AED)")),
    }
    costs = {k: v for k, v in costs.items() if abs(v) > 0}
    if costs:
        fig = px.bar(x=list(costs.keys()), y=list(costs.values()),
                     color=list(costs.keys()), color_discrete_sequence=CHART_COLORS,
                     labels={"x": "Component", "y": "AED"})
        fig.update_layout(height=280, margin=dict(t=10, b=30, l=10, r=10),
                          showlegend=False, font=dict(family="Inter"))
        st.plotly_chart(fig, use_container_width=True, key="rmc_order_bar")

    # Trace from investigation data
    investigation = st.session_state.get("rmc_investigation", [])
    order_inv = next((d for d in investigation if d.get("order") == selected), None)
    if order_inv:
        with st.expander("📋 Detailed Value Comparison (Computed vs Reference)", expanded=False):
            col_data = []
            for col_name, cd in order_inv.get("columns", {}).items():
                col_data.append({
                    "Column": col_name,
                    "Computed": cd.get("computed", 0),
                    "Reference": cd.get("reference", "N/A"),
                    "Diff": cd.get("diff", "N/A"),
                    "Match": "✅" if cd.get("match") else "❌",
                })
            if col_data:
                st.dataframe(pd.DataFrame(col_data), use_container_width=True, hide_index=True)

    # Deep trace using trace module
    idx = st.session_state.get("rmc_idx")
    offsets = st.session_state.get("rmc_offsets", {})
    transfers = st.session_state.get("rmc_transfers", set())
    other_film = st.session_state.get("rmc_other_film", set())

    if idx:
        with st.expander("🔍 SUMIF Trace (per process sheet)", expanded=False):
            try:
                from rmc_engine.trace import trace_order, format_trace_for_display
                trace = trace_order(selected, idx, offsets, transfers, other_film)
                st.code(format_trace_for_display(trace), language="text")
            except Exception as e:
                st.warning(f"Trace error: {e}")


def _render_audit_panel():
    """Per-cell audit trail tab."""
    st.markdown("### 🔍 Per-Cell Audit Trail")
    st.caption("Select an order and column to trace the RMC Summary value back to its source sheet.")

    rmc_rows = st.session_state.get("rmc_rows", [])
    if not rmc_rows:
        st.info("No RMC data. Run the pipeline first.")
        return

    try:
        from dashboard.audit_panel import trace_rmc_cell, get_columns, RMC_COLUMN_MAP
    except ImportError as e:
        st.warning(f"Audit module not available: {e}")
        return

    # Use the filled template workbook if available
    tmpl_bytes = st.session_state.get("rmc_filled_template")

    order_list = [r.get("Order No", "") for r in rmc_rows if r.get("Order No")]
    columns = get_columns()

    c1, c2 = st.columns(2)
    with c1:
        order = st.selectbox("Order:", order_list, key="audit_order_sel")
    with c2:
        col = st.selectbox("Column:", columns, key="audit_col_sel")

    if order and col and tmpl_bytes:
        import openpyxl
        try:
            wb = openpyxl.load_workbook(io.BytesIO(tmpl_bytes), data_only=True)
            trace = trace_rmc_cell(wb, order, col)
            wb.close()

            st.markdown(f"**Value:** `{trace['value']:,.4f}` | "
                        f"**Source:** `{trace['source']}` | "
                        f"**Formula:** `{trace['formula']}`")

            if trace.get('rows'):
                st.markdown(f"**Contributing Rows** ({len(trace['rows'])} rows):")
                st.dataframe(pd.DataFrame(trace['rows']),
                             use_container_width=True, hide_index=True)
        except Exception as e:
            st.warning(f"Audit trace error: {e}")
    elif not tmpl_bytes:
        # Fallback: show data from rmc_rows directly
        row_data = next((r for r in rmc_rows if r.get("Order No") == order), None)
        if row_data and col in RMC_COLUMN_MAP:
            col_num = RMC_COLUMN_MAP[col][0]
            formula = RMC_COLUMN_MAP[col][4]
            source = RMC_COLUMN_MAP[col][1] or "Computed"
            st.markdown(f"**Formula:** `{formula}` | **Source:** `{source}`")
            st.info("Upload a Base RMC Template to enable full row-level audit trace.")


def _render_flag_checker():
    """Manual-check flag list tab."""
    st.markdown("### ⚠️ Manual-Check Flag List")
    st.caption("Automated anomaly detection across all process sheets.")

    tmpl_bytes = st.session_state.get("rmc_filled_template")
    if not tmpl_bytes:
        st.info("Upload a Base RMC Template and run pipeline to enable flag checking.")
        return

    try:
        from dashboard.flag_checker import run_all_checks
    except ImportError as e:
        st.warning(f"Flag checker module not available: {e}")
        return

    import openpyxl
    try:
        wb = openpyxl.load_workbook(io.BytesIO(tmpl_bytes), data_only=True)
        log = st.session_state.get("rmc_log", [])
        flags = run_all_checks(wb, log)
        wb.close()
    except Exception as e:
        st.error(f"Flag check error: {e}")
        return

    if not flags:
        st.success("🎉 No anomalies detected! All checks passed.")
        return

    high = sum(1 for f in flags if 'HIGH' in f['severity'])
    med = sum(1 for f in flags if 'MEDIUM' in f['severity'])

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Flags", len(flags))
    c2.metric("🔴 High", high)
    c3.metric("🟡 Medium", med)

    cats = sorted(set(f['category'] for f in flags))
    sel_cat = st.multiselect("Filter by category:", cats, default=cats, key="flag_cat_filter")
    filtered = [f for f in flags if f['category'] in sel_cat]

    if filtered:
        fdf = pd.DataFrame(filtered)
        show = [c for c in ['severity','category','sheet','order','value','action'] if c in fdf.columns]
        st.dataframe(fdf[show], use_container_width=True, height=400, hide_index=True)

