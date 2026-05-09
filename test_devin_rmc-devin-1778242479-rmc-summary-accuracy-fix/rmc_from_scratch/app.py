"""
RMC Report Generator — Streamlit Application
Generates the Base RMC report from source files with full investigation & CEO dashboard.
"""
import sys
import io
import logging
from pathlib import Path
from datetime import datetime

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent.parent))

from rmc_engine.config import PipelineConfig, SourceFiles
from rmc_engine.pipeline import RMCPipeline
from rmc_engine.rmc_compute import RMC_COL_ORDER, TEXT_COLS
from rmc_engine.data_reader import safe_float

logging.basicConfig(level=logging.INFO)

st.set_page_config(
    page_title="RMC Report Generator",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

COLORS = {
    "primary": "#4F46E5",
    "success": "#059669",
    "warning": "#D97706",
    "danger": "#DC2626",
    "blue": "#2563EB",
    "slate": "#475569",
    "bg": "#F8FAFC",
}

CHART_COLORS = [
    "#4F46E5", "#7C3AED", "#2563EB", "#059669", "#D97706",
    "#DC2626", "#0891B2", "#DB2777", "#EA580C", "#65A30D",
]

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');
    .stApp { background: #F8FAFC; font-family: 'Inter', sans-serif; }
    header[data-testid="stHeader"] { background: rgba(248,250,252,0.95); backdrop-filter: blur(12px); }
    section[data-testid="stSidebar"] { background: #FFFFFF; border-right: 1px solid #E2E8F0; }
    .kpi-row { display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 16px; }
    .kpi-card {
        background: #fff; border: 1px solid #E2E8F0; border-radius: 12px;
        padding: 16px 20px; flex: 1; min-width: 160px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.04); transition: all 0.2s;
    }
    .kpi-card:hover { border-color: #94A3B8; transform: translateY(-1px); box-shadow: 0 4px 8px rgba(0,0,0,0.08); }
    .kpi-label { font-size: 11px; font-weight: 600; color: #64748B; text-transform: uppercase; letter-spacing: 0.8px; margin-bottom: 2px; }
    .kpi-value { font-size: 24px; font-weight: 700; color: #1E293B; line-height: 1.3; }
    .kpi-sub { font-size: 11px; color: #94A3B8; margin-top: 2px; }
    .kpi-value.green { color: #059669; }
    .kpi-value.blue { color: #2563EB; }
    .kpi-value.amber { color: #D97706; }
    .kpi-value.red { color: #DC2626; }
    .kpi-value.purple { color: #7C3AED; }
    .section-hdr { font-size: 17px; font-weight: 700; color: #1E293B; margin: 24px 0 10px; padding-bottom: 6px; border-bottom: 2px solid #E2E8F0; }
    .file-status { padding: 4px 10px; border-radius: 6px; font-size: 12px; font-weight: 600; display: inline-block; margin: 2px 0; }
    .file-ok { background: #ECFDF5; color: #059669; }
    .file-miss { background: #FEF2F2; color: #DC2626; }
    .stButton > button[kind="primary"] { background: linear-gradient(135deg, #4F46E5, #7C3AED) !important; border: none !important; color: white !important; }
    .dash-title { font-size: 22px; font-weight: 700; color: #1E293B; margin-bottom: 4px; }
    .dash-sub { font-size: 13px; color: #64748B; margin-bottom: 16px; }
</style>
""", unsafe_allow_html=True)

st.title("📊 RMC Report Generator")
st.caption("Generate the Base RMC report from source files. Upload, compute, and analyze every value.")

if "output_bytes" not in st.session_state:
    st.session_state.output_bytes = None
if "result" not in st.session_state:
    st.session_state.result = None
if "run_complete" not in st.session_state:
    st.session_state.run_complete = False


def _read_upload(uploaded_file):
    if uploaded_file is None:
        return None
    return io.BytesIO(uploaded_file.getvalue())


# ═══════════════════════════════════════════════════════════════════════════
#  SIDEBAR: FILE UPLOADS
# ═══════════════════════════════════════════════════════════════════════════
with st.sidebar:
    st.header("📁 Source Files")

    use_disk = st.checkbox(
        "Use files from disk (dev mode)", value=True,
        help="Load files from the workspace directory instead of uploading",
    )

    if use_disk:
        WORKSPACE = Path(
            r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP"
        )
        unfilled_dir = WORKSPACE / "Files_need_to_study" / "Unfilled"
        filled_dir = WORKSPACE / "Files_need_to_study" / "Filled_Output"
        template_dir = WORKSPACE / "Template3"

        file_map = {
            "Base RMC Template": unfilled_dir / "1 Base RMC _ 2026 February.xlsx",
            "Filled RMC (validation)": filled_dir / "1 Base RMC _ 2026 February.xlsx",
            "Jobtrack": template_dir / "Job Track Feb 26.xlsx",
            "Stores Recordings": template_dir / "Stores Recordings.xlsx",
            "Purchase Register": unfilled_dir / "2 Purchase Register - 2021 - 2026 _Feb 26.xlsx",
            "Opening WIP": unfilled_dir / "9 Opening WIP Stock.xlsx",
            "Closing WIP": unfilled_dir / "10 Closing WIP Stock.xlsx",
            "Granules (Current)": unfilled_dir / "4 Granules Recipe - February 2026.xlsx",
            "Granules (Previous)": unfilled_dir / "4 Granules Recipe - January 2026.xlsx",
            "MEGAPACK Rates": unfilled_dir / "6 MEGAPACK Rate.xlsx",
            "Ink Consumption": unfilled_dir / "5 Ink Consumption February 2026.xlsx",
            "Components": unfilled_dir / "12 Components Consumptions Dispensed Details.xlsx",
        }

        st.markdown("**File Status:**")
        for name, path in file_map.items():
            exists = path.exists()
            cls = "file-ok" if exists else "file-miss"
            icon = "✅" if exists else "❌"
            st.markdown(
                f'<span class="file-status {cls}">{icon} {name}</span>',
                unsafe_allow_html=True,
            )

        def _path_or_none(key):
            p = file_map.get(key)
            return p if p and p.exists() else None

        source_files = SourceFiles(
            base_rmc_template=_path_or_none("Base RMC Template"),
            filled_rmc_reference=_path_or_none("Filled RMC (validation)"),
            jobtrack=_path_or_none("Jobtrack"),
            stores_recordings=_path_or_none("Stores Recordings"),
            purchase_register=_path_or_none("Purchase Register"),
            opening_wip=_path_or_none("Opening WIP"),
            closing_wip=_path_or_none("Closing WIP"),
            granules_current=_path_or_none("Granules (Current)"),
            granules_prev=_path_or_none("Granules (Previous)"),
            megapack_rates=_path_or_none("MEGAPACK Rates"),
            ink_consumption=_path_or_none("Ink Consumption"),
        )

    else:
        st.markdown("Upload each source file:")
        f_template = st.file_uploader("1. Base RMC Template", type=["xlsx"], key="template")
        f_jobtrack = st.file_uploader("2. Job Track", type=["xlsx"], key="jt")
        f_stores = st.file_uploader("3. Stores Recordings", type=["xlsx"], key="stores")
        f_pr = st.file_uploader("4. Purchase Register", type=["xlsx"], key="pr")
        f_opn_wip = st.file_uploader("5. Opening WIP Stock", type=["xlsx"], key="opnwip")
        f_cls_wip = st.file_uploader("6. Closing WIP Stock", type=["xlsx"], key="clswip")
        f_granules = st.file_uploader("7. Granules Recipe Current", type=["xlsx"], key="gran")
        f_gran_prev = st.file_uploader("8. Granules Recipe Previous", type=["xlsx"], key="granprev")
        f_mega = st.file_uploader("9. MEGAPACK Rates", type=["xlsx"], key="mega")
        f_ink = st.file_uploader("10. Ink Consumption", type=["xlsx"], key="ink")

        st.subheader("Optional (Validation)")
        f_filled_ref = st.file_uploader("Filled RMC Reference", type=["xlsx"], key="filled_ref")
        f_prev_rmc = st.file_uploader("Previous Month RMC Output", type=["xlsx"], key="prev_rmc")

        source_files = SourceFiles(
            base_rmc_template=_read_upload(f_template),
            filled_rmc_reference=_read_upload(f_filled_ref),
            jobtrack=_read_upload(f_jobtrack),
            stores_recordings=_read_upload(f_stores),
            purchase_register=_read_upload(f_pr),
            opening_wip=_read_upload(f_opn_wip),
            closing_wip=_read_upload(f_cls_wip),
            granules_current=_read_upload(f_granules),
            granules_prev=_read_upload(f_gran_prev),
            megapack_rates=_read_upload(f_mega),
            ink_consumption=_read_upload(f_ink),
        )

    st.divider()
    st.subheader("Pipeline Mode")
    mode = st.radio(
        "Computation mode:",
        ["scratch", "reference"],
        index=0,
        help="**scratch**: computes from source files. **reference**: reads filled RMC for validation.",
        captions=[
            "Compute from source files (automation)",
            "Read filled reference (validation only)",
        ],
    )


# ═══════════════════════════════════════════════════════════════════════════
#  MAIN AREA: RUN BUTTON
# ═══════════════════════════════════════════════════════════════════════════
col_run, col_dl = st.columns([2, 1])

with col_run:
    run_btn = st.button("🚀 Generate RMC Report", type="primary", use_container_width=True)

with col_dl:
    if st.session_state.output_bytes:
        st.download_button(
            "📥 Download RMC Output",
            data=st.session_state.output_bytes,
            file_name=f"rmc_output_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

# ═══════════════════════════════════════════════════════════════════════════
#  PIPELINE EXECUTION
# ═══════════════════════════════════════════════════════════════════════════
if run_btn:
    has_ref = source_files.filled_rmc_reference is not None

    if mode == "reference" and not has_ref:
        st.error("Reference mode selected but no filled RMC reference provided!")
        st.stop()

    if mode == "scratch":
        missing = []
        if source_files.jobtrack is None:
            missing.append("Job Track")
        if source_files.purchase_register is None:
            missing.append("Purchase Register")
        if source_files.base_rmc_template is None:
            missing.append("Base RMC Template")
        if missing:
            st.error(f"From-scratch mode requires: {', '.join(missing)}")
            st.stop()

        if has_ref:
            st.info("Filled reference detected — will validate computed results and use for carry-forward data.")

    config = PipelineConfig(
        source_files=source_files,
        output_dir=Path(__file__).parent / "output",
    )

    progress_bar = st.progress(0)
    status_text = st.empty()
    log_container = st.empty()
    logs = []

    def progress_cb(n, total, msg):
        progress_bar.progress(min(n / total, 1.0))
        status_text.text(msg)
        logs.append(f"[{n}/{total}] {msg}")
        log_container.code("\n".join(logs[-10:]), language="text")

    try:
        pipeline = RMCPipeline(config)
        result = pipeline.run(mode=mode, progress_cb=progress_cb)

        progress_bar.progress(1.0)
        status_text.text("Pipeline complete!")

        st.session_state.output_bytes = result.get("output_bytes")
        st.session_state.result = result
        st.session_state.run_complete = True

    except Exception as e:
        st.error(f"Pipeline error: {e}")
        import traceback
        st.code(traceback.format_exc())
        st.stop()


# ═══════════════════════════════════════════════════════════════════════════
#  RESULTS DISPLAY
# ═══════════════════════════════════════════════════════════════════════════
if st.session_state.run_complete and st.session_state.result:
    result = st.session_state.result
    metrics = result.get("metrics", {})
    rmc_rows = result.get("rmc_rows", [])

    st.divider()

    mode_label = metrics.get("mode", "unknown").upper()
    accuracy = metrics.get("accuracy_pct", 0)
    orders_count = metrics.get("rmc_summary_orders", 0)
    elapsed = metrics.get("elapsed_seconds", 0)
    exact = metrics.get("exact_matches", 0)
    total_checks = metrics.get("total_checks", 0)
    mismatches = metrics.get("mismatches_gt1", 0)
    close = metrics.get("close_lt1", 0)

    acc_class = "green" if accuracy >= 99.5 else "amber" if accuracy >= 95 else "red"

    kpi_html = f"""
    <div class="kpi-row">
        <div class="kpi-card">
            <div class="kpi-label">Mode</div>
            <div class="kpi-value blue">{mode_label}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Orders</div>
            <div class="kpi-value">{orders_count:,}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Accuracy</div>
            <div class="kpi-value {acc_class}">{accuracy:.2f}%</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Exact Matches</div>
            <div class="kpi-value green">{exact:,} / {total_checks:,}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Close (&lt;1)</div>
            <div class="kpi-value amber">{close}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Mismatches (&gt;1)</div>
            <div class="kpi-value {'red' if mismatches > 0 else 'green'}">{mismatches}</div>
        </div>
        <div class="kpi-card">
            <div class="kpi-label">Time</div>
            <div class="kpi-value blue">{elapsed:.1f}s</div>
        </div>
    </div>
    """
    st.markdown(kpi_html, unsafe_allow_html=True)

    if st.session_state.output_bytes:
        st.download_button(
            "📥 Download RMC Report (.xlsx)",
            data=st.session_state.output_bytes,
            file_name=f"rmc_output_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    # Build DataFrame from results
    if rmc_rows:
        df = pd.DataFrame(rmc_rows)
        cols_to_show = [c for c in RMC_COL_ORDER if c in df.columns]
        df_display = df[cols_to_show].copy()
        num_cols = [c for c in cols_to_show if c not in TEXT_COLS]
        for c in num_cols:
            df_display[c] = pd.to_numeric(df_display[c], errors="coerce").fillna(0)
    else:
        df_display = pd.DataFrame()

    # ═══════════════════════════════════════════════════════════════════
    #  TABS
    # ═══════════════════════════════════════════════════════════════════
    tab_dash, tab_summary, tab_mismatches, tab_investigate, tab_detail, tab_log = st.tabs([
        "📈 CEO Dashboard", "📊 Summary Table", "⚠️ Mismatches",
        "🔍 Investigate Order", "📋 Full Detail", "📝 Log",
    ])

    # ─── CEO DASHBOARD TAB ───
    with tab_dash:
        if df_display.empty:
            st.info("No data available for dashboard.")
        else:
            st.markdown('<div class="dash-title">Executive Cost Dashboard</div>', unsafe_allow_html=True)
            st.markdown('<div class="dash-sub">Monthly RMC Report — automated analysis</div>', unsafe_allow_html=True)

            total_cost = df_display["Total Cost"].sum()
            total_output = df_display["Prod / Output (Kg)"].sum()
            avg_rmc = total_cost / total_output if total_output > 0 else 0
            total_orders = len(df_display)

            opn_wip_val = df_display["Opening WIP Value (AED)"].sum()
            print_val = df_display["Printing Film Value (AED)"].sum()
            lam_val = df_display["Lam Fresh Mat Value (AED)"].sum()
            other_film_val = df_display["Other Film Value (AED)"].sum()
            ink_val = df_display["Ink & Sol Value (AED)"].sum()
            adh_val = df_display["Adh+ Hard +Sol Value (AED)"].sum()
            zip_val = df_display["Zipper + PE strip +Valve Value (AED)"].sum()
            cls_wip_val = df_display["Closing WIP Value (AED)"].sum()

            # Top KPI row
            st.markdown(f"""
            <div class="kpi-row">
                <div class="kpi-card">
                    <div class="kpi-label">Total Production Cost</div>
                    <div class="kpi-value purple">AED {total_cost:,.0f}</div>
                    <div class="kpi-sub">{total_orders:,} orders</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Total Output</div>
                    <div class="kpi-value blue">{total_output:,.0f} Kg</div>
                    <div class="kpi-sub">{total_output / 1000:,.1f} Tonnes</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Average RMC/Kg</div>
                    <div class="kpi-value green">AED {avg_rmc:,.2f}</div>
                    <div class="kpi-sub">cost per kilogram</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Opening WIP Value</div>
                    <div class="kpi-value">AED {opn_wip_val:,.0f}</div>
                </div>
                <div class="kpi-card">
                    <div class="kpi-label">Closing WIP Value</div>
                    <div class="kpi-value">AED {cls_wip_val:,.0f}</div>
                </div>
            </div>
            """, unsafe_allow_html=True)

            # Charts row 1: Cost breakdown + Top orders
            chart_col1, chart_col2 = st.columns(2)

            with chart_col1:
                st.markdown("#### Cost Breakdown by Category")
                cost_categories = {
                    "Printing Film": print_val,
                    "Lam Fresh Material": lam_val,
                    "Ink & Solvent": ink_val,
                    "Adhesive + Hardener + Solvent": adh_val,
                    "Other Film": other_film_val,
                    "Zipper + PE Strip + Valve": zip_val,
                    "Opening WIP": opn_wip_val,
                }
                cost_categories = {k: v for k, v in cost_categories.items() if v > 0}

                fig_pie = px.pie(
                    names=list(cost_categories.keys()),
                    values=list(cost_categories.values()),
                    color_discrete_sequence=CHART_COLORS,
                    hole=0.45,
                )
                fig_pie.update_traces(
                    textinfo="percent+label",
                    textfont_size=11,
                    hovertemplate="<b>%{label}</b><br>AED %{value:,.0f}<br>%{percent}<extra></extra>",
                )
                fig_pie.update_layout(
                    margin=dict(t=10, b=10, l=10, r=10),
                    height=380,
                    showlegend=False,
                    font=dict(family="Inter"),
                )
                st.plotly_chart(fig_pie, use_container_width=True)

            with chart_col2:
                st.markdown("#### Top 15 Orders by Total Cost")
                top_df = df_display.nlargest(15, "Total Cost")[["Order No", "Total Cost", "Prod / Output (Kg)", "Prod RMC / Kg"]].copy()
                top_df = top_df.sort_values("Total Cost", ascending=True)

                fig_bar = px.bar(
                    top_df,
                    x="Total Cost",
                    y="Order No",
                    orientation="h",
                    color="Prod RMC / Kg",
                    color_continuous_scale="Viridis",
                    labels={"Total Cost": "Total Cost (AED)", "Prod RMC / Kg": "RMC/Kg"},
                )
                fig_bar.update_layout(
                    margin=dict(t=10, b=30, l=10, r=10),
                    height=380,
                    yaxis=dict(tickfont=dict(size=10)),
                    font=dict(family="Inter"),
                    coloraxis_colorbar=dict(title="RMC/Kg", tickformat=",.1f"),
                )
                fig_bar.update_traces(
                    hovertemplate="<b>%{y}</b><br>Cost: AED %{x:,.0f}<extra></extra>",
                )
                st.plotly_chart(fig_bar, use_container_width=True)

            # Charts row 2: Cost components stacked + RMC distribution
            chart_col3, chart_col4 = st.columns(2)

            with chart_col3:
                st.markdown("#### Cost Components — Stacked Bar")
                cost_cols_aed = [
                    ("Printing Film Value (AED)", "Printing Film"),
                    ("Lam Fresh Mat Value (AED)", "Lam Fresh Mat"),
                    ("Ink & Sol Value (AED)", "Ink & Solvent"),
                    ("Adh+ Hard +Sol Value (AED)", "Adh+Hard+Solv"),
                    ("Other Film Value (AED)", "Other Film"),
                    ("Zipper + PE strip +Valve Value (AED)", "Zipper/Valve"),
                ]

                fig_stack = go.Figure()
                for col_name, label in cost_cols_aed:
                    fig_stack.add_trace(go.Bar(
                        name=label,
                        x=[label],
                        y=[df_display[col_name].sum()],
                        hovertemplate=f"<b>{label}</b><br>AED %{{y:,.0f}}<extra></extra>",
                    ))
                fig_stack.update_layout(
                    barmode="group",
                    margin=dict(t=10, b=30, l=10, r=10),
                    height=350,
                    showlegend=False,
                    yaxis=dict(title="AED", tickformat=","),
                    font=dict(family="Inter"),
                )
                st.plotly_chart(fig_stack, use_container_width=True)

            with chart_col4:
                st.markdown("#### RMC per Kg Distribution")
                rmc_data = df_display[df_display["Prod RMC / Kg"] > 0]["Prod RMC / Kg"]
                if not rmc_data.empty:
                    fig_hist = px.histogram(
                        rmc_data,
                        nbins=40,
                        labels={"value": "RMC / Kg (AED)", "count": "Orders"},
                        color_discrete_sequence=[COLORS["primary"]],
                    )
                    fig_hist.update_layout(
                        margin=dict(t=10, b=30, l=10, r=10),
                        height=350,
                        showlegend=False,
                        xaxis=dict(title="RMC per Kg (AED)"),
                        yaxis=dict(title="Number of Orders"),
                        font=dict(family="Inter"),
                    )
                    st.plotly_chart(fig_hist, use_container_width=True)

            # Wastage analysis
            st.markdown("#### Wastage Analysis")
            waste_cols = {
                "BFL Wastage Qty": "BFL",
                "Print Wastage Qty": "Print",
                "Lam Wastage Qty": "Lam",
                "Slit Wastage Qty": "Slit",
                "B&P Wastage Qty": "Bag & Pouch",
                "S&V Wastage Qty": "Spout & Valve",
                "HCI Wastage Qty": "HCI Rew",
                "PTR Wastage Qty": "PTR Rew",
            }
            existing_waste = {k: v for k, v in waste_cols.items() if k in df_display.columns}
            if existing_waste:
                waste_data = []
                for col, label in existing_waste.items():
                    total_w = df_display[col].sum()
                    if total_w > 0:
                        waste_data.append({"Process": label, "Wastage (Kg)": total_w})

                if waste_data:
                    wdf = pd.DataFrame(waste_data).sort_values("Wastage (Kg)", ascending=False)
                    wcol1, wcol2 = st.columns([2, 1])

                    with wcol1:
                        fig_waste = px.bar(
                            wdf, x="Process", y="Wastage (Kg)",
                            color="Process",
                            color_discrete_sequence=CHART_COLORS,
                        )
                        fig_waste.update_layout(
                            margin=dict(t=10, b=30, l=10, r=10),
                            height=300,
                            showlegend=False,
                            yaxis=dict(title="Wastage (Kg)", tickformat=","),
                            font=dict(family="Inter"),
                        )
                        fig_waste.update_traces(
                            hovertemplate="<b>%{x}</b><br>%{y:,.1f} Kg<extra></extra>",
                        )
                        st.plotly_chart(fig_waste, use_container_width=True)

                    with wcol2:
                        total_waste_kg = wdf["Wastage (Kg)"].sum()
                        st.markdown(f"""
                        <div class="kpi-card" style="margin-top:20px">
                            <div class="kpi-label">Total Wastage</div>
                            <div class="kpi-value red">{total_waste_kg:,.0f} Kg</div>
                            <div class="kpi-sub">{total_waste_kg / total_output * 100:.2f}% of output</div>
                        </div>
                        """, unsafe_allow_html=True)

                        waste_val_cols = {
                            "BFL Wastage Val": "BFL",
                            "Print Wastage Val": "Print",
                            "Lam Wastage Val": "Lam",
                            "Slit Wastage Val": "Slit",
                            "B&P Wastage Val": "Bag & Pouch",
                            "S&V Wastage Val": "Spout & Valve",
                            "HCI Wastage Val": "HCI Rew",
                            "PTR Wastage Val": "PTR Rew",
                        }
                        total_waste_val = sum(
                            df_display[c].sum() for c in waste_val_cols if c in df_display.columns
                        )
                        st.markdown(f"""
                        <div class="kpi-card" style="margin-top:12px">
                            <div class="kpi-label">Total Wastage Value</div>
                            <div class="kpi-value red">AED {total_waste_val:,.0f}</div>
                            <div class="kpi-sub">{total_waste_val / total_cost * 100:.2f}% of total cost</div>
                        </div>
                        """, unsafe_allow_html=True)

            # Material breakdown
            if "Material" in df_display.columns:
                st.markdown("#### Orders by Material Type")
                mat_df = df_display.groupby("Material").agg(
                    Orders=("Order No", "count"),
                    Total_Cost=("Total Cost", "sum"),
                    Total_Output=("Prod / Output (Kg)", "sum"),
                ).sort_values("Total_Cost", ascending=False).reset_index()
                mat_df = mat_df[mat_df["Material"].str.strip() != ""]

                if not mat_df.empty:
                    mat_col1, mat_col2 = st.columns(2)
                    with mat_col1:
                        fig_mat = px.treemap(
                            mat_df,
                            path=["Material"],
                            values="Total_Cost",
                            color="Total_Output",
                            color_continuous_scale="Blues",
                            labels={"Total_Cost": "Cost (AED)", "Total_Output": "Output (Kg)"},
                        )
                        fig_mat.update_layout(
                            margin=dict(t=10, b=10, l=10, r=10),
                            height=350,
                            font=dict(family="Inter"),
                        )
                        fig_mat.update_traces(
                            hovertemplate="<b>%{label}</b><br>Cost: AED %{value:,.0f}<br>Output: %{color:,.0f} Kg<extra></extra>",
                        )
                        st.plotly_chart(fig_mat, use_container_width=True)

                    with mat_col2:
                        st.dataframe(
                            mat_df.style.format({
                                "Total_Cost": "AED {:,.0f}",
                                "Total_Output": "{:,.0f} Kg",
                            }),
                            use_container_width=True,
                            hide_index=True,
                            height=350,
                        )

    # ─── SUMMARY TABLE TAB ───
    with tab_summary:
        st.markdown('<div class="section-hdr">RMC Summary — Computed Values</div>', unsafe_allow_html=True)
        if not df_display.empty:
            format_dict = {c: "{:,.2f}" for c in num_cols}
            st.dataframe(
                df_display.style.format(format_dict, na_rep=""),
                use_container_width=True,
                height=500,
            )
            csv = df_display.to_csv(index=False)
            st.download_button("📥 Download as CSV", csv, "rmc_summary.csv", "text/csv")

    # ─── MISMATCHES TAB ───
    with tab_mismatches:
        top_mismatches = metrics.get("top_mismatches", [])
        if top_mismatches:
            st.markdown(f'<div class="section-hdr">Top Mismatches ({len(top_mismatches)})</div>', unsafe_allow_html=True)
            mdf = pd.DataFrame(top_mismatches).rename(columns={
                "order": "Order", "col": "Column",
                "computed": "Computed", "reference": "Reference", "diff": "Difference",
            })
            st.dataframe(
                mdf.style.format({"Computed": "{:,.4f}", "Reference": "{:,.4f}", "Difference": "{:,.4f}"}),
                use_container_width=True, height=400,
            )

            st.markdown('<div class="section-hdr">Mismatches by Column</div>', unsafe_allow_html=True)
            col_summary = mdf.groupby("Column").agg(
                Count=("Column", "size"),
                Avg_Diff=("Difference", lambda x: x.abs().mean()),
                Max_Diff=("Difference", lambda x: x.abs().max()),
            ).sort_values("Count", ascending=False)
            st.dataframe(col_summary, use_container_width=True)
        else:
            st.success("All values match the reference perfectly! 100% accuracy achieved.")

    # ─── INVESTIGATE ORDER TAB ───
    with tab_investigate:
        st.markdown('<div class="section-hdr">Investigate a Specific Order</div>', unsafe_allow_html=True)
        st.caption("Select an order to see exactly how each value was computed.")

        order_list = [r.get("Order No", "") for r in rmc_rows if r.get("Order No")]
        selected_order = st.selectbox("Select Order No:", order_list)

        if selected_order:
            row_data = next((r for r in rmc_rows if r.get("Order No") == selected_order), None)
            if row_data:
                st.markdown(
                    f"**Order:** `{selected_order}` | "
                    f"**Remarks:** `{row_data.get('Remarks', '')}` | "
                    f"**Material:** `{row_data.get('Material', '')}` | "
                    f"**Structure:** `{row_data.get('Structure', '')}`"
                )

                ocol1, ocol2, ocol3 = st.columns(3)
                with ocol1:
                    st.metric("Total Cost", f"AED {safe_float(row_data.get('Total Cost')):,.2f}")
                with ocol2:
                    st.metric("Output (Kg)", f"{safe_float(row_data.get('Prod / Output (Kg)')):,.2f}")
                with ocol3:
                    st.metric("RMC / Kg", f"AED {safe_float(row_data.get('Prod RMC / Kg')):,.2f}")

                # Cost breakdown for this order
                order_costs = {
                    "Opening WIP": safe_float(row_data.get("Opening WIP Value (AED)")),
                    "Printing Film": safe_float(row_data.get("Printing Film Value (AED)")),
                    "Lam Fresh Mat": safe_float(row_data.get("Lam Fresh Mat Value (AED)")),
                    "Other Film": safe_float(row_data.get("Other Film Value (AED)")),
                    "Ink & Solvent": safe_float(row_data.get("Ink & Sol Value (AED)")),
                    "Adh+Hard+Solv": safe_float(row_data.get("Adh+ Hard +Sol Value (AED)")),
                    "Zipper/Valve": safe_float(row_data.get("Zipper + PE strip +Valve Value (AED)")),
                    "Closing WIP (-)": -safe_float(row_data.get("Closing WIP Value (AED)")),
                }
                order_costs = {k: v for k, v in order_costs.items() if abs(v) > 0}

                if order_costs:
                    fig_order = px.bar(
                        x=list(order_costs.keys()),
                        y=list(order_costs.values()),
                        labels={"x": "Cost Component", "y": "Value (AED)"},
                        color=list(order_costs.keys()),
                        color_discrete_sequence=CHART_COLORS,
                    )
                    fig_order.update_layout(
                        height=300,
                        margin=dict(t=10, b=30, l=10, r=10),
                        showlegend=False,
                        yaxis=dict(tickformat=","),
                        font=dict(family="Inter"),
                    )
                    st.plotly_chart(fig_order, use_container_width=True)

                # Deep trace
                idx_data = result.get("idx")
                offsets_data = result.get("offsets", {})
                transfer_orders = result.get("transfer_orders", set())
                other_film_orders = result.get("other_film_orders", set())

                if idx_data:
                    from rmc_engine.trace import trace_order, format_trace_for_display
                    trace = trace_order(
                        selected_order, idx_data, offsets_data,
                        transfer_orders, other_film_orders,
                    )

                    st.markdown("---")
                    st.markdown("### Computation Breakdown by Sheet")

                    for sheet_name, sheet_data in trace.get("sheets", {}).items():
                        if isinstance(sheet_data, dict) and sheet_data.get("status"):
                            continue
                        n_rows = sheet_data.get("matching_rows", 0)
                        if n_rows == 0:
                            continue

                        with st.expander(f"**{sheet_name}** — {n_rows} matching rows", expanded=False):
                            cols_data = sheet_data.get("columns", {})
                            if cols_data:
                                col_rows = []
                                for col_name, cd in cols_data.items():
                                    raw = cd.get("raw_sumif", 0)
                                    offset = cd.get("offset", 0)
                                    final = cd.get("final", 0)
                                    target = cd.get("target", "")
                                    col_rows.append({
                                        "Column": col_name,
                                        "SUMIF Result": f"{raw:,.4f}",
                                        "Offset": f"{offset:,.4f}" if abs(offset) > 0.001 else "-",
                                        "Final Value": f"{final:,.4f}",
                                        "Maps To": target,
                                    })
                                st.dataframe(pd.DataFrame(col_rows), use_container_width=True, hide_index=True)

                            row_details = sheet_data.get("row_details", [])
                            if row_details:
                                st.caption(f"Individual rows from {sheet_name}:")
                                st.dataframe(pd.DataFrame(row_details), use_container_width=True, hide_index=True)

                    with st.expander("Raw Trace Output", expanded=False):
                        st.code(format_trace_for_display(trace), language="text")
                else:
                    qty_cols = list(RMC_COL_ORDER[7:15])
                    val_cols = list(RMC_COL_ORDER[15:23])
                    st.markdown("**Quantities (Kg):**")
                    q_data = {c: f"{safe_float(row_data.get(c)):,.4f}" for c in qty_cols}
                    st.json(q_data)
                    st.markdown("**Values (AED):**")
                    v_data = {c: f"{safe_float(row_data.get(c)):,.4f}" for c in val_cols}
                    st.json(v_data)

    # ─── FULL DETAIL TAB ───
    with tab_detail:
        st.markdown('<div class="section-hdr">Full RMC Data (All Columns)</div>', unsafe_allow_html=True)
        if rmc_rows:
            full_df = pd.DataFrame(rmc_rows)
            st.dataframe(full_df, use_container_width=True, height=600)

    # ─── LOG TAB ───
    with tab_log:
        st.markdown('<div class="section-hdr">Pipeline Log</div>', unsafe_allow_html=True)
        log_text = "\n".join(result.get("log", []))
        st.code(log_text, language="text")

        report_path = result.get("report_path")
        if report_path and Path(report_path).exists():
            with open(report_path, "r") as f:
                st.download_button(
                    "📄 Download Validation Report (JSON)",
                    f.read(), "validation_report.json", "application/json",
                )
