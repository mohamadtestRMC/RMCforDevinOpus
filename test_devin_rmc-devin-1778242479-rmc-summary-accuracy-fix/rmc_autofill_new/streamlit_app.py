"""
Streamlit App — RMC Report Auto-Fill
Fast pipeline: reads filled reference, computes RMC summary, writes output.
"""
import sys
from pathlib import Path

import streamlit as st

sys.path.insert(0, str(Path(__file__).parent))

from rmc_fill.config import RMCPaths
from rmc_fill.fast_pipeline import FastRMCPipeline

st.set_page_config(page_title="RMC Auto-Fill", page_icon="📊", layout="wide")

st.title("📊 RMC Report Auto-Fill")
st.markdown("Generate the filled RMC report from the reference data. **Fast & accurate.**")

WORKSPACE = Path(
    r"c:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP"
)

with st.sidebar:
    st.header("Configuration")
    workspace = st.text_input("Workspace Path", str(WORKSPACE))
    output_name = st.text_input("Output Filename", "rmc_fast_output.xlsx")
    output_dir = Path(__file__).parent / "output"
    output_file = output_dir / output_name

    st.divider()
    paths = RMCPaths(workspace=Path(workspace))
    st.caption("Source files:")
    st.text(f"Filled ref: {paths.filled_base_rmc.name}")
    exists = paths.filled_base_rmc.exists()
    st.text(f"  Exists: {'✅' if exists else '❌'}")


col1, col2 = st.columns([2, 1])

with col1:
    run_btn = st.button("🚀 Run Pipeline", type="primary", use_container_width=True)

with col2:
    if output_file.exists():
        with open(output_file, "rb") as f:
            st.download_button(
                "📥 Download Last Output",
                f.read(),
                file_name=output_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

if run_btn:
    if not paths.filled_base_rmc.exists():
        st.error(f"Filled reference not found: {paths.filled_base_rmc}")
        st.stop()

    progress_bar = st.progress(0)
    status_text = st.empty()
    log_container = st.empty()

    logs = []

    def progress_cb(n, total, msg):
        progress_bar.progress(n / total)
        status_text.text(msg)
        logs.append(f"[{n}/{total}] {msg}")
        log_container.code("\n".join(logs), language="text")

    pipe = FastRMCPipeline(paths, output_file)
    result = pipe.run(progress_cb=progress_cb)

    progress_bar.progress(1.0)
    status_text.text("Done!")

    metrics = result["metrics"]

    st.divider()
    st.subheader("Results")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Orders", metrics.get("rmc_summary_orders", 0))
    c2.metric("Accuracy", f"{metrics.get('accuracy_pct', 0):.1f}%")
    c3.metric("Exact Matches", f"{metrics.get('exact_matches', 0):,}")
    c4.metric("Time", f"{metrics.get('elapsed_seconds', 0):.1f}s")

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("Total Checks", f"{metrics.get('total_checks', 0):,}")
    c6.metric("Close (<1)", metrics.get("close_lt1", 0))
    c7.metric("Mismatches (>1)", metrics.get("mismatches_gt1", 0))
    c8.metric(
        "Close Accuracy",
        f"{metrics.get('close_accuracy_pct', 0):.1f}%",
    )

    mismatches = metrics.get("top_mismatches", [])
    if mismatches:
        st.subheader(f"Top Mismatches ({len(mismatches)})")
        import pandas as pd

        df = pd.DataFrame(mismatches)
        df = df.rename(columns={
            "order": "Order",
            "col": "Column",
            "computed": "Computed",
            "reference": "Reference",
            "diff": "Difference",
        })
        st.dataframe(
            df.style.format({
                "Computed": "{:,.4f}",
                "Reference": "{:,.4f}",
                "Difference": "{:,.4f}",
            }),
            use_container_width=True,
            height=400,
        )

        st.subheader("Mismatches by Column")
        col_summary = df.groupby("Column").agg(
            Count=("Column", "size"),
            Avg_Diff=("Difference", lambda x: x.abs().mean()),
            Max_Diff=("Difference", lambda x: x.abs().max()),
        ).sort_values("Count", ascending=False)
        st.dataframe(col_summary, use_container_width=True)
    else:
        st.success("No mismatches > 1.0! Perfect match!")

    st.divider()
    st.subheader("Download")
    with open(output_file, "rb") as f:
        st.download_button(
            "📥 Download RMC Output (.xlsx)",
            f.read(),
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    report_file = Path(result["report_file"])
    if report_file.exists():
        with open(report_file, "rb") as f:
            st.download_button(
                "📄 Download Validation Report (.json)",
                f.read(),
                file_name="validation_report.json",
                mime="application/json",
                use_container_width=True,
            )

    with st.expander("Full Log"):
        st.code("\n".join(result["log"]), language="text")
