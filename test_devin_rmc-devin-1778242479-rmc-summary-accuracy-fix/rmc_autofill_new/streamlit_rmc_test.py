from __future__ import annotations

import json
from pathlib import Path

import streamlit as st

from rmc_fill.config import RMCPaths
from rmc_fill.pipeline import RMCAutofillPipeline


st.set_page_config(page_title="RMC Autofill Test", layout="wide")
st.title("RMC Autofill Test Runner")
st.caption("Isolated accuracy-first runner for Base RMC generation and validation.")

default_workspace = Path(r"C:\Users\mohamad.al\OneDrive - AL KHAYYAT INVESTMENTS\Desktop\IPP")
workspace_str = st.text_input("Workspace Path", value=str(default_workspace))
copy_jobtrack = st.checkbox("Copy Jobtrack from Template3 (slower)", value=False)

default_out = default_workspace / "rmc_autofill_new" / "output" / "1 Base RMC _ 2026 February_generated.xlsx"
output_str = st.text_input("Output File Path", value=str(default_out))

run_btn = st.button("Generate RMC Report", type="primary")

if run_btn:
    workspace = Path(workspace_str)
    output = Path(output_str)
    output.parent.mkdir(parents=True, exist_ok=True)

    with st.spinner("Running pipeline... this may take a few minutes"):
        pipeline = RMCAutofillPipeline(
            paths=RMCPaths(workspace=workspace),
            output_file=output,
            copy_jobtrack=copy_jobtrack,
        )
        result = pipeline.run()

    st.success("Pipeline completed.")
    st.write("Output workbook:", str(result.output_file))
    st.write("Validation report:", str(result.validation_report_file))

    st.subheader("Metrics")
    st.json(result.metrics)

    if result.validation_report_file.exists():
        report = json.loads(result.validation_report_file.read_text(encoding="utf-8"))
        st.subheader("Validation Summary")
        st.write("Compared range:", report.get("compared_range"))
        st.write("Mismatch count:", report.get("mismatch_count"))

        mismatches = report.get("mismatches", [])
        if mismatches:
            st.subheader("Top mismatches")
            st.dataframe(mismatches[:200], use_container_width=True)

        with open(result.validation_report_file, "rb") as f:
            st.download_button(
                "Download validation_report.json",
                data=f.read(),
                file_name="validation_report.json",
                mime="application/json",
            )

    if result.output_file.exists():
        with open(result.output_file, "rb") as f:
            st.download_button(
                "Download generated workbook",
                data=f.read(),
                file_name=result.output_file.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

