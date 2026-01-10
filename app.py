import io
from datetime import datetime
from pathlib import Path

import streamlit as st

from src.config import AppConfig, ensure_dirs
from src.orchestrator import run_kpi_update


st.set_page_config(page_title="SAP KPI Report Generator", layout="wide")
st.title("SAP KPI Report Generator")

config = AppConfig.from_env()
ensure_dirs(config)

with st.sidebar:
    st.header("Inputs")
    variant = st.text_input("SAP Variant", value=f"CLV-PG{datetime.now().year}")
    sap_system = st.text_input("SAP System", value=config.sap_system)
    output_dir = st.text_input("Output Directory", value=str(config.output_dir))

    st.markdown("---")
    st.caption("Provide either the main workbook and optionally _woStatus.xlsx if SAP extraction is not available.")
    main_wb = st.file_uploader("Main Workbook (.xlsx)", type=["xlsx"]) 
    wo_status = st.file_uploader("Optional _woStatus.xlsx", type=["xlsx"]) 

run = st.button("Run KPI Update", type="primary")

log_container = st.expander("Logs", expanded=False)

if run:
    if not main_wb:
        st.error("Upload the main workbook.")
    else:
        # Persist uploads
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)

        main_wb_path = output_path / main_wb.name
        with open(main_wb_path, "wb") as f:
            f.write(main_wb.read())

        wo_status_path = None
        if wo_status is not None:
            wo_status_path = output_path / "_woStatus.xlsx"
            with open(wo_status_path, "wb") as f:
                f.write(wo_status.read())

        with st.spinner("Running... this can take a few minutes"):
            result = run_kpi_update(
                variant=variant,
                main_wb_file=str(main_wb_path),
                output_dir=str(output_path),
                sap_system=sap_system,
                wo_status_file=str(wo_status_path) if wo_status_path else None,
            )

        with log_container:
            st.text(result.logs)

        if result.success and result.output_path:
            st.success("Completed successfully.")
            with open(result.output_path, "rb") as f:
                st.download_button(
                    label="Download Updated Workbook",
                    data=f,
                    file_name=Path(result.output_path).name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        else:
            st.error(result.message)

st.markdown("---")
st.caption("If SAP extraction is not available on this host, upload _woStatus.xlsx explicitly.")
