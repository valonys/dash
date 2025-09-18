from __future__ import annotations

import io
import logging
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Tuple

from .data_process import process_wo_status, update_main_workbook
from .sap_extract import extract_sap_data
from .sheet_generate import generate_all_sheets


@dataclass
class RunResult:
    success: bool
    message: str
    output_path: Optional[str]
    logs: str


def _configure_logging_to_buffer() -> io.StringIO:
    log_stream = io.StringIO()
    handler = logging.StreamHandler(log_stream)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)

    root = logging.getLogger()
    root.setLevel(logging.INFO)
    # Avoid duplicate handlers if called multiple times
    for h in list(root.handlers):
        root.removeHandler(h)
    root.addHandler(handler)
    return log_stream


def run_kpi_update(variant: str, main_wb_file: str | Path, output_dir: str | Path, sap_system: str, wo_status_file: Optional[str | Path] = None) -> RunResult:
    log_buffer = _configure_logging_to_buffer()
    try:
        logging.info("Starting KPI update pipeline")
        output_dir_path = Path(output_dir).resolve()
        output_dir_path.mkdir(parents=True, exist_ok=True)

        main_path = Path(main_wb_file).resolve()
        updated_main_path = main_path  # We update in-place

        # Step 1: Obtain _woStatus.xlsx
        wo_status_path = None
        if wo_status_file:
            wo_status_path = Path(wo_status_file).resolve()
            logging.info("Using provided _woStatus.xlsx: %s", wo_status_path)
        else:
            ok, saved = extract_sap_data(variant, output_dir_path, sap_system)
            if not ok or not saved:
                logging.warning("SAP extraction not executed. Please upload _woStatus.xlsx.")
                return RunResult(False, "SAP extraction not available. Provide _woStatus.xlsx.", None, log_buffer.getvalue())
            wo_status_path = Path(saved)

        # Step 2: Process _woStatus.xlsx
        if not process_wo_status(wo_status_path):
            return RunResult(False, "Processing _woStatus.xlsx failed.", None, log_buffer.getvalue())

        # Step 3: Update main workbook
        if not update_main_workbook(updated_main_path, wo_status_path):
            return RunResult(False, "Updating main workbook failed.", None, log_buffer.getvalue())

        # Step 4: Generate categorized sheets
        if not generate_all_sheets(updated_main_path):
            return RunResult(False, "Generating categorized sheets failed.", None, log_buffer.getvalue())

        logging.info("KPI update pipeline completed")
        return RunResult(True, "Success", str(updated_main_path), log_buffer.getvalue())
    except Exception as exc:
        logging.exception("Run failed: %s", exc)
        return RunResult(False, f"Run failed: {exc}", None, log_buffer.getvalue())
