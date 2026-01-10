from __future__ import annotations

import logging
from datetime import datetime
from pathlib import Path
from typing import Iterable

import pandas as pd


def process_wo_status(wo_status_path: str | Path) -> bool:
    """Clean and normalize the _woStatus.xlsx export.

    This function attempts to:
    - Trim column names
    - Ensure key identifier columns are string-typed
    - Derive a lightweight 'Last System Status' column when possible
    The exact column indices may differ by system; operations are guarded.
    """
    try:
        wo_path = Path(wo_status_path).resolve()
        df = pd.read_excel(wo_path)
        df.columns = [str(c).strip() for c in df.columns]

        # Best-effort normalization
        for candidate in ("Order", "Order Number", "WO Number", "OrderNo"):
            if candidate in df.columns:
                df[candidate] = df[candidate].astype(str).str.strip()

        # Derive 'Last System Status' from any candidate status column
        status_col = next((c for c in df.columns if str(c).lower().startswith("system status")), None)
        if status_col and "Last System Status" not in df.columns:
            df["Last System Status"] = df[status_col].astype(str).str[:4]

        # If there is a User Status or similar, attempt to reflect QCAP/EXDO
        user_status_col = next((c for c in df.columns if str(c).lower().startswith("user status")), None)
        if user_status_col and status_col:
            mask = df["Last System Status"].isin(["CLSD", "CLOT", "TCLO"])  # closed statuses
            # Prefer not to overwrite; add a normalized status field
            df["Normalized Status"] = df[user_status_col].astype(str)
            df.loc[mask, "Normalized Status"] = "QCAP"

        df.to_excel(wo_path, index=False)
        logging.info("Processed and saved WO status: %s", wo_path)
        return True
    except Exception as exc:
        logging.exception("Error processing _woStatus.xlsx: %s", exc)
        return False


def _safe_month(dt) -> int | None:
    try:
        return int(pd.to_datetime(dt).month)
    except Exception:
        return None


def update_main_workbook(main_wb_path: str | Path, wo_status_path: str | Path) -> bool:
    """Update the 'Data Base' sheet by merging status and due dates from _woStatus.xlsx.

    Since sheet structures vary, this function uses heuristic merges by common keys.
    Expected combined outputs:
    - Fill status column if a matching order is found
    - Compute simple delay bucket field if due dates exist
    - Derive month helper columns when dates are present
    """
    try:
        main_path = Path(main_wb_path).resolve()
        wo_path = Path(wo_status_path).resolve()

        db_df = pd.read_excel(main_path, sheet_name="Data Base", skiprows=4)
        wo_df = pd.read_excel(wo_path)

        db_df.columns = [str(c).strip() for c in db_df.columns]
        wo_df.columns = [str(c).strip() for c in wo_df.columns]

        # Identify key columns heuristically
        order_key_candidates: Iterable[str] = (
            "Order", "Order Number", "WO Number", "OrderNo", "Notification"
        )
        db_order_col = next((c for c in db_df.columns if c in order_key_candidates), None)
        wo_order_col = next((c for c in wo_df.columns if c in order_key_candidates), None)

        if not db_order_col or not wo_order_col:
            logging.warning("Could not identify order key columns. Skipping merge.")
            return False

        # Normalize key types
        db_df[db_order_col] = db_df[db_order_col].astype(str).str.strip()
        wo_df[wo_order_col] = wo_df[wo_order_col].astype(str).str.strip()

        # Candidate fields from WO status
        status_src_col = next((c for c in ["Normalized Status", "Last System Status", "User Status"] if c in wo_df.columns), None)
        due_date_src_col = next((c for c in ["Basic finish", "Basic Finish", "Finish date", "Finish Date", "Sched finish date"] if c in wo_df.columns), None)

        merged = db_df.merge(
            wo_df[[wo_order_col] + [c for c in [status_src_col, due_date_src_col] if c]],
            left_on=db_order_col,
            right_on=wo_order_col,
            how="left",
            suffixes=("", "_wo"),
        )

        # Write back into existing or new columns
        if status_src_col:
            target_status_col = next((c for c in ("SECE STATUS", "Status", "User Status") if c in merged.columns), None)
            if target_status_col is None:
                target_status_col = "Status"
            merged[target_status_col] = merged[status_src_col]

        if due_date_src_col:
            target_due_col = next((c for c in ("Due Date", "Next Insp", "Basic finish") if c in merged.columns), None)
            if target_due_col is None:
                target_due_col = "Due Date"
            merged[target_due_col] = pd.to_datetime(merged[due_date_src_col], errors="coerce")

            # Month helpers
            merged["Due Month"] = merged[target_due_col].apply(_safe_month)

            # Delay buckets relative to today
            today = pd.Timestamp(datetime.now().date())
            def bucket(d):
                if pd.isna(d):
                    return ""
                delta = (today - pd.Timestamp(d)).days
                if delta > 365 * 3:
                    return "> 3 Yrs"
                if delta > 365 * 2:
                    return "2 Yrs < x < 3 Yrs"
                if delta > 365:
                    return "1 Yrs < x < 2 Yrs"
                if delta > 182:
                    return "6 Months < x < 1 Yrs"
                return "< 6 Months"

            merged["Delay"] = merged[target_due_col].apply(bucket)

        # Persist back to the workbook, replacing the Data Base data region
        with pd.ExcelWriter(main_path, engine="openpyxl", mode="a", if_sheet_exists="overlay") as writer:
            startrow = 5  # 0-indexed in pandas but Excel is 1; we want after 5 header rows
            merged.to_excel(writer, sheet_name="Data Base", index=False, startrow=startrow)

        logging.info("Main workbook updated: %s", main_path)
        return True
    except Exception as exc:
        logging.exception("Failed updating main workbook: %s", exc)
        return False
