from __future__ import annotations

import logging
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd


INSPECTION_TYPES: List[str] = [
    "Pressure Vessel (VII)",
    "Pressure Vessel (VIE)",
    "Pressure Safety Device",
    "Piping",
    "FU Items",
    "Structure",
    "Flare TIP",
    "Lifting",
    "Non Structural Tank",
    "Corrosion Monitoring",
    "Campaign",
    "Intelligent Pigging",
    "Flame Arrestor",
]


def _compute_sheet(inspection_type: str, db_df: pd.DataFrame) -> Tuple[str, pd.DataFrame]:
    df = db_df.copy()
    if "Item Class" in df.columns:
        mask = df["Item Class"].astype(str).str.strip() == inspection_type
        sheet_df = df.loc[mask].copy()
    else:
        # Fallback: if there is a generic classifier column 'B'
        mask = df.columns[1] if len(df.columns) > 1 else df.columns[0]
        sheet_df = df[df[mask].astype(str).str.strip() == inspection_type].copy()

    if sheet_df.empty:
        return inspection_type, sheet_df

    # Re-map minimal set of expected columns if present
    mapping = {
        "WO Number": next((c for c in sheet_df.columns if c in ("WO Number", "Order", "Order Number")), None),
        "Status": next((c for c in sheet_df.columns if c in ("SECE STATUS", "Status", "User Status")), None),
        "TAG": next((c for c in sheet_df.columns if c in ("TAG", "Tag", "Technical ID")), None),
        "FL": next((c for c in sheet_df.columns if c in ("FL", "Functional Location")), None),
        "Description": next((c for c in sheet_df.columns if c in ("Description", "Object description")), None),
        "Due Date": next((c for c in sheet_df.columns if c in ("Due Date", "Next Insp", "Basic finish")), None),
    }

    out = pd.DataFrame()
    out["#"] = range(1, len(sheet_df) + 1)
    for target, src in mapping.items():
        if src:
            out[target] = sheet_df[src].values

    # Derive Due Month and QCAP/EXDO flag
    if "Due Date" in out.columns:
        out["Due Month"] = pd.to_datetime(out["Due Date"], errors="coerce").dt.month
    if "Status" in out.columns:
        out["QCAP/EXDO"] = out["Status"].apply(lambda s: "YES" if str(s) in ("QCAP", "EXDO") else "NO")

    return inspection_type, out


def generate_all_sheets(main_wb_path: str | Path) -> bool:
    try:
        main_path = Path(main_wb_path).resolve()
        db_df = pd.read_excel(main_path, sheet_name="Data Base", skiprows=4)
        db_df.columns = [str(c).strip() for c in db_df.columns]

        # Compute in parallel, write sequentially
        results: Dict[str, pd.DataFrame] = {}
        with ProcessPoolExecutor() as executor:
            futures = {executor.submit(_compute_sheet, itype, db_df): itype for itype in INSPECTION_TYPES}
            for fut in as_completed(futures):
                itype, out_df = fut.result()
                results[itype] = out_df

        with pd.ExcelWriter(main_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            for itype, out_df in results.items():
                if out_df is not None and not out_df.empty:
                    out_df.to_excel(writer, sheet_name=itype, index=False)

        logging.info("Generated %d categorized sheets", sum(1 for v in results.values() if not v.empty))
        return True
    except Exception as exc:
        logging.exception("Failed generating categorized sheets: %s", exc)
        return False
