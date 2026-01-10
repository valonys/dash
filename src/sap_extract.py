from __future__ import annotations

import logging
import os
import platform
import time
from pathlib import Path


def is_windows() -> bool:
    return platform.system().lower() == "windows"


def extract_sap_data(variant: str, output_dir: str | Path, sap_system: str) -> tuple[bool, str | None]:
    """Extract data from SAP IW39 into _woStatus.xlsx.

    Returns (success, path_str or None). On non-Windows, returns (False, None) with info logged.
    """
    output_path = Path(output_dir).resolve() / "_woStatus.xlsx"
    if not is_windows():
        logging.warning("SAP GUI extraction is only supported on Windows. Provide _woStatus.xlsx manually.")
        return False, None

    try:
        import win32com.client  # type: ignore
    except Exception as exc:  # pragma: no cover
        logging.error("pywin32 is not available: %s", exc)
        return False, None

    try:  # pragma: no cover - requires Windows + SAP GUI
        sap_gui = None
        application = None
        try:
            sap_gui = win32com.client.GetObject("SAPGUI")
            application = sap_gui.GetScriptingEngine
        except Exception:
            logging.info("Starting SAP Logon...")
            os.system(r'"C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\saplogon.exe"')
            time.sleep(2)
            sap_gui = win32com.client.GetObject("SAPGUI")
            application = sap_gui.GetScriptingEngine

        # Try to reuse an existing session
        session = None
        for conn in application.Children:
            for sess in conn.Children:
                session = sess
                break
        if session is None:
            connection = application.OpenConnection(sap_system, True)
            session = connection.Children(0)

        session.FindById("wnd[0]").Maximize()
        session.StartTransaction("IW39")

        # Example of loading a variant and exporting to Excel
        session.FindById("wnd[0]/mbar/menu[2]/menu[0]/menu[0]").Select()
        session.FindById("wnd[1]/usr/txtV-LOW").Text = variant
        session.FindById("wnd[1]/tbar[0]/btn[8]").Press()
        session.FindById("wnd[0]/tbar[1]/btn[8]").Press()  # Execute
        session.FindById("wnd[0]/tbar[1]/btn[16]").Press()  # Export
        session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
        session.FindById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").Select()
        session.FindById("wnd[1]/tbar[0]/btn[0]").Press()
        session.FindById("wnd[1]/tbar[0]/btn[0]").Press()

        # Save the exported workbook
        excel = win32com.client.Dispatch("Excel.Application")
        time.sleep(2)
        for wb in excel.Workbooks:
            try:
                wb.SaveAs(str(output_path), FileFormat=51)
                wb.Close(False)
                logging.info("Saved SAP export to %s", output_path)
                break
            except Exception:  # keep trying others
                continue

        try:
            session.FindById("wnd[0]").Close()
            session.FindById("wnd[1]/usr/btnSPOP-OPTION1").Press()
        except Exception:
            pass

        return output_path.exists(), str(output_path) if output_path.exists() else None
    except Exception as exc:  # pragma: no cover
        logging.exception("SAP extraction failed: %s", exc)
        return False, None
