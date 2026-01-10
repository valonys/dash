# SAP KPI Report Generator (Python + Streamlit)

Modular Python reimplementation of VBA macros for SAP IW39 extraction, Excel processing, and KPI sheet generation with a Streamlit web UI.

## Requirements
- Python 3.10+
- On Windows (optional): SAP GUI + GUI scripting enabled for automated extraction
- Main workbook (.xlsx) with a `Data Base` sheet

## Setup
```bash
pip install -r requirements.txt
```

## Run UI
```bash
streamlit run app.py
```

## Usage
- Upload the main workbook and, if SAP extraction is not available on this host, also upload `_woStatus.xlsx`.
- On Windows with SAP GUI, provide the variant and system; the app will attempt to extract IW39 and proceed automatically.

## Notes
- Pivot tables are not refreshed natively; rebuild with pandas if needed or refresh via Excel COM on Windows.
