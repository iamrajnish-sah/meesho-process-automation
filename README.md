# Meesho Monthly Update Pipeline

Automates the monthly Excel reporting workflow for Meesho category-level GMV and Orders data.

## What it does

- Reads a raw CSV (Category, Orders, GMV) and a master workbook
- Populates six downstream Excel sheets for the new month:
  1. Raw Working Sheet
  2. Raw Data (with MoM and YoY)
  3. Meesho Prelim View
  4. Error Margin
  5. FK Prelim View
  6. Client Prelim
- Flags MoM/YoY changes above the threshold in red with a "Please re-check again" note
- Never modifies the original master file — always saves to a new output file

## Quick start (Web UI)

1. Install dependencies:
   ```
   pip install openpyxl flask werkzeug google-generativeai
   ```

2. Start the server:
   ```
   python web_app.py
   ```

3. Open your browser at `http://127.0.0.1:8080`

4. Upload Raw Data CSV → Upload Master Workbook → click **Process the Data** → Download

## Monthly loop

- Run May → download output → use that output file as master for June
- No code changes needed between months

## Raw data CSV format

```
Category,Orders,GMV
Automotive Accessories,4900000,876300000
Beauty and Personal Care,19930000,3505100000
...
```

GMV must be in absolute INR units (not Crores).

## Project structure

```
web_app.py          — Flask web UI
run_pipeline.py     — CLI entry point and pipeline orchestrator
raw_working_sheet.py
raw_data.py
meesho_prelim.py
error_margin.py
fk_prelim.py
client_prelim.py
shared/
  anchors.py        — Sheet structural anchors
  contracts.py      — Data contracts between pipeline steps
  formula_utils.py  — Column arithmetic and formula utilities
  month_utils.py    — Month label parsing and validation
  style_utils.py    — Cell styling utilities
  raw_input.py      — Raw CSV/Excel loading
  workbook_io.py    — Safe workbook loading
gemini_ingest.py    — Gemini API integration for file conversion
web_uploads.py      — Session-scoped upload storage
templates/
  index.html        — Web UI template
requirements-web.txt
```

## CLI usage

```
python run_pipeline.py MasterFile.xlsx --month "May'26" --raw-data raw_may26.csv
python run_pipeline.py MasterFile.xlsx --month "May'26" --raw-data raw_may26.csv --dry-run
```
