# Xplor Data Migration Validator — Infrastructure Handover Guide

## Overview

This is a **Streamlit web application** that validates and transforms CSV files prepared for Xplor data migration. It is designed to run as an Azure App Service.

The tool:
- Validates child/parent migration CSVs against Xplor's mandatory field requirements
- Auto-fixes common issues (Service ID mapping, state normalisation, duplicate email removal)
- Splits validated output into one CSV per service
- Generates Excel audit reports (internal + client-facing)

---

## Files to Deploy

Share **only** these files with the Infrastructure team:

| File | Purpose |
|------|---------|
| `app_azure.py` | **Main entry point** — Streamlit UI optimised for Azure (file upload/download, no local folder picker) |
| `validator_v2.py` | Core validation and transformation engine — imported by `app_azure.py` |
| `requirements.txt` | Python package dependencies |
| `.streamlit/config.toml` | Streamlit server configuration (port 8000, headless mode) |
| `serviceIDs.csv` | Service ID mapping table (QK legacy IDs → Xplor IDs + service names) — **must be kept up to date** |

> **Do not include:** `app.py` (local-only version), `launcher.py` (Windows .exe launcher), `transform.py` (older standalone script), `build/`, or `dist/` folders.

---

## Azure App Service — Deployment Steps

### 1. Runtime

- **Python 3.11+** (Python 3.12 recommended)

### 2. Install Dependencies

```bash
pip install -r requirements.txt
```

### 3. Startup Command

Set this as the App Service **startup command**:

```bash
python -m streamlit run app_azure.py --server.port 8000 --server.address 0.0.0.0 --server.headless true
```

The `.streamlit/config.toml` file also pre-configures these settings — include it in the deployment so Streamlit picks it up automatically.

### 4. Port Configuration

The app listens on **port 8000**. Ensure the App Service is configured to route inbound traffic to this port.

### 5. File Structure in the Deployment Root

```
/
├── app_azure.py
├── validator_v2.py
├── requirements.txt
├── serviceIDs.csv
└── .streamlit/
    └── config.toml
```

---

## How the Application Works

### User Workflow

1. User uploads a **migration CSV** (the child/parent data file prepared for Xplor import).
2. *(Optional)* User uploads a custom **serviceIDs.csv** to override the default service mapping.
3. The app validates the data and displays a summary of issues and fixes.
4. User downloads:
   - **Ready-to-import CSVs** (one per service, bundled as a ZIP)
   - **Validation audit report** (.xlsx, internal use)
   - **Client report** (.xlsx, client-facing summary)

### What the Validator Checks & Fixes

| Category | Behaviour |
|----------|-----------|
| Mandatory fields | Reports `ERROR` for missing required fields |
| Service ID mapping | Auto-maps QK legacy Service IDs to Xplor Service IDs using `serviceIDs.csv` |
| State normalisation | Converts free-text state names (e.g. "New South Wales") to AU abbreviations (e.g. "NSW") |
| Email deduplication | Removes duplicate emails within Parent / Emergency Contact fields |
| CRN format | Validates Australian CRN format |
| Date fields | Validates date formats (DOB, enrolment dates, etc.) |
| Boolean fields | Accepts Yes/No/True/False/1/0 |
| Duplicate detection | Flags duplicate Child Legacy IDs and duplicate Parent CRNs across services |
| Waitlist records | Exempts waitlist records from `Enrolment_Start_Date` requirement |
| Field length limits | Reports fields exceeding Xplor's column size limits |

---

## serviceIDs.csv Format

This file maps QikKids (legacy) service IDs to Xplor service IDs. It must be a three-column CSV:

```csv
QK_Service_ID,Xplor_Service_ID,Service_Name
182,XPL-001,Acacia Ridge OSHC
183,XPL-002,Sunnybank Hills Early Learning
```

- The app bundles a default `serviceIDs.csv`. Users can upload a custom one at runtime to override it.
- Keep this file updated as new services are onboarded.

---

## Key Python API (for scripting / runbook integration)

If the Infrastructure team wants to call the validator programmatically rather than via the UI, the following public functions are available in `validator_v2.py`:

```python
from validator_v2 import (
    ServiceMapping,
    run_v2_from_bytes,
    write_split_csvs,
    write_excel_report,
    write_client_excel_report,
)

# Load service mapping
service_map = ServiceMapping("path/to/serviceIDs.csv")

# Run validation on a CSV file loaded as bytes
with open("migration_data.csv", "rb") as f:
    file_bytes = f.read()

recorder, all_rows, fieldnames = run_v2_from_bytes(
    file_bytes,
    "migration_data.csv",
    service_map,
)

# Write output files
write_split_csvs(all_rows, fieldnames, output_dir="./output")
write_excel_report(recorder, all_rows, "audit_report.xlsx", service_map)
write_client_excel_report(recorder, all_rows, "client_report.xlsx", service_map)
```

### `run_v2_from_bytes` return values

| Value | Type | Description |
|-------|------|-------------|
| `recorder` | `IssueRecorder` | All issues and fixes found during validation |
| `all_rows` | `list[dict]` | Transformed rows (one dict per CSV row) |
| `fieldnames` | `list[str]` | Column names from the input CSV |

---

## Dependencies

```
streamlit>=1.35.0
pandas>=2.0.0
openpyxl>=3.1.0
```

No external API calls are made. All processing is local/in-memory.

---

## Contact

For questions about the validation logic or migration data requirements, contact Amy - Onboarding team.
