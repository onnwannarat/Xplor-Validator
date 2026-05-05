# Xplor Data Migration Tools

A suite of five Streamlit tools for validating, cleaning, and preparing QikKids data for Xplor import. All tools share a single `serviceIDs.csv` upload that persists as you navigate between pages.

---

## Requirements

- Python 3.11+
- Dependencies listed in `requirements.txt`

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## Running the App

```bash
py -m streamlit run Home.py
```

The app opens at `http://localhost:8501` by default.

---

## serviceIDs.csv

Every tool requires a `serviceIDs.csv` file. Upload it once in the sidebar — it is shared across all pages for the session.

Required columns:

| Column | Description |
|---|---|
| `Service Name` | Xplor service name |
| `Service Type` | e.g. Long Day Care, OSHC |
| `QKDBID` | QikKids database ID |
| `QKServiceID` | QikKids service legacy ID |
| `Xplor Service ID` | Xplor service ID |

---

## Tools

### 1 — Xplor Validator

Validates a full QikKids migration export (CSV or XLSX) against Xplor import rules.

**What it does:**
- Maps QK legacy service IDs to Xplor service IDs
- Normalises state abbreviations and date formats
- Deduplicates parent emails within a row
- Removes `Child_CRN` when it matches a parent CRN (a child and parent cannot share the same CRN) — logged as **FIXED** in the audit report
- Flags errors (must fix), warnings (review), and auto-fixes (applied automatically)
- Cross-checks parents against other already-imported service files to catch duplicate profiles
- Saves split CSVs per service, a full audit report, a client-facing audit report, and a duplicate parents report

**Inputs:**
- Migration CSV or XLSX
- Destination folder (output files saved to `output/` inside it)
- (Optional) existing service files for cross-service duplicate check

**Severity levels:**

| Level | Meaning |
|---|---|
| 🔴 ERROR | Must be fixed before importing |
| 🟡 WARNING | Review recommended |
| 🟢 FIXED | Auto-corrected by the tool |

---

### 2 — Room & Fee Names

Compares the fee names and room names used in QikKids booking exports against the names defined in Xplor, and identifies mismatches with word-overlap suggestions.

**What it does:**
- Loads Xplor fee names per service ID
- Loads Xplor room names per centre name
- Scans one or more QikKids booking exports for fee and room names used
- Produces an Excel report with two sheets: **Fee Mismatches** and **Room Mismatches**
- Each mismatch row includes a "Possible Match in Xplor" suggestion

**Inputs:**
- Xplor Active Fees CSV (columns: `Service ID`, `Fee Name`)
- Xplor Active Rooms CSV (columns: `Centre_Name`, `Room_Name`)
- One or more QikKids bookings CSV/XLSX (recurring and/or casual)

**Output:** `Fee_Room_Name_Mismatch_Report.xlsx` (download button + optional save to folder)

---

### 3 — Bookings Import

Processes QikKids recurring and casual booking exports, deduplicates rows, removes schedule conflicts, maps service IDs, and produces split CSV files per service ready for Xplor import.

**What it does:**
- Combines all uploaded booking files into one dataset
- Detects and removes exact duplicate booking rows (all fields identical)
- Detects and removes recurring schedule overlaps (same child, overlapping date range, shared booked day)
- Detects and removes casual bookings that fall within an existing recurring booking's pattern
- Maps QK service IDs to Xplor service IDs using `serviceIDs.csv`
- Reformats dates to `DD/MM/YYYY` and sets default end dates
- Saves split CSVs to `Output/Recurring/` and `Output/Casual/` inside the chosen output folder
- Saves a duplicate bookings report and a removed casual overlaps report as Excel files

**Upload order for Xplor import:**
1. Upload all files in `Output/Recurring/` first
2. Upload all files in `Output/Casual/` after

---

### 4 — Payment Plans

Validates a payment plan CSV against Onboarding Tool error rules, auto-fixes common issues, and splits output by service.

**What it does:**
- Strips trailing whitespace from all fields
- Converts date values to `DD/MM/YYYY` format
- Converts full weekday names to 3-letter abbreviations (e.g. `Monday` → `Mon`)
- Validates against all Onboarding Tool error rules:

| Error Key | Description |
|---|---|
| `ERROR_INVALID_PAYMENT_DAY` | Weekday is Saturday or Sunday |
| `ERROR_MISSING_BOOKING_START_DATE` | Start Date is empty |
| `ERROR_MISSING_PAYMENT_DAY` | Weekday is empty |
| `ERROR_MISSING_GUARDIAN` | Parent first or last name is missing |
| `ERROR_INVALID_FREQUENCY` | Billing Cycle is not Weekly / Fortnightly / Monthly |
| `ERROR_NEGATIVE_DIRECT_DEBIT_LIMIT` | Direct Debit Limit is negative |
| `ERROR_NEGATIVE_FIXED_LIMIT` | Fixed Amount is negative |
| `ERROR_ONLY_ONE_AMOUNT_ALLOWED` | Both Limit and Fixed Amount are greater than zero |
| `MANUAL_PLAN_NOT_MONDAY` | Manual (Paused) plan start date is not a Monday |

- Saves a cleaned CSV, a colour-coded Excel error report, and split CSVs per service

**Column mapping:** If your CSV uses different column names, expand the **Column Mapping** section on the page to remap before running.

---

### 5 — Balance Adjustments

Reads balance adjustment data, maps centre names to validated Xplor service names, removes demo accounts, and produces one styled XLSX file per centre using the upload template.

**What it does:**
- Accepts CSV, XLSX, and HTML-exported XLS files
- Parses `Account Name` field in "Last Name, First Name" format
- Handles both `Credit`/`Owing` columns and a single signed `Amount Due` column
- Removes rows where `Account Name` contains "demo parent"
- Only produces output files for centres whose name matches a `Service Name` in `serviceIDs.csv`
- Applies the styling and layout from the Balance Adjustments Details Upload Template

**Inputs:**
- One or more balance adjustment files (CSV/XLSX/XLS)
- Template is bundled in `assets/Balance Adjustments Details Upload Template.xlsx` — no upload needed
- (Optional) override template via the Advanced expander

**Output:** One `<Centre Name>_Balance_Import.xlsx` per centre

---

## Project Structure

```
Validator/
├── Home.py                               ← App entry point (landing page)
├── launcher.py                           ← PyInstaller entry point (do not modify)
├── validator_v2.py                       ← Core validation engine
├── requirements.txt
├── xplor_validator.spec                  ← PyInstaller build spec
│
├── pages/
│   ├── 1_Xplor_Validator.py
│   ├── 2_Room_Fee_Names.py
│   ├── 3_Bookings_Import.py
│   ├── 4_Payment_Plans.py
│   └── 5_Balance_Adjustments.py
│
├── shared/
│   ├── service_map.py                    ← Shared sidebar + session state helper
│   └── styles.py                         ← Shared CSS
│
├── scripts/
│   ├── check_names.py
│   ├── prepare_bookings_import.py
│   ├── payment_plan_checker.py
│   └── process_balance_adjustments.py
│
└── assets/
    └── Balance Adjustments Details Upload Template.xlsx
```

---

## Packaging (PyInstaller)

Build a standalone `.exe` on Windows:

```bash
pyinstaller xplor_validator.spec
```

The packaged app launches automatically in the default browser. Do not modify `launcher.py` or `xplor_validator.spec`.

---

## Notes

- None of the tools modify your original input files.
- The `serviceIDs.csv` upload persists in session state — upload it once on any page and it remains available on all other pages until the browser tab is closed.
- The `app.py` file is kept as a reference copy of the original single-page validator. The active version is `pages/1_Xplor_Validator.py`.
