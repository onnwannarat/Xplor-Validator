"""
Xplor Data Migration — Validation Tool v2 (Streamlit UI)
=========================================================
Run locally:
    streamlit run app.py

Deploy to Azure App Service — startup command:
    python -m streamlit run app.py --server.port 8000 --server.address 0.0.0.0 --server.headless true

Author: Amy Boonyaratanakornkit (Onboarding team)
"""

import hashlib
import io
import os
import sys
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from validator_v2 import (
    IssueRecorder,
    ServiceMapping,
    run_v2_from_bytes,
    write_client_excel_report,
    write_excel_report,
    write_split_csvs,
)

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Xplor Data Migration Validator",
    page_icon="🔍",
    layout="wide",
)

# ─────────────────────────────────────────────────────────────────────────────
# CUSTOM STYLING
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("""
<style>
    #MainMenu { visibility: hidden; }
    footer    { visibility: hidden; }

    .tool-header {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
        padding: 2rem 2.5rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        color: white;
    }
    .tool-header h1 { margin: 0; font-size: 1.8rem; font-weight: 700; }
    .tool-header p  { margin: 0.4rem 0 0; opacity: 0.75; font-size: 0.95rem; }

    .status-banner {
        padding: 0.9rem 1.5rem;
        border-radius: 8px;
        font-weight: 600;
        margin-bottom: 1.5rem;
        font-size: 1rem;
    }
    .status-error   { background: #fef2f2; color: #991b1b; border-left: 4px solid #dc2626; }
    .status-warning { background: #fffbeb; color: #92400e; border-left: 4px solid #d97706; }
    .status-ok      { background: #f0fdf4; color: #14532d; border-left: 4px solid #16a34a; }

    .how-to-step {
        display: flex;
        align-items: flex-start;
        gap: 0.75rem;
        margin-bottom: 0.9rem;
    }
    .step-num {
        background: #0f3460;
        color: white;
        border-radius: 50%;
        width: 24px;
        height: 24px;
        min-width: 24px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.75rem;
        font-weight: 700;
        margin-top: 2px;
    }
    .step-text { font-size: 0.88rem; line-height: 1.5; color: #1e293b; }
    .step-text strong { color: #0f3460; }

    .stDataFrame { font-size: 0.875rem; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _find_service_map_path() -> str:
    if hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, "serviceIDs.csv")
    return str(Path(__file__).parent / "serviceIDs.csv")


@st.cache_resource
def _load_default_service_map() -> ServiceMapping:
    return ServiceMapping(_find_service_map_path())


@st.cache_data(show_spinner=False)
def _load_service_map_from_bytes(file_bytes: bytes) -> ServiceMapping:
    with tempfile.NamedTemporaryFile(suffix=".csv", delete=False, mode="wb") as tmp:
        tmp.write(file_bytes)
        tmp_path = tmp.name
    try:
        return ServiceMapping(tmp_path)
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass


@st.cache_data(show_spinner=False)
def _run_and_build_outputs(
    file_bytes: bytes,
    filename: str,
    svc_map_bytes: bytes | None,
) -> tuple[IssueRecorder, bytes, bytes, bytes]:
    """
    Runs the full v2 pipeline and returns everything the UI needs.
    Cached by (file content + serviceIDs content) so re-runs only when
    either file actually changes.

    Returns: recorder, zip_bytes, excel_bytes, client_excel_bytes
    """
    if svc_map_bytes is not None:
        service_map = _load_service_map_from_bytes(svc_map_bytes)
    else:
        service_map = _load_default_service_map()

    recorder, all_rows, fieldnames = run_v2_from_bytes(file_bytes, filename, service_map)

    # Write split CSVs to a temp dir — only to capture row_num_map for reports
    with tempfile.TemporaryDirectory() as tmp_dir:
        _, row_num_map = write_split_csvs(all_rows, fieldnames, tmp_dir)

        # Build ZIP of split CSVs for download
        zip_buf = io.BytesIO()
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for csv_file in Path(tmp_dir).glob("ready_to_import_*.csv"):
                zf.write(csv_file, csv_file.name)
        zip_bytes = zip_buf.getvalue()

    # Build Excel audit report bytes
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_xl:
        tmp_xl_path = tmp_xl.name
    try:
        write_excel_report(recorder, all_rows, tmp_xl_path, service_map, row_num_map)
        excel_bytes = Path(tmp_xl_path).read_bytes()
    finally:
        try:
            os.unlink(tmp_xl_path)
        except OSError:
            pass

    # Build client Excel report bytes
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp_cl:
        tmp_cl_path = tmp_cl.name
    try:
        write_client_excel_report(recorder, all_rows, tmp_cl_path, service_map, row_num_map)
        client_excel_bytes = Path(tmp_cl_path).read_bytes() if Path(tmp_cl_path).exists() else b""
    finally:
        try:
            os.unlink(tmp_cl_path)
        except OSError:
            pass

    return recorder, zip_bytes, excel_bytes, client_excel_bytes


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    # ── Service ID mapping file ───────────────────────────────────────────────
    st.markdown("## ⚙️ Service ID Mapping")
    st.markdown("---")
    st.markdown(
        "Upload a **serviceIDs.csv** for this project. "
        "If not uploaded, the built-in default file will be used.",
        help="Required columns: Service Name, Service Type, QKDBID, QKServiceID, Xplor Service ID",
    )
    uploaded_service_map = st.file_uploader(
        "serviceIDs.csv (optional)",
        type=["csv"],
        key="service_map_upload",
        label_visibility="collapsed",
    )
    if uploaded_service_map:
        st.success(f"✅ Using: **{uploaded_service_map.name}**")
    else:
        default_path = _find_service_map_path()
        if Path(default_path).exists():
            st.info("ℹ️ Using built-in default serviceIDs.csv")
        else:
            st.warning("⚠️ No serviceIDs.csv found — service ID mapping will be skipped.")

    st.markdown("---")

    # ── How to use ────────────────────────────────────────────────────────────
    st.markdown("## 📖 How to Use")
    st.markdown("---")

    st.markdown("""
<div class="how-to-step">
  <div class="step-num">1</div>
  <div class="step-text"><strong>(Optional) Upload serviceIDs.csv</strong><br>
  If this project uses a different service ID mapping, upload it above first.</div>
</div>

<div class="how-to-step">
  <div class="step-num">2</div>
  <div class="step-text"><strong>Upload your migration file</strong><br>
  Drag and drop a <strong>CSV</strong> or <strong>XLSX</strong> onto the upload box.</div>
</div>

<div class="how-to-step">
  <div class="step-num">3</div>
  <div class="step-text"><strong>Wait for validation</strong><br>
  The tool automatically checks all rows, maps service IDs, normalises states, and deduplicates emails.</div>
</div>

<div class="how-to-step">
  <div class="step-num">4</div>
  <div class="step-text"><strong>Read the results</strong><br>
  A summary shows <strong style="color:#dc2626">Errors</strong>, <strong style="color:#d97706">Warnings</strong>, and <strong style="color:#16a34a">Auto-Fixes</strong>.</div>
</div>

<div class="how-to-step">
  <div class="step-num">5</div>
  <div class="step-text"><strong>Download the outputs</strong><br>
  <em>Audit Report (.xlsx)</em> — colour-coded report per service.<br>
  <em>Client Report (.xlsx)</em> — duplicate emails and redundant ECs.<br>
  <em>Import-Ready CSVs (.zip)</em> — files ready to upload to Xplor.</div>
</div>

<div class="how-to-step">
  <div class="step-num">6</div>
  <div class="step-text"><strong>Fix & re-upload</strong><br>
  Correct any errors in your spreadsheet, re-export, and upload again.</div>
</div>
""", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 🚦 Severity levels")
    st.markdown("""
**🔴 ERROR** — Must be fixed before importing.

**🟡 WARNING** — Review recommended before importing.

**🟢 FIXED** — Automatically corrected by the tool.
""")

    st.markdown("---")
    st.markdown("### ⚙️ Auto-fixes applied")
    st.markdown("""
- **Service ID mapping** — QK legacy IDs → Xplor Service IDs
- **State normalisation** — free-text → standard abbreviations
- **Email deduplication** — removes duplicate emails within a row
- **CRN fix** — clears parent CRN that matches child CRN
- **Blank first name** — copies last name into blank first name fields
""")

    st.markdown("---")
    st.markdown("### 📋 Validation checklist")
    with st.expander("View full checklist"):
        st.markdown("""
**Mandatory Fields**
- Child: Service ID, Legacy ID, First/Last Name, DOB, Status, CRN, Room, Enrolment Date (Active/Inactive only)
- Parent: CRN, Legacy ID, DOB, Email (when parent block is populated)
- Emergency contacts: Legacy ID per contact

**Format Checks**
- CRN: 9 digits + 1 letter (e.g. `123456789A`)
- Dates: `YYYY-MM-DD` or `D/MM/YYYY` or `DD/MM/YYYY`
- Emails: standard format
- Phone: Australian mobile (04xx) or landline
- State: NSW, VIC, QLD, SA, WA, TAS, ACT, NT
- Postcodes: 4 digits
- Booleans: 0/1, Yes/No, True/False

**Business Logic**
- Child DOB not in the future
- Enrolment date not before DOB
- Child CRN ≠ Parent CRN
- Waitlist logic

**Duplicate Checks**
- Child Legacy ID and CRN unique
- Parent Email / CRN consistency
""")

    st.markdown("---")
    st.caption("Xplor Data Migration Validator · Does not modify your original file.")

# ─────────────────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="tool-header">
    <h1>🔍 Xplor Data Migration Validator</h1>
    <p>Upload your migration CSV or XLSX to validate and transform it for Xplor import.</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# FILE UPLOAD
# ─────────────────────────────────────────────────────────────────────────────

uploaded_file = st.file_uploader(
    "Upload your migration file (CSV or XLSX)",
    type=["csv", "xlsx", "xls"],
    help="Your original file is not modified.",
)

if not uploaded_file:
    st.info("👆 Upload a migration file above to begin.  \n💡 See the **How to Use** guide in the sidebar if you need help.")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# RUN VALIDATION
# ─────────────────────────────────────────────────────────────────────────────

svc_map_bytes = uploaded_service_map.read() if uploaded_service_map else None

with st.spinner("Validating and transforming your file — please wait…"):
    try:
        recorder, zip_bytes, excel_bytes, client_excel_bytes = _run_and_build_outputs(
            uploaded_file.read(),
            uploaded_file.name,
            svc_map_bytes,
        )
    except Exception as exc:
        st.error(f"Could not process the file: {exc}")
        st.stop()

errors   = recorder.error_count()
warnings = recorder.warning_count()
fixed    = recorder.fixed_count()
total    = len(recorder.issues)
df       = recorder.to_dataframe()
ts       = datetime.now().strftime("%Y%m%d_%H%M%S")

# ─────────────────────────────────────────────────────────────────────────────
# SUMMARY METRICS
# ─────────────────────────────────────────────────────────────────────────────

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Total Entries", total)
with col2:
    st.metric("🔴 Errors", errors, help="Must be resolved before importing.")
with col3:
    st.metric("🟡 Warnings", warnings, help="Review before importing.")
with col4:
    st.metric("🟢 Auto-Fixed", fixed, help="Automatically corrected by the tool.")

if errors > 0:
    st.markdown(
        f'<div class="status-banner status-error">❌ {errors} error(s) found — please resolve all errors before importing.</div>',
        unsafe_allow_html=True,
    )
elif warnings > 0:
    st.markdown(
        f'<div class="status-banner status-warning">⚠️ {warnings} warning(s) found — please review before importing.</div>',
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        '<div class="status-banner status-ok">✅ No errors or warnings — this file appears ready for import.</div>',
        unsafe_allow_html=True,
    )

# ─────────────────────────────────────────────────────────────────────────────
# DOWNLOAD BUTTONS
# ─────────────────────────────────────────────────────────────────────────────

dl_col1, dl_col2, dl_col3, dl_col4 = st.columns(4)

with dl_col1:
    st.download_button(
        label="⬇️ Audit Report (CSV)",
        data=recorder.to_csv_bytes(),
        file_name=f"validation_audit_report_{ts}.csv",
        mime="text/csv",
        use_container_width=True,
        help="Flat CSV listing all issues and auto-fixes.",
    )

with dl_col2:
    st.download_button(
        label="⬇️ Audit Report (Excel)",
        data=excel_bytes,
        file_name=f"validation_audit_report_{ts}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        disabled=not excel_bytes,
        help="Colour-coded Excel report with one tab per service.",
    )

with dl_col3:
    st.download_button(
        label="⬇️ Client Report (Excel)",
        data=client_excel_bytes,
        file_name=f"client_audit_report_{ts}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        disabled=not client_excel_bytes,
        help="Client-facing report: duplicate emails and redundant emergency contacts.",
    )

with dl_col4:
    st.download_button(
        label="⬇️ Import-Ready CSVs (ZIP)",
        data=zip_bytes,
        file_name=f"import_ready_{ts}.zip",
        mime="application/zip",
        use_container_width=True,
        disabled=not zip_bytes,
        help="Transformed CSVs split by service, ready to upload to Xplor.",
    )

st.divider()

# ─────────────────────────────────────────────────────────────────────────────
# RESULTS TABLE
# ─────────────────────────────────────────────────────────────────────────────

if total == 0:
    st.success("No issues to display.")
    st.stop()

st.subheader("Validation Results")

filter_col1, filter_col2, filter_col3 = st.columns([1, 2, 2])

with filter_col1:
    severity_filter = st.selectbox("Filter by Severity", options=["All", "ERROR", "WARNING", "FIXED"], index=0)

with filter_col2:
    child_options = ["All children"] + sorted(df["Child_Name"].dropna().unique().tolist())
    child_filter  = st.selectbox("Filter by Child", options=child_options, index=0)

with filter_col3:
    field_options = ["All fields"] + sorted(df["Field"].dropna().unique().tolist())
    field_filter  = st.selectbox("Filter by Field", options=field_options, index=0)

filtered_df = df.copy()
if severity_filter != "All":
    filtered_df = filtered_df[filtered_df["Severity_Level"] == severity_filter]
if child_filter != "All children":
    filtered_df = filtered_df[filtered_df["Child_Name"] == child_filter]
if field_filter != "All fields":
    filtered_df = filtered_df[filtered_df["Field"] == field_filter]

st.caption(f"Showing {len(filtered_df):,} of {total:,} entries")


def colour_severity(val):
    if val == "ERROR":
        return "color: #dc2626; font-weight: 700;"
    elif val == "WARNING":
        return "color: #d97706; font-weight: 700;"
    elif val == "FIXED":
        return "color: #16a34a; font-weight: 700;"
    return ""


_cell_count = filtered_df.shape[0] * filtered_df.shape[1]
pd.set_option("styler.render.max_elements", max(_cell_count, 262144))

styled = (
    filtered_df
    .reset_index(drop=True)
    .style
    .map(colour_severity, subset=["Severity_Level"])
)

st.dataframe(
    styled,
    use_container_width=True,
    height=min(600, 60 + len(filtered_df) * 35),
    column_config={
        "Row":               st.column_config.NumberColumn("Row",        width="small"),
        "Child_Name":        st.column_config.TextColumn("Child",        width="medium"),
        "Field":             st.column_config.TextColumn("Field",        width="medium"),
        "Issue_Description": st.column_config.TextColumn("Issue",        width="large"),
        "Severity_Level":    st.column_config.TextColumn("Severity",     width="small"),
        "Action_Taken":      st.column_config.TextColumn("Action Taken", width="large"),
    },
)

with st.expander("📊 Issue breakdown by field"):
    breakdown = (
        df.groupby(["Field", "Severity_Level"])
        .size()
        .reset_index(name="Count")
        .sort_values(["Count"], ascending=False)
    )
    st.dataframe(breakdown, use_container_width=True, hide_index=True)
