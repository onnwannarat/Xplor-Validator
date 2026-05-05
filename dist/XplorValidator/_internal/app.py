"""
Xplor Data Migration — Validation Tool (Streamlit UI)
======================================================
Run with:
    streamlit run app.py
"""

import io
from datetime import datetime

import pandas as pd
import streamlit as st

from validator import run_validation

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
    /* Hide the default Streamlit menu and footer for a cleaner look */
    #MainMenu { visibility: hidden; }
    footer    { visibility: hidden; }

    /* Header bar */
    .tool-header {
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
        padding: 2rem 2.5rem;
        border-radius: 12px;
        margin-bottom: 1.5rem;
        color: white;
    }
    .tool-header h1 { margin: 0; font-size: 1.8rem; font-weight: 700; }
    .tool-header p  { margin: 0.4rem 0 0; opacity: 0.75; font-size: 0.95rem; }

    /* Status banner */
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

    /* Sidebar how-to styling */
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

    /* Severity colour in table */
    .stDataFrame { font-size: 0.875rem; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR — HOW TO USE
# ─────────────────────────────────────────────────────────────────────────────

with st.sidebar:
    st.markdown("## 📖 How to Use")
    st.markdown("---")

    st.markdown("""
<div class="how-to-step">
  <div class="step-num">1</div>
  <div class="step-text"><strong>Prepare your file</strong><br>
  Export your migration data as a <strong>CSV (UTF-8)</strong> file from the onboarding spreadsheet.</div>
</div>

<div class="how-to-step">
  <div class="step-num">2</div>
  <div class="step-text"><strong>Upload the CSV</strong><br>
  Drag and drop the CSV file onto the upload box on the right, or click <em>Browse files</em> to select it.</div>
</div>

<div class="how-to-step">
  <div class="step-num">3</div>
  <div class="step-text"><strong>Wait for validation</strong><br>
  The tool will automatically check all rows for errors. This usually takes a few seconds.</div>
</div>

<div class="how-to-step">
  <div class="step-num">4</div>
  <div class="step-text"><strong>Read the results</strong><br>
  A summary shows the number of <strong style="color:#dc2626">Errors</strong> and <strong style="color:#d97706">Warnings</strong> found.</div>
</div>

<div class="how-to-step">
  <div class="step-num">5</div>
  <div class="step-text"><strong>Download the report</strong><br>
  Click <em>Download Audit Report</em> to save a CSV with all issues listed by row and child name.</div>
</div>

<div class="how-to-step">
  <div class="step-num">6</div>
  <div class="step-text"><strong>Fix & re-upload</strong><br>
  Correct the issues in your spreadsheet, re-export the CSV, and upload again to re-validate.</div>
</div>
""", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 🚦 What do the severities mean?")
    st.markdown("""
**🔴 ERROR** — Must be fixed before importing. The record will fail or cause data corruption in Xplor.

**🟡 WARNING** — Review recommended. The record may import but could cause issues later.
""")

    st.markdown("---")
    st.markdown("### 📋 What is checked?")
    with st.expander("View full checklist"):
        st.markdown("""
**Mandatory Fields**
- Child: Service ID, Legacy ID, First/Last Name, DOB, Status, CRN, Room, Enrolment Date (Active/Inactive only)
- Parent: CRN, Legacy ID, DOB, Email (when parent block is populated)
- Emergency contacts: Legacy ID per contact

**Format Checks**
- CRN: must be 9 digits + 1 letter (e.g. `123456789A`)
- Dates: `YYYY-MM-DD` or `D/MM/YYYY` or `DD/MM/YYYY`
- Emails: standard format (e.g. `name@email.com`)
- Phone: Australian mobile (04xx) or landline (0x xxxx xxxx), or 9-digit starting with 4
- State: abbreviation only (NSW, VIC, QLD, SA, WA, TAS, ACT, NT)
- Postcodes: 4 digits
- Booleans: 0/1, Yes/No, or True/False

**Business Logic**
- Child DOB cannot be in the future
- Enrolment date cannot be before DOB
- Child CRN must not equal Parent CRN
- Waitlist children: no Enrolment Date required; flags if date already passed
- Incomplete waitlist: Waitlist status with no guardian info

**Duplicate Checks (cross-row)**
- Child Legacy ID — must be unique
- Child CRN — must be unique
- Parent Email — same email for different CRNs = error
- Parent CRN — same CRN must always use the same email (1 CRN = 1 email)

**Name Pairing**
- First name always requires a last name, and vice versa
""")

    st.markdown("---")
    st.caption("Xplor Data Migration Validator · Read-only · Does not modify your data.")

# ─────────────────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="tool-header">
    <h1>🔍 Xplor Data Migration Validator</h1>
    <p>Upload your migration CSV to audit it against Xplor import rules. This tool is read-only and does not modify your data.</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# FILE UPLOAD
# ─────────────────────────────────────────────────────────────────────────────

uploaded_file = st.file_uploader(
    "Drag and drop your migration CSV here, or click to browse",
    type=["csv"],
    help="Upload the CSV file prepared for Xplor Data Migration. The file is not stored or sent anywhere.",
)

if not uploaded_file:
    st.info("👆 Upload a CSV file above to begin validation.  \n💡 **Tip:** See the **How to Use** guide in the left-hand sidebar if you need help.")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# RUN VALIDATION
# ─────────────────────────────────────────────────────────────────────────────

with st.spinner("Validating your file — please wait…"):
    try:
        recorder = run_validation(uploaded_file)
    except ValueError as exc:
        st.error(f"Could not read the CSV file: {exc}")
        st.stop()

errors   = recorder.error_count()
warnings = recorder.warning_count()
total    = len(recorder.issues)
df       = recorder.to_dataframe()

# ─────────────────────────────────────────────────────────────────────────────
# SUMMARY METRICS
# ─────────────────────────────────────────────────────────────────────────────

col1, col2, col3 = st.columns(3)
with col1:
    st.metric("Total Issues", total)
with col2:
    st.metric("🔴 Errors", errors, help="Must be resolved before importing.")
with col3:
    st.metric("🟡 Warnings", warnings, help="Review before importing.")

# Status banner
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
        '<div class="status-banner status-ok">✅ No issues found — this file appears ready for import.</div>',
        unsafe_allow_html=True,
    )

# ─────────────────────────────────────────────────────────────────────────────
# DOWNLOAD BUTTON
# ─────────────────────────────────────────────────────────────────────────────

timestamp    = datetime.now().strftime("%Y%m%d_%H%M%S")
report_name  = f"validation_audit_report_{timestamp}.csv"
report_bytes = recorder.to_csv_bytes()

st.download_button(
    label="⬇️ Download Audit Report (CSV)",
    data=report_bytes,
    file_name=report_name,
    mime="text/csv",
    use_container_width=True,
)

st.divider()

# ─────────────────────────────────────────────────────────────────────────────
# RESULTS TABLE
# ─────────────────────────────────────────────────────────────────────────────

if total == 0:
    st.success("No issues to display.")
    st.stop()

st.subheader("Validation Results")

# ── Filters ──────────────────────────────────────────────────────────────────

filter_col1, filter_col2, filter_col3 = st.columns([1, 2, 2])

with filter_col1:
    severity_filter = st.selectbox(
        "Filter by Severity",
        options=["All", "ERROR", "WARNING"],
        index=0,
    )

with filter_col2:
    child_options = ["All children"] + sorted(df["Child_Name"].dropna().unique().tolist())
    child_filter  = st.selectbox("Filter by Child", options=child_options, index=0)

with filter_col3:
    field_options = ["All fields"] + sorted(df["Field"].dropna().unique().tolist())
    field_filter  = st.selectbox("Filter by Field", options=field_options, index=0)

# ── Apply filters ─────────────────────────────────────────────────────────────

filtered_df = df.copy()

if severity_filter != "All":
    filtered_df = filtered_df[filtered_df["Severity_Level"] == severity_filter]

if child_filter != "All children":
    filtered_df = filtered_df[filtered_df["Child_Name"] == child_filter]

if field_filter != "All fields":
    filtered_df = filtered_df[filtered_df["Field"] == field_filter]

st.caption(f"Showing {len(filtered_df):,} of {total:,} issue(s)")

# ── Colour-coded severity column ──────────────────────────────────────────────

def colour_severity(val):
    if val == "ERROR":
        return "color: #dc2626; font-weight: 700;"
    elif val == "WARNING":
        return "color: #d97706; font-weight: 700;"
    return ""

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
        "Row":               st.column_config.NumberColumn("Row",     width="small"),
        "Child_Name":        st.column_config.TextColumn("Child",     width="medium"),
        "Field":             st.column_config.TextColumn("Field",     width="medium"),
        "Issue_Description": st.column_config.TextColumn("Issue",     width="large"),
        "Severity_Level":    st.column_config.TextColumn("Severity",  width="small"),
    },
)

# ─────────────────────────────────────────────────────────────────────────────
# BREAKDOWN BY FIELD
# ─────────────────────────────────────────────────────────────────────────────

with st.expander("📊 Issue breakdown by field"):
    breakdown = (
        df.groupby(["Field", "Severity_Level"])
        .size()
        .reset_index(name="Count")
        .sort_values(["Count"], ascending=False)
    )
    st.dataframe(breakdown, use_container_width=True, hide_index=True)
