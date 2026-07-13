"""
Xplor Validator — Page 1
========================
Migrated from app.py with minimal changes. The sidebar service map section is
now handled by shared.service_map.load_service_map_sidebar().
"""

import hashlib
import os
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from shared.service_map import load_service_map_sidebar
from shared.styles import apply_styles
from validator_v2 import (
    IssueRecorder,
    ServiceMapping,
    load_existing_parent_profiles_from_bytes,
    run_v2_from_bytes,
    write_client_excel_report,
    write_duplicate_parents_report,
    write_excel_report,
    write_split_csvs,
)

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Xplor Validator",
    page_icon="🔍",
    layout="wide",
)

apply_styles()

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def _pick_folder(title: str = "Select destination folder") -> str:
    if sys.platform == "darwin":
        try:
            import subprocess
            script = f'POSIX path of (choose folder with prompt "{title}")'
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True, text=True, timeout=120,
            )
            return result.stdout.strip().rstrip("/") if result.returncode == 0 else ""
        except Exception:
            return ""
    else:
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            root.wm_attributes("-topmost", True)
            folder = filedialog.askdirectory(title=title)
            root.destroy()
            return folder or ""
        except Exception:
            return ""


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


# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────

svc_map_bytes = load_service_map_sidebar()

with st.sidebar:
    st.markdown("## 📖 How to Use")
    st.markdown("---")

    st.markdown("""
<div class="how-to-step">
  <div class="step-num">1</div>
  <div class="step-text"><strong>Upload the service ID mapping file</strong><br>
  Upload the service ID mapping CSV for this project above. Any filename is accepted.</div>
</div>

<div class="how-to-step">
  <div class="step-num">2</div>
  <div class="step-text"><strong>Upload your migration file</strong><br>
  Drag and drop a <strong>CSV</strong> or <strong>XLSX</strong> onto the upload box.</div>
</div>

<div class="how-to-step">
  <div class="step-num">3</div>
  <div class="step-text"><strong>Select destination folder</strong><br>
  Click <em>📂 Browse</em> to pick where output files will be saved. An <strong>output/</strong> subfolder will be created inside it.</div>
</div>

<div class="how-to-step">
  <div class="step-num">4</div>
  <div class="step-text"><strong>Select import scope</strong><br>
  Choose <strong>Active only</strong> to include only Active children, or <strong>Active + WaitList</strong> to include both.</div>
</div>

<div class="how-to-step">
  <div class="step-num">5</div>
  <div class="step-text"><strong>(Optional) Add existing service files</strong><br>
  Upload or point to CSV/XLSX files from <em>other</em> services already in Xplor. The tool will flag duplicate parents.</div>
</div>

<div class="how-to-step">
  <div class="step-num">6</div>
  <div class="step-text"><strong>Wait for validation</strong><br>
  The tool checks all rows, maps service IDs, normalises states, deduplicates emails, and runs the cross-service parent check.</div>
</div>

<div class="how-to-step">
  <div class="step-num">7</div>
  <div class="step-text"><strong>Collect your outputs</strong><br>
  All output files are saved automatically to the <strong>output/</strong> folder inside your chosen destination.</div>
</div>

<div class="how-to-step">
  <div class="step-num">8</div>
  <div class="step-text"><strong>Fix & re-upload</strong><br>
  Correct any errors, re-export, and upload again to re-validate.</div>
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
    st.markdown("### 🔁 Cross-service duplicate check")
    st.markdown("""
When existing service files are provided (Step 5), the tool checks every parent in the migration file against those profiles. A duplicate is flagged when:
- **First name + last name** match, AND
- At least one of **DOB**, **contact number**, or **email** also matches.
""")

    st.markdown("---")
    st.markdown("### 📋 Validation checklist")
    with st.expander("View full checklist"):
        st.markdown("""
**Mandatory Fields**
- Child: Service ID, Legacy ID, First/Last Name, DOB, Status, CRN, Room, Enrolment Date
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
""")

    st.markdown("---")
    st.caption("Xplor Data Migration Validator · Does not modify your original file.")

# ─────────────────────────────────────────────────────────────────────────────
# GUARD — service map required
# ─────────────────────────────────────────────────────────────────────────────

if not svc_map_bytes:
    st.info("👆 Upload a service ID mapping file in the sidebar to begin.")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="tool-header">
    <h1>🔍 Xplor Data Migration Validator</h1>
    <p>Upload your migration file and choose a destination folder — all output files will be saved there automatically.</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# STEP 1 — FILE UPLOAD
# ─────────────────────────────────────────────────────────────────────────────

uploaded_file = st.file_uploader(
    "**Step 1** — Upload your migration file (CSV or XLSX)",
    type=["csv", "xlsx", "xls"],
    help="Your original file is not modified.",
)

# ─────────────────────────────────────────────────────────────────────────────
# STEP 2 — DESTINATION FOLDER
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("**Step 2** — Select destination folder")

col_input, col_btn = st.columns([5, 1])

with col_btn:
    st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
    if st.button("📂 Browse", use_container_width=True):
        picked = _pick_folder()
        if picked:
            st.session_state["dest_folder"] = picked

with col_input:
    dest_folder_input = st.text_input(
        label="Destination folder",
        value=st.session_state.get("dest_folder", ""),
        placeholder="Click Browse or type a folder path…",
        label_visibility="collapsed",
    )
    if dest_folder_input:
        st.session_state["dest_folder"] = dest_folder_input

dest_folder = st.session_state.get("dest_folder", "").strip()

# ─────────────────────────────────────────────────────────────────────────────
# GUARD — both inputs required
# ─────────────────────────────────────────────────────────────────────────────

if not uploaded_file and not dest_folder:
    st.info("👆 Upload a file and select a destination folder to begin.")
    st.stop()

if not uploaded_file:
    st.info("👆 Upload a migration file to continue.")
    st.stop()

if not dest_folder:
    st.info("📂 Select a destination folder to continue.")
    st.stop()

if not Path(dest_folder).is_dir():
    st.error(f"Folder not found: `{dest_folder}`")
    st.stop()

dest_folder = str(Path(dest_folder).resolve())

st.divider()

# ─────────────────────────────────────────────────────────────────────────────
# STEP 3 — IMPORT SCOPE
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("**Step 3** — Select import scope")
import_mode = st.radio(
    "Which children should be included?",
    options=["Active only", "Active + WaitList"],
    horizontal=True,
    key="import_mode",
    help=(
        "**Active only** — processes children whose Status is 'Active'. "
        "Waitlist rows are excluded from validation and all output files.\n\n"
        "**Active + WaitList** — processes both Active and Waitlist children."
    ),
)
include_waitlist = import_mode == "Active + WaitList"

# ─────────────────────────────────────────────────────────────────────────────
# STEP 4 — EXISTING SERVICE FILES (OPTIONAL)
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("**Step 4 (Optional)** — Cross-check against existing service data")
st.caption(
    "Provide CSV/XLSX files from other services already in Xplor. "
    "Any parent whose name + at least one of DOB / contact / email matches will be flagged as a duplicate."
)

existing_source = st.radio(
    "existing_source_radio",
    options=["Upload files", "Select from folder"],
    horizontal=True,
    label_visibility="collapsed",
    key="existing_source",
)

existing_files_data: list[tuple[bytes, str]] = []

if existing_source == "Upload files":
    uploaded_existing = st.file_uploader(
        "Existing service files (CSV or XLSX, multiple allowed)",
        type=["csv", "xlsx", "xls"],
        accept_multiple_files=True,
        key="existing_files_upload",
        label_visibility="collapsed",
    )
    if uploaded_existing:
        existing_files_data = [(f.read(), f.name) for f in uploaded_existing]
        st.success(f"✅ {len(existing_files_data)} existing file(s) ready for cross-check.")
    else:
        st.info("No files uploaded — cross-service duplicate check will be skipped.")

else:
    col_ex, col_ex_btn = st.columns([5, 1])
    with col_ex_btn:
        st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
        if st.button("📂 Browse", key="browse_existing", use_container_width=True):
            picked = _pick_folder(title="Select folder containing existing service files")
            if picked:
                st.session_state["existing_folder"] = picked
    with col_ex:
        existing_folder_input = st.text_input(
            "Existing service files folder",
            value=st.session_state.get("existing_folder", ""),
            placeholder="Click Browse or type a folder path…",
            key="existing_folder_text",
            label_visibility="collapsed",
        )
        if existing_folder_input:
            st.session_state["existing_folder"] = existing_folder_input

    existing_folder = st.session_state.get("existing_folder", "").strip()
    if existing_folder:
        if not Path(existing_folder).is_dir():
            st.error(f"Folder not found: `{existing_folder}`")
        else:
            found_paths = sorted(
                f for f in Path(existing_folder).iterdir()
                if f.suffix.lower() in (".csv", ".xlsx", ".xls")
                and f.name.lower() != "serviceids.csv"
            )
            if found_paths:
                for fp in found_paths:
                    try:
                        existing_files_data.append((fp.read_bytes(), fp.name))
                    except Exception as exc:
                        st.warning(f"Could not read {fp.name}: {exc}")
                names = ", ".join(f.name for f in found_paths)
                st.success(f"✅ {len(existing_files_data)} file(s) found: {names}")
            else:
                st.info("No CSV/XLSX files found in the selected folder — cross-service check will be skipped.")
    else:
        st.info("No folder selected — cross-service duplicate check will be skipped.")

st.divider()

# ─────────────────────────────────────────────────────────────────────────────
# RUN VALIDATION
# ─────────────────────────────────────────────────────────────────────────────

svc_hash     = hashlib.md5(svc_map_bytes).hexdigest()
file_bytes   = uploaded_file.read()
file_hash    = hashlib.md5(file_bytes).hexdigest()
existing_hash = (
    hashlib.md5(b"".join(d for d, _ in existing_files_data)).hexdigest()
    if existing_files_data else "none"
)
run_key = file_hash + dest_folder + svc_hash + existing_hash + import_mode

if st.session_state.get("confirmed_key") != run_key:
    if st.button("🚀 Let's do this!", type="primary", use_container_width=False):
        st.session_state["confirmed_key"] = run_key
        st.rerun()
    st.stop()

if st.session_state.get("last_run_key") != run_key:
    with st.spinner("Validating and transforming your file — please wait…"):
        service_map = _load_service_map_from_bytes(svc_map_bytes)

        try:
            recorder, all_rows, fieldnames = run_v2_from_bytes(
                file_bytes, uploaded_file.name, service_map,
                existing_files=existing_files_data or None,
                include_waitlist=include_waitlist,
            )
        except Exception as exc:
            st.error(f"Could not process the file: {exc}")
            st.stop()

        output_dir = Path(dest_folder) / "output"
        try:
            os.makedirs(str(output_dir), exist_ok=True)
        except PermissionError:
            import sys
            if sys.platform == "darwin":
                st.error(
                    f"Cannot create output folder in:\n\n`{dest_folder}`\n\n"
                    "macOS is blocking access to this folder.\n\n"
                    "**Fix:** Go to **System Settings → Privacy & Security → Files and Folders** "
                    "(or **Full Disk Access**) and allow this app to access the folder, "
                    "then try again.\n\n"
                    "Alternatively, choose a folder on the Desktop."
                )
            else:
                st.error(
                    f"Cannot create output folder in:\n\n`{dest_folder}`\n\n"
                    "**Permission denied.** You may not have write access to this folder.\n\n"
                    "Try choosing a different destination, such as the Desktop."
                )
            st.stop()
        except FileNotFoundError:
            st.error(
                f"Cannot create output folder in:\n\n`{dest_folder}`\n\n"
                "This folder appears to be a **OneDrive cloud-only placeholder** "
                "that hasn't been synced to this device yet.\n\n"
                "**Fix:** Right-click the folder in File Explorer → "
                "**Always keep on this device** → wait for sync to finish, then try again.\n\n"
                "Alternatively, choose a folder outside OneDrive (e.g. Desktop or C:\\Temp)."
            )
            st.stop()

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        write_key = f"files_written_{run_key}"
        if write_key not in st.session_state:
            try:
                _, row_num_map = write_split_csvs(all_rows, fieldnames, str(output_dir))
                excel_path = output_dir / f"validation_audit_report_{ts}.xlsx"
                write_excel_report(recorder, all_rows, str(excel_path), service_map, row_num_map)
                client_path = output_dir / f"client_audit_report_{ts}.xlsx"
                write_client_excel_report(recorder, all_rows, str(client_path), service_map, row_num_map)
                dup_parents_path = output_dir / f"duplicate_parents_report_{ts}.xlsx"
                dup_cross_count, dup_intra_count = write_duplicate_parents_report(
                    recorder, str(dup_parents_path), service_map, all_rows, row_num_map
                )
                output_saved   = True
                output_dir_str = str(output_dir)
            except Exception as exc:
                st.warning(f"Validation complete, but could not write output files: {exc}")
                output_saved   = False
                output_dir_str = ""
                dup_cross_count = 0
                dup_intra_count = 0

            st.session_state[write_key] = True
        else:
            output_saved    = st.session_state.get("output_saved", False)
            output_dir_str  = st.session_state.get("output_dir", "")
            dup_cross_count = st.session_state.get("dup_cross_count", 0)
            dup_intra_count = st.session_state.get("dup_intra_count", 0)

        st.session_state.update({
            "last_run_key":    run_key,
            "recorder":        recorder,
            "output_saved":    output_saved,
            "output_dir":      output_dir_str,
            "timestamp":       ts,
            "dup_cross_count": dup_cross_count,
            "dup_intra_count": dup_intra_count,
        })

recorder        = st.session_state["recorder"]
output_saved    = st.session_state["output_saved"]
output_dir      = st.session_state["output_dir"]
dup_cross_count = st.session_state.get("dup_cross_count", 0)
dup_intra_count = st.session_state.get("dup_intra_count", 0)

errors   = recorder.error_count()
warnings = recorder.warning_count()
fixed    = recorder.fixed_count()
total    = len(recorder.issues)
df       = recorder.to_dataframe()

# ─────────────────────────────────────────────────────────────────────────────
# OUTPUT FOLDER NOTICE
# ─────────────────────────────────────────────────────────────────────────────

if output_saved:
    st.success(f"✅ Output files saved to: `{output_dir}`")

# ─────────────────────────────────────────────────────────────────────────────
# SUMMARY METRICS
# ─────────────────────────────────────────────────────────────────────────────

col1, col2, col3, col4, col5, col6 = st.columns(6)
with col1:
    st.metric("Total Entries", total)
with col2:
    st.metric("🔴 Errors", errors, help="Must be resolved before importing.")
with col3:
    st.metric("🟡 Warnings", warnings, help="Review before importing.")
with col4:
    st.metric("🟢 Auto-Fixed", fixed, help="Automatically corrected by the tool.")
with col5:
    st.metric("🔁 Within-Upload Dupes", dup_intra_count)
with col6:
    st.metric("🔁 Cross-Service Dupes", dup_cross_count)

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
