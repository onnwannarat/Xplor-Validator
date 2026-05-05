"""
Bookings Import — Page 3
========================
Wraps scripts/prepare_bookings_import.py.  Processes QikKids recurring and
casual booking exports and produces split CSV files per service.
"""

import os
import sys
from pathlib import Path

import streamlit as st

from shared.service_map import load_service_map_sidebar
from shared.styles import apply_styles

sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))
from prepare_bookings_import import main as run_bookings_import  # noqa: E402

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Bookings Import",
    page_icon="📅",
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
  <div class="step-text"><strong>Upload serviceIDs.csv</strong><br>
  Already done in the Service ID Mapping section above.</div>
</div>

<div class="how-to-step">
  <div class="step-num">2</div>
  <div class="step-text"><strong>Upload booking files</strong><br>
  Upload one or more QikKids exports (recurring and casual files can be mixed together).</div>
</div>

<div class="how-to-step">
  <div class="step-num">3</div>
  <div class="step-text"><strong>Select output folder</strong><br>
  Click Browse to choose where to save the split CSVs and reports.</div>
</div>

<div class="how-to-step">
  <div class="step-num">4</div>
  <div class="step-text"><strong>Run</strong><br>
  Click <em>Run Bookings Import</em> and review the summary.</div>
</div>

<div class="how-to-step">
  <div class="step-num">5</div>
  <div class="step-text"><strong>Upload order</strong><br>
  Always upload <strong>Recurring/</strong> files first, then <strong>Casual/</strong> files.</div>
</div>
""", unsafe_allow_html=True)
    st.markdown("---")
    st.caption("Exact duplicates and overlapping recurring schedules are removed automatically.")

# ─────────────────────────────────────────────────────────────────────────────
# GUARD
# ─────────────────────────────────────────────────────────────────────────────

if not svc_map_bytes:
    st.info("👆 Upload a service ID mapping file in the sidebar to begin.")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("""
<div class="tool-header">
    <h1>📅 Bookings Import Preparer</h1>
    <p>Processes QikKids booking exports: removes duplicates and schedule overlaps, maps service IDs, and saves split CSV files per service ready for Xplor import.</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# FILE UPLOADERS
# ─────────────────────────────────────────────────────────────────────────────

bookings_files = st.file_uploader(
    "**QikKids Booking Files** — CSV or XLSX (required, upload multiple at once)",
    type=["csv", "xlsx"],
    accept_multiple_files=True,
    help="Recurring and casual files can be mixed. Required columns include Service Legacy ID, Child Legacy ID, Fee Name, Room Name, Frequency.",
    key="bookings_files_upload",
)

# ─────────────────────────────────────────────────────────────────────────────
# OUTPUT FOLDER
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("**Output folder** (required)")
col_path, col_browse = st.columns([5, 1])

with col_browse:
    st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
    if st.button("📂 Browse", key="browse_bookings_out", use_container_width=True):
        picked = _pick_folder("Select output folder for booking files")
        if picked:
            st.session_state["bookings_output_folder"] = picked

with col_path:
    out_folder_input = st.text_input(
        "Output folder",
        value=st.session_state.get("bookings_output_folder", ""),
        placeholder="Click Browse or type a folder path…",
        label_visibility="collapsed",
        key="bookings_output_folder_input",
    )
    if out_folder_input:
        st.session_state["bookings_output_folder"] = out_folder_input

output_folder = st.session_state.get("bookings_output_folder", "").strip()

# ─────────────────────────────────────────────────────────────────────────────
# RUN BUTTON
# ─────────────────────────────────────────────────────────────────────────────

all_ready = bool(bookings_files) and bool(output_folder)
run_btn = st.button(
    "▶ Run Bookings Import",
    type="primary",
    disabled=not all_ready,
    help="Upload files and select output folder to enable." if not all_ready else "",
)

if not run_btn:
    if not bookings_files:
        st.info("Upload QikKids booking files to continue.")
    elif not output_folder:
        st.info("Select an output folder to continue.")
    st.stop()

if not Path(output_folder).is_dir():
    st.error(f"Output folder not found: `{output_folder}`")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# EXECUTE
# ─────────────────────────────────────────────────────────────────────────────

with st.spinner("Processing bookings — please wait…"):
    try:
        input_files = [(f.name, f.read()) for f in bookings_files]
        result = run_bookings_import(
            input_files=input_files,
            service_ids_bytes=svc_map_bytes,
            output_dir=output_folder,
        )
    except Exception as exc:
        st.error(f"Error processing bookings: {exc}")
        st.stop()

if "error" in result:
    st.error(result["error"])
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# RESULTS
# ─────────────────────────────────────────────────────────────────────────────

st.success(f"✅ Processing complete. Files saved to: `{output_folder}`")

m1, m2, m3, m4, m5 = st.columns(5)
with m1:
    st.metric("Input Files", result["n_input_files"])
with m2:
    st.metric("Raw Rows", result["n_raw_rows"])
with m3:
    st.metric("Recurring Bookings", result["n_recurring"])
with m4:
    st.metric("Casual Bookings", result["n_casual"])
with m5:
    st.metric("Duplicates Removed", result["n_dupe_rows"])

if result["n_sched_conflict_rows"]:
    st.warning(f"⚠️ {result['n_sched_conflict_rows']} recurring booking(s) removed due to schedule conflicts across {result['n_sched_conflict_groups']} group(s).")

if result["n_casual_removed"]:
    st.warning(f"⚠️ {result['n_casual_removed']} casual booking(s) removed due to overlap with recurring bookings.")

# Unmapped IDs
if result["unmapped_ids"]:
    st.markdown(
        '<div class="status-banner status-error">❌ Some QK Service IDs could not be mapped — check the list below.</div>',
        unsafe_allow_html=True,
    )
    with st.expander(f"⚠️ Unmapped Service IDs ({len(result['unmapped_ids'])})"):
        for uid in sorted(result["unmapped_ids"]):
            st.write(f"- `{uid}`")

# Output file lists
col_r, col_c = st.columns(2)

with col_r:
    st.markdown(f"**Recurring Files** ({len(result['recurring_files'])})")
    if result["recurring_files"]:
        for fname, n_rows in sorted(result["recurring_files"]):
            st.write(f"✅ `{fname}` — {n_rows} rows")
    else:
        st.caption("No recurring files produced.")

with col_c:
    st.markdown(f"**Casual Files** ({len(result['casual_files'])})")
    if result["casual_files"]:
        for fname, n_rows in sorted(result["casual_files"]):
            st.write(f"✅ `{fname}` — {n_rows} rows")
    else:
        st.caption("No casual files produced.")
