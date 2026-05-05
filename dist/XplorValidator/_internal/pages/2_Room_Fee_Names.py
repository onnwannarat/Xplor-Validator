"""
Room & Fee Names — Page 2
=========================
Wraps scripts/check_names.py.  Compares fee and room names in QikKids booking
exports against those defined in Xplor and produces an Excel mismatch report.
"""

import sys
from pathlib import Path

import streamlit as st

from shared.service_map import load_service_map_sidebar
from shared.styles import apply_styles

sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))
from check_names import run_check_names  # noqa: E402

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Room & Fee Names",
    page_icon="🏷️",
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
  <div class="step-text"><strong>Upload Xplor Active Fees CSV</strong><br>
  Export from Xplor: active fee names per service.</div>
</div>

<div class="how-to-step">
  <div class="step-num">3</div>
  <div class="step-text"><strong>Upload Xplor Active Rooms CSV</strong><br>
  Export from Xplor: active room names per centre.</div>
</div>

<div class="how-to-step">
  <div class="step-num">4</div>
  <div class="step-text"><strong>Upload QikKids Bookings</strong><br>
  Upload one or more QikKids booking exports (recurring and/or casual).</div>
</div>

<div class="how-to-step">
  <div class="step-num">5</div>
  <div class="step-text"><strong>Run and download report</strong><br>
  Click <em>Run Check</em> and download the Excel mismatch report.</div>
</div>
""", unsafe_allow_html=True)
    st.markdown("---")
    st.caption("Only fee/room names that do NOT match are reported.")

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
    <h1>🏷️ Room &amp; Fee Names Checker</h1>
    <p>Compares fee and room names in QikKids bookings against names defined in Xplor. Produces an Excel report with mismatch rows and fuzzy-match suggestions.</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# FILE UPLOADERS
# ─────────────────────────────────────────────────────────────────────────────

col_a, col_b = st.columns(2)

with col_a:
    fees_file = st.file_uploader(
        "**Xplor Active Fees CSV** (required)",
        type=["csv"],
        help="Columns required: Service ID, Fee Name",
        key="fees_upload",
    )
    rooms_file = st.file_uploader(
        "**Xplor Active Rooms CSV** (required)",
        type=["csv"],
        help="Columns required: Centre_Name, Room_Name",
        key="rooms_upload",
    )

with col_b:
    bookings_files = st.file_uploader(
        "**QikKids Bookings CSV** (required — upload one or more)",
        type=["csv", "xlsx"],
        accept_multiple_files=True,
        help="Recurring and/or casual bookings exports. Columns required: Service Legacy ID, Fee Name, Room Name",
        key="bookings_upload",
    )

# ─────────────────────────────────────────────────────────────────────────────
# SAVE TO FOLDER (optional)
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("**Save report to folder** (optional)")
col_path, col_browse = st.columns([5, 1])

with col_browse:
    st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
    if st.button("📂 Browse", key="browse_names_folder", use_container_width=True):
        picked = _pick_folder("Select folder to save report")
        if picked:
            st.session_state["names_save_folder"] = picked

with col_path:
    save_folder_input = st.text_input(
        "Save folder",
        value=st.session_state.get("names_save_folder", ""),
        placeholder="Leave blank to download only, or click Browse…",
        label_visibility="collapsed",
        key="names_save_folder_input",
    )
    if save_folder_input:
        st.session_state["names_save_folder"] = save_folder_input

save_folder = st.session_state.get("names_save_folder", "").strip()

# ─────────────────────────────────────────────────────────────────────────────
# RUN BUTTON
# ─────────────────────────────────────────────────────────────────────────────

all_ready = fees_file and rooms_file and bookings_files
run_btn = st.button(
    "▶ Run Check",
    type="primary",
    disabled=not all_ready,
    help="Upload all required files above to enable." if not all_ready else "",
)

if not run_btn:
    if not all_ready:
        st.info("Upload the required files above to enable the run button.")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# EXECUTE
# ─────────────────────────────────────────────────────────────────────────────

with st.spinner("Comparing names — please wait…"):
    try:
        bookings_bytes_list = [f.read() for f in bookings_files]
        excel_bytes, n_fees, n_rooms = run_check_names(
            service_ids_bytes=svc_map_bytes,
            fees_bytes=fees_file.read(),
            rooms_bytes=rooms_file.read(),
            bookings_bytes_list=bookings_bytes_list,
        )
    except Exception as exc:
        st.error(f"Error running check: {exc}")
        st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# RESULTS
# ─────────────────────────────────────────────────────────────────────────────

m1, m2 = st.columns(2)
with m1:
    st.metric("Fee Mismatches", n_fees)
with m2:
    st.metric("Room Mismatches", n_rooms)

if n_fees == 0 and n_rooms == 0:
    st.markdown(
        '<div class="status-banner status-ok">✅ No mismatches found — all fee and room names match Xplor.</div>',
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        f'<div class="status-banner status-warning">⚠️ {n_fees} fee mismatch(es) and {n_rooms} room mismatch(es) found. Download the report for details.</div>',
        unsafe_allow_html=True,
    )

st.download_button(
    label="⬇️ Download Mismatch Report (Excel)",
    data=excel_bytes,
    file_name="Fee_Room_Name_Mismatch_Report.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

if save_folder:
    from pathlib import Path as _Path
    if _Path(save_folder).is_dir():
        out_path = _Path(save_folder) / "Fee_Room_Name_Mismatch_Report.xlsx"
        out_path.write_bytes(excel_bytes)
        st.success(f"✅ Report also saved to: `{out_path}`")
    else:
        st.warning(f"Save folder not found: `{save_folder}` — use the download button instead.")
