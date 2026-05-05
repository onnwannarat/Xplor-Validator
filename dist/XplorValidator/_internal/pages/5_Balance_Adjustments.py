"""
Balance Adjustments — Page 5
=============================
Wraps scripts/process_balance_adjustments.py.  Reads balance adjustment data,
maps centre names to Xplor service names, and produces one styled XLSX per centre.
"""

import sys
from pathlib import Path

import streamlit as st

from shared.service_map import load_service_map_sidebar
from shared.styles import apply_styles

sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))
from process_balance_adjustments import main as run_balance_adjustments  # noqa: E402

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Balance Adjustments",
    page_icon="⚖️",
    layout="wide",
)

apply_styles()

# Bundled template path
_ASSETS_DIR = Path(__file__).parent.parent / "assets"
_BUNDLED_TEMPLATE = _ASSETS_DIR / "Balance Adjustments Details Upload Template.xlsx"

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
  <div class="step-text"><strong>Upload balance adjustment files</strong><br>
  Upload one or more CSV or XLSX files (including HTML-exported XLS).</div>
</div>

<div class="how-to-step">
  <div class="step-num">3</div>
  <div class="step-text"><strong>Select output folder</strong><br>
  Click Browse to choose where to save the per-centre XLSX files.</div>
</div>

<div class="how-to-step">
  <div class="step-num">4</div>
  <div class="step-text"><strong>Run</strong><br>
  Click <em>Run Balance Adjustments</em> and review the results.</div>
</div>
""", unsafe_allow_html=True)
    st.markdown("---")
    st.caption("Centre names must match Service Names in serviceIDs.csv exactly. Demo accounts are excluded automatically.")

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
    <h1>⚖️ Balance Adjustments</h1>
    <p>Maps centre names to Xplor service names, removes demo accounts, and produces one styled XLSX output file per centre using the upload template.</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# TEMPLATE
# ─────────────────────────────────────────────────────────────────────────────

template_bytes: bytes | None = None

if _BUNDLED_TEMPLATE.exists():
    template_bytes = _BUNDLED_TEMPLATE.read_bytes()
    st.success(f"✅ Using bundled template: `{_BUNDLED_TEMPLATE.name}`")
else:
    st.warning("Bundled template not found in `assets/`. Upload a template below.")

with st.expander("Advanced — Override template"):
    override_template = st.file_uploader(
        "Template XLSX (optional override)",
        type=["xlsx"],
        key="template_override",
        help="Overrides the bundled Balance Adjustments Details Upload Template.xlsx",
    )
    if override_template:
        template_bytes = override_template.read()
        st.success(f"✅ Using override template: `{override_template.name}`")

if not template_bytes:
    st.error("A template XLSX is required. Upload one in the Advanced section above.")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# FILE UPLOADERS
# ─────────────────────────────────────────────────────────────────────────────

input_files_uploaded = st.file_uploader(
    "**Balance Adjustment Files** — CSV or XLSX (required, upload multiple at once)",
    type=["csv", "xlsx", "xls"],
    accept_multiple_files=True,
    help="Columns required: Center Name (or Centre Name), Account Name. Credit/Owing or Amount Due columns are used for amounts.",
    key="balance_files_upload",
)

# ─────────────────────────────────────────────────────────────────────────────
# OUTPUT FOLDER
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("**Output folder** (required)")
col_path, col_browse = st.columns([5, 1])

with col_browse:
    st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
    if st.button("📂 Browse", key="browse_balance_out", use_container_width=True):
        picked = _pick_folder("Select output folder for balance adjustment files")
        if picked:
            st.session_state["balance_output_folder"] = picked

with col_path:
    out_folder_input = st.text_input(
        "Output folder",
        value=st.session_state.get("balance_output_folder", ""),
        placeholder="Click Browse or type a folder path…",
        label_visibility="collapsed",
        key="balance_output_folder_input",
    )
    if out_folder_input:
        st.session_state["balance_output_folder"] = out_folder_input

output_folder = st.session_state.get("balance_output_folder", "").strip()

# ─────────────────────────────────────────────────────────────────────────────
# RUN BUTTON
# ─────────────────────────────────────────────────────────────────────────────

all_ready = bool(input_files_uploaded) and bool(output_folder)
run_btn = st.button(
    "▶ Run Balance Adjustments",
    type="primary",
    disabled=not all_ready,
    help="Upload files and select an output folder to enable." if not all_ready else "",
)

if not run_btn:
    if not input_files_uploaded:
        st.info("Upload balance adjustment files to continue.")
    elif not output_folder:
        st.info("Select an output folder to continue.")
    st.stop()

if not Path(output_folder).is_dir():
    st.error(f"Output folder not found: `{output_folder}`")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# EXECUTE
# ─────────────────────────────────────────────────────────────────────────────

with st.spinner("Processing balance adjustments — please wait…"):
    try:
        input_files = [(f.name, f.read()) for f in input_files_uploaded]
        result = run_balance_adjustments(
            input_files=input_files,
            service_ids_bytes=svc_map_bytes,
            template_bytes=template_bytes,
            output_dir=output_folder,
        )
    except Exception as exc:
        st.error(f"Error processing balance adjustments: {exc}")
        st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# RESULTS
# ─────────────────────────────────────────────────────────────────────────────

st.success(f"✅ Processing complete. Files saved to: `{output_folder}`")

m1, m2 = st.columns(2)
with m1:
    st.metric("Output Files Created", result["total_outputs"])
with m2:
    st.metric("Total Data Rows Written", result["total_rows"])

# Created files
if result["created_files"]:
    st.markdown("### Created Files")
    for cf in sorted(result["created_files"], key=lambda x: x["centre"]):
        st.write(f"✅ **{cf['centre']}** — `{Path(cf['path']).name}` ({cf['rows']} rows)")

# Skipped centres
if result["skipped_centres"]:
    st.markdown(
        f'<div class="status-banner status-warning">⚠️ {len(result["skipped_centres"])} centre(s) skipped — name not found in serviceIDs.csv.</div>',
        unsafe_allow_html=True,
    )
    with st.expander(f"Skipped Centres ({len(result['skipped_centres'])})"):
        for c in sorted(result["skipped_centres"]):
            st.write(f"- `{c}`")
    st.caption("Tip: Centre names must match the **Service Name** column in serviceIDs.csv exactly (case-sensitive).")

# File read errors
if result["errors"]:
    with st.expander(f"⚠️ File Read Errors ({len(result['errors'])})"):
        for err in result["errors"]:
            st.error(err)
