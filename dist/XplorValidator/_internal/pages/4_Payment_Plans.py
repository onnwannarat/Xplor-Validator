"""
Payment Plans — Page 4
======================
Wraps scripts/payment_plan_checker.py.  Validates payment plan CSVs against
Onboarding Tool error rules, auto-fixes formats, and splits output by service.
"""

import sys
from pathlib import Path

import streamlit as st

from shared.service_map import load_service_map_sidebar
from shared.styles import apply_styles

sys.path.insert(0, str(Path(__file__).parent.parent / "scripts"))
from payment_plan_checker import (  # noqa: E402
    COLUMN_LABELS,
    DEFAULT_COLUMNS,
    run_payment_plan_checker,
)

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────

st.set_page_config(
    page_title="Payment Plans",
    page_icon="💳",
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
  <div class="step-text"><strong>Upload payment plan CSV</strong><br>
  Upload the payment plan CSV exported from QikKids.</div>
</div>

<div class="how-to-step">
  <div class="step-num">3</div>
  <div class="step-text"><strong>(Optional) Adjust column mapping</strong><br>
  If your CSV uses different column names, expand the Column Mapping section and update them.</div>
</div>

<div class="how-to-step">
  <div class="step-num">4</div>
  <div class="step-text"><strong>Select output folder</strong><br>
  Click Browse to choose where to save the cleaned CSV, error report, and split files.</div>
</div>

<div class="how-to-step">
  <div class="step-num">5</div>
  <div class="step-text"><strong>Run and review</strong><br>
  Click <em>Run Checker</em> and fix any errors shown before importing.</div>
</div>
""", unsafe_allow_html=True)
    st.markdown("---")
    st.caption("Date formats and weekday abbreviations are fixed automatically.")

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
    <h1>💳 Payment Plan Checker</h1>
    <p>Validates payment plan CSVs against Onboarding Tool error rules, auto-fixes date formats and weekday abbreviations, and splits output by service.</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# FILE UPLOADER
# ─────────────────────────────────────────────────────────────────────────────

plan_file = st.file_uploader(
    "**Payment Plan CSV** (required)",
    type=["csv"],
    help="The raw payment plan export from QikKids.",
    key="plan_file_upload",
)

# ─────────────────────────────────────────────────────────────────────────────
# COLUMN MAPPING
# ─────────────────────────────────────────────────────────────────────────────

with st.expander("⚙️ Column Mapping (expand to customise if your CSV uses different column names)"):
    col_mapping: dict[str, str] = {}
    keys = list(COLUMN_LABELS.keys())
    n_cols = 3
    rows_needed = -(-len(keys) // n_cols)

    for row_i in range(rows_needed):
        grid_cols = st.columns(n_cols)
        for col_i, c in enumerate(grid_cols):
            idx = row_i * n_cols + col_i
            if idx >= len(keys):
                break
            key = keys[idx]
            label = COLUMN_LABELS[key]
            default = DEFAULT_COLUMNS.get(key, "")
            with c:
                val = st.text_input(
                    label,
                    value=default,
                    key=f"col_map_{key}",
                )
                col_mapping[key] = val

# ─────────────────────────────────────────────────────────────────────────────
# OUTPUT FOLDER
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("**Output folder** (required)")
col_path, col_browse = st.columns([5, 1])

with col_browse:
    st.markdown("<div style='height:28px'></div>", unsafe_allow_html=True)
    if st.button("📂 Browse", key="browse_plans_out", use_container_width=True):
        picked = _pick_folder("Select output folder for payment plan files")
        if picked:
            st.session_state["plans_output_folder"] = picked

with col_path:
    out_folder_input = st.text_input(
        "Output folder",
        value=st.session_state.get("plans_output_folder", ""),
        placeholder="Click Browse or type a folder path…",
        label_visibility="collapsed",
        key="plans_output_folder_input",
    )
    if out_folder_input:
        st.session_state["plans_output_folder"] = out_folder_input

output_folder = st.session_state.get("plans_output_folder", "").strip()

# ─────────────────────────────────────────────────────────────────────────────
# RUN BUTTON
# ─────────────────────────────────────────────────────────────────────────────

all_ready = bool(plan_file) and bool(output_folder)
run_btn = st.button(
    "▶ Run Checker",
    type="primary",
    disabled=not all_ready,
    help="Upload a file and select an output folder to enable." if not all_ready else "",
)

if not run_btn:
    if not plan_file:
        st.info("Upload a payment plan CSV to continue.")
    elif not output_folder:
        st.info("Select an output folder to continue.")
    st.stop()

if not Path(output_folder).is_dir():
    st.error(f"Output folder not found: `{output_folder}`")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
# EXECUTE
# ─────────────────────────────────────────────────────────────────────────────

with st.spinner("Validating payment plans — please wait…"):
    try:
        result_dict = run_payment_plan_checker(
            input_bytes=plan_file.read(),
            filename=plan_file.name,
            service_ids_bytes=svc_map_bytes,
            output_dir=output_folder,
            col_mapping=col_mapping or None,
        )
    except Exception as exc:
        st.error(f"Error processing payment plans: {exc}")
        st.stop()

stats  = result_dict["result"]["stats"]
errors = result_dict["result"]["errors"]
total_err = sum(len(v) for v in errors.values())

# ─────────────────────────────────────────────────────────────────────────────
# METRICS
# ─────────────────────────────────────────────────────────────────────────────

st.success(f"✅ Processing complete. Files saved to: `{output_folder}`")

m1, m2, m3, m4, m5 = st.columns(5)
with m1:
    st.metric("Total Rows", stats["total"])
with m2:
    st.metric("Date Format Fixed", stats["date_fixed"])
with m3:
    st.metric("Weekday Abbreviated", stats["weekday_fixed"])
with m4:
    st.metric("Trailing Spaces Removed", stats["spaces_fixed"])
with m5:
    st.metric("Total Errors", total_err)

# Status banner
if total_err == 0:
    st.markdown(
        '<div class="status-banner status-ok">✅ No errors found — this file appears ready for import.</div>',
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        f'<div class="status-banner status-error">❌ {total_err} error(s) found — review and fix before importing.</div>',
        unsafe_allow_html=True,
    )

# ─────────────────────────────────────────────────────────────────────────────
# ERROR SUMMARY TABLE
# ─────────────────────────────────────────────────────────────────────────────

PRIORITY_DISPLAY = [
    ("weekend",            "WEEKEND — Sat / Sun",                    "🔴"),
    ("missing_date",       "ERROR_MISSING_BOOKING_START_DATE",       "🔴"),
    ("missing_weekday",    "ERROR_MISSING_PAYMENT_DAY",              "🔴"),
    ("missing_parent",     "ERROR_MISSING_GUARDIAN",                 "🔴"),
    ("missing_service_id", "ERROR_MISSING_SERVICE_ID",               "🔴"),
    ("invalid_cycle",      "ERROR_INVALID_FREQUENCY",                "🔴"),
    ("manual_not_monday",  "Manual Plan — Not Monday",               "🟠"),
    ("negative_limit",     "ERROR_NEGATIVE_DIRECT_DEBIT_LIMIT",      "🔴"),
    ("negative_fixed",     "ERROR_NEGATIVE_FIXED_LIMIT",             "🔴"),
    ("both_amounts",       "ERROR_ONLY_ONE_AMOUNT_ALLOWED",          "🔴"),
    ("unparseable_date",   "Unparseable Date",                       "🟡"),
    ("unknown_weekday",    "Unknown Weekday Value",                   "🟡"),
]

summary_rows = []
for key, label, icon in PRIORITY_DISPLAY:
    n = len(errors.get(key, []))
    summary_rows.append({"": icon, "Error Category": label, "Count": n})

import pandas as pd
summary_df = pd.DataFrame(summary_rows)

def _colour_count(val):
    if val > 0:
        return "color: #dc2626; font-weight: 700;"
    return "color: #16a34a;"

styled_summary = summary_df.style.map(_colour_count, subset=["Count"])
st.dataframe(styled_summary, use_container_width=True, hide_index=True)

# ─────────────────────────────────────────────────────────────────────────────
# OUTPUT FILES
# ─────────────────────────────────────────────────────────────────────────────

st.markdown("### Output Files")
st.write(f"📄 Cleaned CSV: `{result_dict['cleaned_path']}`")
st.write(f"📊 Error Report: `{result_dict['error_path']}`")

split = result_dict.get("split_result", {})
known = split.get("known", {})
unknown = split.get("unknown", {})

if known:
    with st.expander(f"Split files by service ({len(known)} service(s))"):
        for svc_name, path in sorted(known.items()):
            st.write(f"✅ **{svc_name}** — `{path}`")

if unknown:
    with st.expander(f"⚠️ Unmapped Service IDs ({len(unknown)} group(s))", expanded=True):
        for sid, path in sorted(unknown.items()):
            st.write(f"⚠️ Service ID `{sid}` — `{path}`")
