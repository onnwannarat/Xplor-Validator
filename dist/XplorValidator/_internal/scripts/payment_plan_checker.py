"""
Payment Plan Import Checker
============================
Validates and prepares data for Payment Plan Import into the Onboarding Tool.

Validation rules aligned with the Onboarding Tool Error List:
  ERROR_MISSING_GUARDIAN           → No parent name provided
  ERROR_MISSING_PAYMENT_DAY        → Weekday is missing
  ERROR_INVALID_PAYMENT_DAY        → Weekday is Saturday or Sunday
  ERROR_INVALID_FREQUENCY          → Billing Cycle is not Weekly / Fortnightly / Monthly
  ERROR_MISSING_BOOKING_START_DATE → Start Date is missing
  ERROR_NEGATIVE_DIRECT_DEBIT_LIMIT → Direct Debit Limit is negative
  ERROR_NEGATIVE_FIXED_LIMIT       → Fixed Amount is negative
  ERROR_ONLY_ONE_AMOUNT_ALLOWED    → Both Limit and Fixed Amount are greater than zero

Additional checks:
  - Manual (Paused) plan Start Date must be a Monday
  - Service ID is required
  - Trailing spaces in names are stripped
  - Date format converted to DD/MM/YYYY
  - Weekday converted to 3-letter abbreviation
"""

import csv
import io
import os
import tempfile
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side


# ─────────────────────────── Constants ──────────────────────────────────────

DEFAULT_COLUMNS = {
    "service_id": "Service ID",
    "service": "Service Name",
    "parent_id": "Parent Legacy ID",
    "parent_fn": "Primary Guardian First Name",
    "parent_ln": "Primary Guardian Last Name",
    "child_id": "Child Legacy ID",
    "child_fn": "Child First Name",
    "child_ln": "Child Last Name",
    "date": "Start Date",
    "weekday": "Weekday",
    "manual": "Manual",
    "cycle": "Cycle",
    "limit": "Limit",
    "fixed_amount": "Fixed Amount",
    "gateway": "Gateway Reference",
}

# Label shown in UI for each column key
COLUMN_LABELS = {
    "service_id": "Service ID  *required",
    "service": "Service Name",
    "parent_id": "Parent Legacy ID",
    "parent_fn": "Parent First Name",
    "parent_ln": "Parent Last Name",
    "child_id": "Child Legacy ID",
    "child_fn": "Child First Name",
    "child_ln": "Child Last Name",
    "date": "Start Date",
    "weekday": "Weekday",
    "manual": "Manual (Yes/No)",
    "cycle": "Billing Cycle",
    "limit": "Direct Debit Limit",
    "fixed_amount": "Fixed Amount",
    "gateway": "Gateway Reference",
}

WEEKDAY_MAP = {
    "monday": "Mon", "tuesday": "Tue", "wednesday": "Wed",
    "thursday": "Thu", "friday": "Fri", "saturday": "Sat",
    "sunday": "Sun",
    "mon": "Mon", "tue": "Tue", "wed": "Wed", "thu": "Thu",
    "fri": "Fri", "sat": "Sat", "sun": "Sun",
}

VALID_CYCLES = {"weekly", "fortnightly", "monthly"}
WEEKEND_DAYS = {"Sat", "Sun"}

# Output CSV template — must match payment_plan_onboarding_tools.csv
TEMPLATE_COLUMNS = [
    "Service_ID", "Service_Name", "Parent_Legacy_Id",
    "Parent_First_Name", "Parent_Last_Name", "Child_Legacy_Id",
    "Child_First_Name", "Child_Last_Name", "Direct_Debit_Start_Date",
    "Direct_Debit_Day", "Manual", "Billing_Cycle",
    "Direct_Debit_Limit", "Fixed_Amount", "Gateway_Reference",
]

DATE_FORMATS = [
    "%d/%m/%Y %I:%M:%S %p", "%d/%m/%Y %I:%M %p",       # 12-hour with AM/PM
    "%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y",
    "%Y-%m-%d %I:%M:%S %p", "%Y-%m-%d %I:%M %p",
    "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d",
    "%m/%d/%Y %I:%M:%S %p", "%m/%d/%Y %I:%M %p",
    "%m/%d/%Y %H:%M:%S", "%m/%d/%Y %H:%M", "%m/%d/%Y",
    "%d/%m/%y",
]


# ─────────────────────────── Helpers ────────────────────────────────────────

def parse_date(raw: str):
    raw = raw.strip()
    if not raw:
        return None
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(raw, fmt)
        except ValueError:
            pass
    return None


def is_numeric(val: str) -> bool:
    try:
        float(val)
        return True
    except (ValueError, TypeError):
        return False


def _entry(row, col, key, default=""):
    col_name = col.get(key, "")
    return row.get(col_name, default).strip() if col_name else default


# ─────────────────────────── Core processor ─────────────────────────────────

def process_csv(filepath: str, col: dict) -> dict:
    with open(filepath, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames or []
        rows = list(reader)

    # Error buckets — order here controls display priority
    errors = {
        "weekend": [],   # ERROR_INVALID_PAYMENT_DAY (Sat/Sun)
        "missing_date": [],   # ERROR_MISSING_BOOKING_START_DATE
        "missing_weekday": [],   # ERROR_MISSING_PAYMENT_DAY
        "missing_parent": [],   # ERROR_MISSING_GUARDIAN
        "missing_service_id": [],   # Service ID required
        "invalid_cycle": [],   # ERROR_INVALID_FREQUENCY
        "manual_not_monday": [],   # Manual plan must start on Monday
        "negative_limit": [],   # ERROR_NEGATIVE_DIRECT_DEBIT_LIMIT
        "negative_fixed": [],   # ERROR_NEGATIVE_FIXED_LIMIT
        "both_amounts": [],   # ERROR_ONLY_ONE_AMOUNT_ALLOWED
        "unparseable_date": [],   # Date string unrecognised
        "unknown_weekday": [],   # Weekday value not in map
    }

    stats = {
        "total": len(rows),
        "date_fixed": 0,
        "weekday_fixed": 0,
        "spaces_fixed": 0,
    }

    processed = []
    for i, raw_row in enumerate(rows, start=2):
        row = dict(raw_row)

        # ── Strip all whitespace ──────────────────────────────────────────
        for key in row:
            stripped = row[key].strip()
            if stripped != row[key]:
                stats["spaces_fixed"] += 1
            row[key] = stripped

        # Shorthand helpers for this row
        def g(key): return _entry(row, col, key)

        parent_fn = g("parent_fn")
        parent_ln = g("parent_ln")
        parent_name = f"{parent_fn} {parent_ln}".strip()
        service = g("service")
        parent_id = g("parent_id")
        child_id = g("child_id")
        service_id = g("service_id")

        ctx = dict(row=i, parent_id=parent_id, child_id=child_id,
                   parent_name=parent_name, service=service)

        # ── 1. Fix & validate Start Date ─────────────────────────────────
        raw_date = g("date")
        parsed_dt = None
        if raw_date:
            parsed_dt = parse_date(raw_date)
            if parsed_dt:
                fixed = parsed_dt.strftime("%d/%m/%Y")
                if fixed != raw_date:
                    stats["date_fixed"] += 1
                col_name = col.get("date", "")
                if col_name:
                    row[col_name] = fixed
            else:
                errors["unparseable_date"].append({**ctx, "value": raw_date})
        else:
            errors["missing_date"].append({**ctx, "weekday": g("weekday") or "(empty)"})

        # ── 2. Fix & validate Weekday ─────────────────────────────────────
        raw_wd = g("weekday")
        wd_fixed = WEEKDAY_MAP.get(raw_wd.lower(), "")
        col_name_wd = col.get("weekday", "")
        if wd_fixed:
            if wd_fixed != raw_wd:
                stats["weekday_fixed"] += 1
            if col_name_wd:
                row[col_name_wd] = wd_fixed
        elif raw_wd:
            errors["unknown_weekday"].append({**ctx, "value": raw_wd})
        else:
            errors["missing_weekday"].append({**ctx, "date": raw_date or "(empty)"})

        # ── 3. Weekend check ──────────────────────────────────────────────
        if wd_fixed in WEEKEND_DAYS:
            date_str = row.get(col.get("date", ""), "")
            errors["weekend"].append({**ctx, "weekday": wd_fixed, "date": date_str})

        # ── 4. Parent name ────────────────────────────────────────────────
        if not parent_fn or not parent_ln:
            errors["missing_parent"].append({
                **ctx,
                "first_name": parent_fn or "(empty)",
                "last_name": parent_ln or "(empty)",
            })

        # ── 5. Service ID required ────────────────────────────────────────
        if not service_id:
            errors["missing_service_id"].append(ctx)

        # ── 6. Billing Cycle validation ───────────────────────────────────
        cycle = g("cycle")
        if cycle and cycle.lower() not in VALID_CYCLES:
            errors["invalid_cycle"].append({**ctx, "value": cycle})

        # ── 7. Manual plan → start date must be Monday ───────────────────
        manual_val = g("manual").lower()
        is_manual = manual_val in ("yes", "1", "true")
        if is_manual and parsed_dt and parsed_dt.weekday() != 0:   # 0 = Monday
            day_name = parsed_dt.strftime("%A")
            errors["manual_not_monday"].append({
                **ctx,
                "date": parsed_dt.strftime("%d/%m/%Y"),
                "day": day_name,
            })

        # ── 8. Limit / Fixed Amount checks ───────────────────────────────
        limit_raw = g("limit")
        fixed_raw = g("fixed_amount")

        if limit_raw:
            if is_numeric(limit_raw):
                if float(limit_raw) < 0:
                    errors["negative_limit"].append({**ctx, "value": limit_raw})
            # (non-numeric → the tool imports as 0, so we don't flag it here)

        if fixed_raw:
            if is_numeric(fixed_raw):
                if float(fixed_raw) < 0:
                    errors["negative_fixed"].append({**ctx, "value": fixed_raw})

        # Both amounts > 0 (ERROR_ONLY_ONE_AMOUNT_ALLOWED)
        try:
            limit_num = float(limit_raw) if limit_raw else 0
            fixed_num = float(fixed_raw) if fixed_raw else 0
            if limit_num > 0 and fixed_num > 0:
                errors["both_amounts"].append({
                    **ctx,
                    "limit": limit_raw,
                    "fixed": fixed_raw,
                })
        except (ValueError, TypeError):
            pass

        row["_row_num"] = i   # preserve original file row number for later mapping
        processed.append(row)

    return {
        "processed_rows": processed,
        "fieldnames": fieldnames,
        "errors": errors,
        "stats": stats,
    }


# ─────────────────────────── File writers ───────────────────────────────────

def _output_dir(filepath: str) -> Path:
    """Return (and create) the Output folder next to filepath."""
    src = Path(filepath).parent
    base = src.parent if src.name.lower() == "input" else src
    out = base / "Output"
    out.mkdir(exist_ok=True)
    return out


def _resolve_out_dir(filepath: str, out_dir: str | None) -> Path:
    """Return out_dir as a Path (creating it), or fall back to _output_dir."""
    if out_dir:
        p = Path(out_dir)
        p.mkdir(parents=True, exist_ok=True)
        return p
    return _output_dir(filepath)


def write_cleaned_csv(filepath: str, fieldnames: list, rows: list, out_dir: str | None = None) -> str:
    out = _resolve_out_dir(filepath, out_dir) / f"{Path(filepath).stem}_cleaned.csv"
    with open(out, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    return str(out)


def load_service_mapping(csv_path: str) -> dict:
    """Load serviceIDs.csv → {qk_service_id_str: {"xplor_id": str, "name": str}}."""
    mapping = {}
    with open(csv_path, "r", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            qk_id = row.get("QKServiceID", "").strip()
            xplor_id = row.get("Xplor Service ID", "").strip()
            name = row.get("Service Name", "").strip()
            if qk_id:
                mapping[qk_id] = {"xplor_id": xplor_id, "name": name}
    return mapping


def _safe_filename(name: str) -> str:
    """Remove characters that are invalid in Windows filenames."""
    for ch in r'\/:*?"<>|':
        name = name.replace(ch, "_")
    return name.strip()


def write_split_csvs(filepath: str, col: dict, rows: list, service_map: dict, out_dir: str | None = None) -> dict:
    """
    Split processed rows by Service ID and write one CSV per service.

    - Rows whose Service ID is found in service_map  → Output/<ServiceName>_payment_plan_import.csv
    - Rows whose Service ID is NOT found             → Output/unknown/<id>_payment_plan_import.csv
    - Service Name column is replaced with the correct name from service_map.
    - All output files follow TEMPLATE_COLUMNS column order.

    Returns {"known": {service_name: path}, "unknown": {service_id: path}}
    """
    out_dir = _resolve_out_dir(filepath, out_dir)
    unknown_dir = out_dir / "unknown"

    # Group rows by service_id value
    groups: dict[str, list] = {}
    for row in rows:
        sid = _entry(row, col, "service_id")
        groups.setdefault(sid, []).append(row)

    known_files: dict[str, str] = {}
    unknown_files: dict[str, str] = {}
    row_map: dict[int, int] = {}   # {original_row_num: service_file_row_num}

    for sid, group_rows in groups.items():
        svc = service_map.get(sid)
        is_known = svc is not None
        if is_known:
            service_name = svc["name"]
            xplor_id = svc["xplor_id"]
            dest = out_dir / (_safe_filename(service_name) + "_payment_plan_import.csv")
        else:
            unknown_dir.mkdir(exist_ok=True)
            label = sid if sid else "no_service_id"
            dest = unknown_dir / (_safe_filename(label) + "_payment_plan_import.csv")
            service_name = _entry(group_rows[0], col, "service")   # keep original
            xplor_id = sid

        template_rows = []
        for file_row_num, row in enumerate(group_rows, start=2):  # row 1 = header
            orig = row.get("_row_num")
            if orig is not None:
                row_map[orig] = file_row_num

            def g(key, _row=row): return _entry(_row, col, key)
            template_rows.append({
                "Service_ID": xplor_id,
                "Service_Name": service_name,
                "Parent_Legacy_Id": g("parent_id"),
                "Parent_First_Name": g("parent_fn"),
                "Parent_Last_Name": g("parent_ln"),
                "Child_Legacy_Id": f"{xplor_id}_{g('child_id')}" if g("child_id") else "",
                "Child_First_Name": g("child_fn"),
                "Child_Last_Name": g("child_ln"),
                "Direct_Debit_Start_Date": g("date"),
                "Direct_Debit_Day": g("weekday"),
                "Manual": g("manual"),
                "Billing_Cycle": g("cycle"),
                "Direct_Debit_Limit": g("limit"),
                "Fixed_Amount": g("fixed_amount"),
                "Gateway_Reference": g("gateway"),
            })

        with open(dest, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.DictWriter(f, fieldnames=TEMPLATE_COLUMNS)
            writer.writeheader()
            writer.writerows(template_rows)

        if is_known:
            known_files[service_name] = str(dest)
        else:
            unknown_files[sid or "(empty)"] = str(dest)

    return {"known": known_files, "unknown": unknown_files, "row_map": row_map}


def _translate_error_rows(errors: dict, row_map: dict) -> None:
    """Replace original-file row numbers with service-file row numbers in-place."""
    for items in errors.values():
        for e in items:
            orig = e.get("row")
            if orig in row_map:
                e["row"] = row_map[orig]


def write_error_report(filepath: str, errors: dict, out_dir: str | None = None) -> str:
    """Write a colour-coded Excel error report (.xlsx)."""
    out = _resolve_out_dir(filepath, out_dir) / f"{Path(filepath).stem}_error_report.xlsx"

    # ── Colour palette ───────────────────────────────────────────────────
    C_WEEKEND_HEADER = "C0392B"   # deep red
    C_WEEKEND_ROW = "FADBD8"   # light pink
    C_ERROR_HEADER = "E67E22"   # orange
    C_ERROR_ROW = "FDEBD0"   # light orange
    C_WARN_HEADER = "F1C40F"   # yellow
    C_WARN_ROW = "FEF9E7"   # light yellow
    C_COL_HEADER = "1A5276"   # navy (column header row)
    C_WHITE = "FFFFFF"
    C_LIGHT_GRAY = "F2F2F2"

    def fill(hex_color):
        return PatternFill("solid", fgColor=hex_color)

    def bold_font(color=None, size=10):
        return Font(bold=True, color=color or "000000", size=size)

    def std_font(color=None):
        return Font(color=color or "000000", size=10)

    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    # ── Error definitions: (key, tool_error_key, display_label, note_fn, severity)
    # severity: "weekend" | "error" | "warn"
    PRIORITY = [
        ("weekend",
         "ERROR_INVALID_PAYMENT_DAY",
         "WEEKEND — Sat / Sun",
         lambda e: f"Falls on a {
             'Sunday' if e.get('weekday') == 'Sun' else 'Saturday'} ({
             e.get(
                 'date',
                 '')}) — please verify before importing",
            "weekend"),
        ("missing_date",
         "ERROR_MISSING_BOOKING_START_DATE",
         "Missing Start Date",
         lambda e: "Start Date is empty — cannot import",
         "error"),
        ("missing_weekday",
         "ERROR_MISSING_PAYMENT_DAY",
         "Missing Weekday",
         lambda e: "Weekday is empty — cannot import",
         "error"),
        ("missing_parent",
         "ERROR_MISSING_GUARDIAN",
         "Missing Parent Name",
         lambda e: f"First: \"{
             e.get(
                 'first_name',
                 '')}\"  Last: \"{
             e.get(
                 'last_name',
                 '')}\"",
         "error"),
        ("missing_service_id",
         "ERROR_MISSING_SERVICE_ID",
         "Missing Service ID",
         lambda e: "Service ID is empty (required field)",
         "error"),
        ("invalid_cycle",
         "ERROR_INVALID_FREQUENCY",
         "Invalid Billing Cycle",
         lambda e: f"Value: \"{
             e.get(
                 'value',
                 '')}\" — must be Weekly / Fortnightly / Monthly",
         "error"),
        ("manual_not_monday",
         "MANUAL_PLAN_NOT_MONDAY",
         "Manual Plan — Not Monday",
         lambda e: f"Starts on {
             e.get(
                 'day',
                 '')} ({
             e.get(
                 'date',
                 '')}) — Manual plans must start on Monday",
         "error"),
        ("negative_limit",
         "ERROR_NEGATIVE_DIRECT_DEBIT_LIMIT",
         "Negative Direct Debit Limit",
         lambda e: f"Value = {
             e.get(
                 'value',
                 '')}  (must be >= 0)",
         "error"),
        ("negative_fixed",
         "ERROR_NEGATIVE_FIXED_LIMIT",
         "Negative Fixed Amount",
         lambda e: f"Value = {
             e.get(
                 'value',
                 '')}  (must be >= 0)",
         "error"),
        ("both_amounts",
         "ERROR_ONLY_ONE_AMOUNT_ALLOWED",
         "Both Limit + Fixed Amount Set",
         lambda e: f"Limit={
             e.get(
                 'limit',
                 '')}  AND  Fixed={
             e.get(
                 'fixed',
                 '')}  (only one allowed)",
         "error"),
        ("unparseable_date",
         "UNPARSEABLE_DATE",
         "Unparseable Date Value",
         lambda e: f"Cannot parse: \"{
             e.get(
                 'value',
                 '')}\"",
         "warn"),
        ("unknown_weekday",
         "UNKNOWN_WEEKDAY",
         "Unknown Weekday Value",
         lambda e: f"Unrecognised value: \"{
             e.get(
                 'value',
                 '')}\"",
         "warn"),
    ]

    wb = Workbook()

    # ════════════════════════════════════════════════════════════════════
    # Sheet 1: Summary
    # ════════════════════════════════════════════════════════════════════
    ws_sum = wb.active
    ws_sum.title = "Summary"
    ws_sum.sheet_view.showGridLines = False

    # Title
    ws_sum.merge_cells("A1:D1")
    ws_sum["A1"] = "Payment Plan Import — Error Report"
    ws_sum["A1"].font = Font(bold=True, size=16, color=C_WHITE)
    ws_sum["A1"].fill = fill(C_COL_HEADER)
    ws_sum["A1"].alignment = center
    ws_sum.row_dimensions[1].height = 30

    ws_sum.merge_cells("A2:D2")
    ws_sum["A2"] = f"Generated: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}  |  Source: {Path(filepath).name}"
    ws_sum["A2"].font = Font(italic=True, size=9, color="666666")
    ws_sum["A2"].alignment = left
    ws_sum.row_dimensions[2].height = 16

    ws_sum.append([])

    # Column headers
    ws_sum.append(["Error Category", "Onboarding Tool Error Key", "Count", "Severity"])
    for col_idx, cell in enumerate(ws_sum[ws_sum.max_row], 1):
        cell.font = bold_font(C_WHITE)
        cell.fill = fill(C_COL_HEADER)
        cell.alignment = center
        cell.border = border
    ws_sum.row_dimensions[ws_sum.max_row].height = 20

    total = 0
    for key, tool_key, label, _, sev in PRIORITY:
        n = len(errors.get(key, []))
        total += n
        row_data = [label, tool_key, n, sev.upper()]
        ws_sum.append(row_data)
        r = ws_sum.max_row
        c_fill = fill(C_WEEKEND_ROW if sev == "weekend"
                      else C_ERROR_ROW if sev == "error"
                      else C_WARN_ROW)
        c_font_count = bold_font("C0392B") if n > 0 else std_font("2ECC71")
        for col_idx, cell in enumerate(ws_sum[r], 1):
            cell.fill = c_fill
            cell.alignment = left if col_idx <= 2 else center
            cell.border = border
            cell.font = std_font()
        ws_sum[r][2].font = c_font_count  # count cell
        ws_sum.row_dimensions[r].height = 18

    # Total row
    ws_sum.append(["TOTAL ERRORS", "", total, ""])
    r = ws_sum.max_row
    t_fill = fill("C0392B") if total > 0 else fill("1E8449")
    for cell in ws_sum[r]:
        cell.font = bold_font(C_WHITE, size=11)
        cell.fill = t_fill
        cell.alignment = center
        cell.border = border
    ws_sum.row_dimensions[r].height = 22

    ws_sum.column_dimensions["A"].width = 36
    ws_sum.column_dimensions["B"].width = 38
    ws_sum.column_dimensions["C"].width = 10
    ws_sum.column_dimensions["D"].width = 14

    # ════════════════════════════════════════════════════════════════════
    # Sheet 2: Detail (all errors)
    # ════════════════════════════════════════════════════════════════════
    ws_det = wb.create_sheet("Error Detail")
    ws_det.sheet_view.showGridLines = False
    ws_det.freeze_panes = "A3"

    # Title row
    ws_det.merge_cells("A1:I1")
    ws_det["A1"] = "Error Detail — All Issues Found"
    ws_det["A1"].font = Font(bold=True, size=13, color=C_WHITE)
    ws_det["A1"].fill = fill(C_COL_HEADER)
    ws_det["A1"].alignment = center
    ws_det.row_dimensions[1].height = 24

    # Column headers
    COLS = ["#", "Error Category", "Tool Error Key",
            "Row #", "Parent Legacy ID", "Child Legacy ID",
            "Parent Name", "Service Name", "Note / Detail"]
    ws_det.append(COLS)
    for cell in ws_det[2]:
        cell.font = bold_font(C_WHITE)
        cell.fill = fill(C_COL_HEADER)
        cell.alignment = center
        cell.border = border
    ws_det.row_dimensions[2].height = 20

    seq = 0
    for key, tool_key, label, note_fn, sev in PRIORITY:
        items = errors.get(key, [])
        if not items:
            continue

        # Section divider
        ws_det.append(["", f"── {label.upper()}  ({len(items)} rows) ──"] + [""] * 7)
        r = ws_det.max_row
        ws_det.merge_cells(f"B{r}:I{r}")
        c_hdr = (C_WEEKEND_HEADER if sev == "weekend"
                 else C_ERROR_HEADER if sev == "error"
                 else C_WARN_HEADER)
        for cell in ws_det[r]:
            cell.fill = fill(c_hdr)
            cell.font = bold_font(C_WHITE)
            cell.alignment = left
        ws_det.row_dimensions[r].height = 18

        c_row_fill = fill(C_WEEKEND_ROW if sev == "weekend"
                          else C_ERROR_ROW if sev == "error"
                          else C_WARN_ROW)
        alt_fill = fill(C_LIGHT_GRAY)

        for i, e in enumerate(items):
            seq += 1
            note = note_fn(e)
            ws_det.append([
                seq, label, tool_key,
                e.get("row", ""), e.get("parent_id", ""), e.get("child_id", ""),
                e.get("parent_name", ""), e.get("service", ""), note,
            ])
            r = ws_det.max_row
            row_fill = c_row_fill if i % 2 == 0 else alt_fill
            for col_idx, cell in enumerate(ws_det[r], 1):
                cell.fill = row_fill
                cell.border = border
                cell.alignment = center if col_idx in (1, 4) else left
                cell.font = std_font()
            ws_det.row_dimensions[r].height = 16

    ws_det.column_dimensions["A"].width = 5
    ws_det.column_dimensions["B"].width = 30
    ws_det.column_dimensions["C"].width = 34
    ws_det.column_dimensions["D"].width = 7
    ws_det.column_dimensions["E"].width = 16
    ws_det.column_dimensions["F"].width = 16
    ws_det.column_dimensions["G"].width = 28
    ws_det.column_dimensions["H"].width = 40
    ws_det.column_dimensions["I"].width = 55

    # ════════════════════════════════════════════════════════════════════
    # Sheet 3: Weekends only (quick-check sheet)
    # ════════════════════════════════════════════════════════════════════
    ws_wk = wb.create_sheet("Weekend Errors")
    ws_wk.sheet_view.showGridLines = False
    ws_wk.freeze_panes = "A3"

    ws_wk.merge_cells("A1:G1")
    ws_wk["A1"] = "*** WEEKEND ERRORS — Please verify before importing ***"
    ws_wk["A1"].font = Font(bold=True, size=13, color=C_WHITE)
    ws_wk["A1"].fill = fill(C_WEEKEND_HEADER)
    ws_wk["A1"].alignment = center
    ws_wk.row_dimensions[1].height = 24

    WK_COLS = ["Row #", "Weekday", "Start Date", "Parent Legacy ID",
               "Child Legacy ID", "Parent Name", "Service Name"]
    ws_wk.append(WK_COLS)
    for cell in ws_wk[2]:
        cell.font = bold_font(C_WHITE)
        cell.fill = fill(C_WEEKEND_HEADER)
        cell.alignment = center
        cell.border = border
    ws_wk.row_dimensions[2].height = 20

    weekend_items = errors.get("weekend", [])
    if weekend_items:
        for i, e in enumerate(weekend_items):
            ws_wk.append([
                e.get("row", ""), e.get("weekday", ""), e.get("date", ""),
                e.get("parent_id", ""), e.get("child_id", ""),
                e.get("parent_name", ""), e.get("service", ""),
            ])
            r = ws_wk.max_row
            row_fill = fill(C_WEEKEND_ROW) if i % 2 == 0 else fill("F5B7B1")
            for col_idx, cell in enumerate(ws_wk[r], 1):
                cell.fill = row_fill
                cell.border = border
                cell.alignment = center if col_idx <= 3 else left
                cell.font = std_font()
                if col_idx == 2:
                    cell.font = bold_font("C0392B")
            ws_wk.row_dimensions[r].height = 16
    else:
        ws_wk.merge_cells("A3:G3")
        ws_wk["A3"] = "✓ No Weekend Errors found"
        ws_wk["A3"].font = Font(bold=True, size=11, color="1E8449")
        ws_wk["A3"].alignment = center
        ws_wk["A3"].fill = fill("D5F5E3")

    ws_wk.column_dimensions["A"].width = 7
    ws_wk.column_dimensions["B"].width = 10
    ws_wk.column_dimensions["C"].width = 14
    ws_wk.column_dimensions["D"].width = 16
    ws_wk.column_dimensions["E"].width = 16
    ws_wk.column_dimensions["F"].width = 30
    ws_wk.column_dimensions["G"].width = 45

    wb.save(out)
    return str(out)


# ─────────────────────────── Streamlit wrapper ──────────────────────────────

def run_payment_plan_checker(
    input_bytes: bytes,
    filename: str,
    service_ids_bytes: bytes,
    output_dir: str,
    col_mapping: dict | None = None,
) -> dict:
    """Validate and process a payment plan CSV from in-memory bytes.

    Parameters
    ----------
    input_bytes:
        Raw bytes of the input CSV.
    filename:
        Original filename (used for output file naming).
    service_ids_bytes:
        Raw bytes of serviceIDs.csv.
    output_dir:
        Absolute path where all output files will be saved.
    col_mapping:
        Column name mapping dict.  Defaults to DEFAULT_COLUMNS if None.

    Returns
    -------
    dict with keys: result (process_csv output), cleaned_path, error_path,
    split_result, col (the mapping used).
    """
    col = col_mapping or DEFAULT_COLUMNS.copy()

    # Write bytes to a temp file so process_csv can read it
    with tempfile.NamedTemporaryFile(
        suffix=".csv", delete=False, mode="wb", prefix=Path(filename).stem + "_"
    ) as tmp:
        tmp.write(input_bytes)
        tmp_path = tmp.name

    try:
        result = process_csv(tmp_path, col)
    finally:
        try:
            os.unlink(tmp_path)
        except OSError:
            pass

    # Use a temp path for naming purposes only (out_dir override handles location)
    naming_path = str(Path(output_dir) / filename)

    cleaned_path = write_cleaned_csv(naming_path, result["fieldnames"], result["processed_rows"], out_dir=output_dir)

    # Load service mapping from bytes
    service_map: dict = {}
    if service_ids_bytes:
        text = service_ids_bytes.decode("utf-8-sig", errors="replace")
        import csv as _csv
        reader = _csv.DictReader(io.StringIO(text))
        for row in reader:
            qk_id = row.get("QKServiceID", "").strip()
            xplor_id = row.get("Xplor Service ID", "").strip()
            name = row.get("Service Name", "").strip()
            if qk_id:
                service_map[qk_id] = {"xplor_id": xplor_id, "name": name}

    split_result = write_split_csvs(naming_path, col, result["processed_rows"], service_map, out_dir=output_dir)
    _translate_error_rows(result["errors"], split_result["row_map"])
    error_path = write_error_report(naming_path, result["errors"], out_dir=output_dir)

    return {
        "result": result,
        "cleaned_path": cleaned_path,
        "error_path": error_path,
        "split_result": split_result,
        "col": col,
    }


# ─────────────────────────── Entry point ────────────────────────────────────

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 4:
        print("Usage: python payment_plan_checker.py <input.csv> <serviceIDs.csv> <output_dir>")
        sys.exit(1)

    input_path = sys.argv[1]
    svc_path = sys.argv[2]
    out_dir = sys.argv[3]

    input_bytes = Path(input_path).read_bytes()
    svc_bytes = Path(svc_path).read_bytes()

    result_dict = run_payment_plan_checker(input_bytes, Path(input_path).name, svc_bytes, out_dir)
    stats = result_dict["result"]["stats"]
    errors = result_dict["result"]["errors"]
    total_err = sum(len(v) for v in errors.values())

    print(f"Rows processed   : {stats['total']}")
    print(f"Date fixed       : {stats['date_fixed']}")
    print(f"Weekday fixed    : {stats['weekday_fixed']}")
    print(f"Spaces stripped  : {stats['spaces_fixed']}")
    print(f"Total errors     : {total_err}")
    print(f"Cleaned CSV      : {result_dict['cleaned_path']}")
    print(f"Error report     : {result_dict['error_path']}")
