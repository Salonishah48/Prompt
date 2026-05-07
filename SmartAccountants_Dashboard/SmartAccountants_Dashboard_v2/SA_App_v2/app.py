"""
Smart Accountants — Timesheet Dashboard Generator v2
Includes: Missing Time Entries tab
Run via: run_dashboard.bat
"""

import os, sys, json, math, glob, webbrowser, shutil
from datetime import datetime, date
from pathlib import Path

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas not installed. Please run install_once.bat first.")
    input("Press Enter to exit..."); sys.exit(1)

try:
    import openpyxl
except ImportError:
    print("ERROR: openpyxl not installed. Please run install_once.bat first.")
    input("Press Enter to exit..."); sys.exit(1)

BASE_DIR   = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "output"
TEMPLATE   = BASE_DIR / "dashboard_template.html"
OUTPUT_DIR.mkdir(exist_ok=True)

# ── Find Excel file ───────────────────────────────────────────────────────────
def find_excel():
    candidates = [f for f in BASE_DIR.glob("*.xls*") if not f.name.startswith("~$")]
    if not candidates:
        return None
    candidates.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    return candidates[0]

# ── Rounding ──────────────────────────────────────────────────────────────────
def r1(v):  return round(float(v), 1)
def r0(v):  return round(float(v))

# ── Categorise ────────────────────────────────────────────────────────────────
LEAVE_KEYWORDS = ["sick", "personal", "holiday", "vacation"]

def categorise(row):
    sl = row["service"].lower()
    for k in LEAVE_KEYWORDS:
        if k in sl: return "leave"
    if "Non Billable" in row["client"] or "NONB" in row["client"]:
        return "non_billable"
    return "billable"

# ── Working days in date range ────────────────────────────────────────────────
def count_working_days(dates):
    """Count Mon–Sat working days (excluding Sun) in the date list."""
    working = 0
    for d in dates:
        dt = datetime.strptime(d, "%Y-%m-%d")
        if dt.weekday() != 6:   # 6 = Sunday
            working += 1
    return working

# ── Parse Time and Expense sheet ─────────────────────────────────────────────
def parse_timesheet(path):
    print(f"  Reading timesheet: {path.name}")
    df_raw = pd.read_excel(path, sheet_name="Time and Expense", header=None)
    records = []
    current_staff = None
    rows = list(df_raw.iterrows())
    i = 0
    while i < len(rows):
        idx, row = rows[i]
        val1 = row[1]
        if isinstance(val1, str) and "Staff :" in val1:
            current_staff = val1.replace("Staff :", "").strip()
            i += 1; continue
        if pd.notna(val1) and hasattr(val1, "year"):
            rec = {
                "staff":    current_staff,
                "date":     val1.strftime("%Y-%m-%d"),
                "client":   str(row[3]),
                "service":  str(row[4]),
                "status":   str(row[6]),
                "hours":    float(row[7]) if pd.notna(row[7]) else 0,
                "inv_desc": "",
                "int_note": "",
            }
            if i+1 < len(rows):
                r1v = str(rows[i+1][1][2]) if pd.notna(rows[i+1][1][2]) else ""
                if "Invoice Description:" in r1v:
                    rec["inv_desc"] = r1v.replace("Invoice Description:", "").strip()
            if i+2 < len(rows):
                r2v = str(rows[i+2][1][2]) if pd.notna(rows[i+2][1][2]) else ""
                if "Internal Note:" in r2v:
                    rec["int_note"] = r2v.replace("Internal Note:", "").strip()
            records.append(rec)
        i += 1
    return pd.DataFrame(records)

# ── Parse All Employee Details sheet ─────────────────────────────────────────
def parse_employee_details(path):
    """
    Reads employee master data from the Excel file.

    Sheet name priority (first match wins):
      1. Exact: 'List of employees - SA'   <- primary target
      2. Contains 'list of employees'
      3. Contains 'employee' + 'detail'
      4. Contains 'all employee'
      5. Any sheet containing 'employee'

    Expected columns (flexible matching):
      Emp ID / Employee ID / Staff ID
      Employee Name / Name / Staff Name / Staff
      Department / Dept
      T1 Manager / Manager / T1
      Reporting Manager / Report Manager
      Expected Hours / Expected Hrs / Expected / Target Hours
    """
    xl = pd.ExcelFile(path)
    all_sheets = xl.sheet_names

    # Priority sheet name matching
    PRIORITY_PATTERNS = [
        lambda s: s.strip().lower() == "list of employees - sa",
        lambda s: "list of employees" in s.lower(),
        lambda s: "employee" in s.lower() and "detail" in s.lower(),
        lambda s: "all employee" in s.lower(),
        lambda s: "employee" in s.lower(),
    ]

    sheet_name = None
    for pattern in PRIORITY_PATTERNS:
        for s in all_sheets:
            if pattern(s):
                sheet_name = s
                break
        if sheet_name:
            break

    if sheet_name is None:
        print("  NOTE: Employee list sheet not found — Missing Time Entries will use default expected hours.")
        print(f"        Available sheets: {all_sheets}")
        return None

    print(f"  Reading employee details: '{sheet_name}'")
    df = pd.read_excel(path, sheet_name=sheet_name)
    df.dropna(how="all", inplace=True)
    df.columns = [str(c).strip() for c in df.columns]
    print(f"        Columns found: {list(df.columns)}")
    print(f"        Rows: {len(df)}")

    # Flexible column mapping
    col_map = {}
    for col in df.columns:
        cl = col.lower().strip()

        if "emp_id" not in col_map:
            if any(x in cl for x in ["emp id", "emp_id", "empid", "employee id", "staff id", "staffid"]):
                col_map["emp_id"] = col
            elif cl == "id":
                col_map["emp_id"] = col

        if "emp_name" not in col_map:
            if any(x in cl for x in ["employee name", "emp name", "staff name", "full name", "name of employee"]):
                col_map["emp_name"] = col
            elif cl in ("name", "staff"):
                col_map["emp_name"] = col

        if "dept" not in col_map:
            if any(x in cl for x in ["department", "dept", "division", "team"]):
                col_map["dept"] = col

        if "t1_manager" not in col_map:
            if ("t1" in cl and "manager" in cl) or ("t1" in cl and "mgr" in cl):
                col_map["t1_manager"] = col
            elif cl in ("t1 manager", "t1manager", "manager", "t1"):
                col_map["t1_manager"] = col

        if "reporting_manager" not in col_map:
            if "reporting" in cl and ("manager" in cl or "mgr" in cl):
                col_map["reporting_manager"] = col
            elif "report" in cl and "manager" in cl:
                col_map["reporting_manager"] = col

        if "expected" not in col_map:
            if any(x in cl for x in ["expected hours", "expected hrs", "target hours",
                                      "target hrs", "weekly hours", "weekly hrs",
                                      "std hours", "standard hours", "exp hours", "exp hrs"]):
                col_map["expected"] = col
            elif cl in ("expected", "target", "hours expected"):
                col_map["expected"] = col

    print(f"        Column mapping: { {k: v for k, v in col_map.items()} }")

    if "emp_name" not in col_map:
        print("  WARNING: Could not identify Employee Name column — check sheet column headers.")
        return None

    return df, col_map, sheet_name

# ── Build all dashboard data ──────────────────────────────────────────────────
def build_data(df, emp_df_info=None):
    df["category"] = df.apply(categorise, axis=1)

    # ── Summary per staff ─────────────────────────────────────────────────────
    summary = []
    for staff, g in df.groupby("staff"):
        b  = r1(g[g.category == "billable"]["hours"].sum())
        l  = r1(g[g.category == "leave"]["hours"].sum())
        n  = r1(g[g.category == "non_billable"]["hours"].sum())
        t  = r1(g["hours"].sum())
        e  = r1(t - l)
        be = r0(b / e * 100) if e > 0 else 0
        ne = r0(n / e * 100) if e > 0 else 0
        summary.append({
            "name": staff.split("(")[0].strip(),
            "billable": b, "leave": l, "non_billable": n,
            "total": t, "effective": e,
            "bill_eff": be, "nonbill_eff": ne, "total_eff": be + ne,
        })
    summary.sort(key=lambda x: -x["billable"])

    # ── Daily detail ──────────────────────────────────────────────────────────
    daily = {}
    for staff, g in df.groupby("staff"):
        nm = staff.split("(")[0].strip()
        days = {}
        for date_val, dg in g.groupby("date"):
            entries = []
            for _, row in dg.iterrows():
                cl = row["client"].split("(")[0].strip()
                if len(cl) > 40: cl = cl[:40] + "…"
                svc = row["service"]
                if "(" in svc: svc = svc[:svc.rfind("(")].strip()
                entries.append({
                    "client":   cl, "service":  svc,
                    "hours":    r1(row["hours"]),
                    "category": row["category"], "status": row["status"],
                    "inv_desc": row["inv_desc"][:120] if row["inv_desc"] else "",
                    "int_note": row["int_note"][:200] if row["int_note"] else "",
                })
            days[date_val] = entries
        daily[nm] = days

    # ── KPI data ──────────────────────────────────────────────────────────────
    kpi_billable = []
    for staff, g in df[df.category == "billable"].groupby("staff"):
        nm = staff.split("(")[0].strip()
        by_day = {d: r1(h) for d, h in g.groupby("date")["hours"].sum().items()}
        kpi_billable.append({"name": nm, "total": r1(g["hours"].sum()), "by_day": by_day})
    kpi_billable.sort(key=lambda x: -x["total"])

    kpi_leave = []
    for staff, g in df[df.category == "leave"].groupby("staff"):
        nm = staff.split("(")[0].strip()
        entries = []
        for _, row in g.iterrows():
            svc = row["service"]
            if "(" in svc: svc = svc[:svc.rfind("(")].strip()
            entries.append({"date": row["date"], "service": svc,
                            "hours": r1(row["hours"]),
                            "note": row["int_note"][:100] if row["int_note"] else ""})
        kpi_leave.append({"name": nm, "total": r1(g["hours"].sum()), "entries": entries})
    kpi_leave.sort(key=lambda x: -x["total"])

    kpi_nonbill = []
    for staff, g in df[df.category == "non_billable"].groupby("staff"):
        nm = staff.split("(")[0].strip()
        entries = []
        for _, row in g.iterrows():
            svc = row["service"]
            if "(" in svc: svc = svc[:svc.rfind("(")].strip()
            entries.append({"date": row["date"], "service": svc,
                            "hours": r1(row["hours"]),
                            "note": row["int_note"][:100] if row["int_note"] else ""})
        kpi_nonbill.append({"name": nm, "total": r1(g["hours"].sum()), "entries": entries})
    kpi_nonbill.sort(key=lambda x: -x["total"])

    kpi_total = [{"name": s["name"], "total": s["total"],
                  "billable": s["billable"], "leave": s["leave"],
                  "non_billable": s["non_billable"]} for s in summary]

    # ── Grand totals ──────────────────────────────────────────────────────────
    t_bill  = r1(df[df.category == "billable"]["hours"].sum())
    t_leave = r1(df[df.category == "leave"]["hours"].sum())
    t_nonb  = r1(df[df.category == "non_billable"]["hours"].sum())
    t_total = r1(df["hours"].sum())
    t_eff   = r1(t_total - t_leave)
    gt_be   = r0(t_bill / t_eff * 100) if t_eff > 0 else 0
    gt_ne   = r0(t_nonb / t_eff * 100) if t_eff > 0 else 0
    totals  = {"total": t_total, "billable": t_bill, "leave": t_leave,
               "non_billable": t_nonb, "effective": t_eff,
               "bill_eff": gt_be, "nonbill_eff": gt_ne, "total_eff": gt_be + gt_ne}

    # ── Date range ────────────────────────────────────────────────────────────
    dates = sorted(df["date"].unique())
    date_range = f"{dates[0]} to {dates[-1]}" if dates else "Unknown"
    try:
        d1 = datetime.strptime(dates[0], "%Y-%m-%d")
        d2 = datetime.strptime(dates[-1], "%Y-%m-%d")
        date_range = f"{d1.strftime('%b %d')} – {d2.strftime('%b %d, %Y')}"
    except Exception:
        pass

    # ── Missing Time Entries ──────────────────────────────────────────────────
    missing_entries = build_missing_entries(df, summary, dates, emp_df_info)

    return {
        "summary":        summary,
        "daily":          daily,
        "kpi_billable":   kpi_billable,
        "kpi_leave":      kpi_leave,
        "kpi_nonbill":    kpi_nonbill,
        "kpi_total":      kpi_total,
        "totals":         totals,
        "date_range":     date_range,
        "record_count":   len(df),
        "staff_count":    len(summary),
        "leave_count":    len(kpi_leave),
        "missing":        missing_entries,
    }

# ── Safe Emp ID converter ─────────────────────────────────────────────────────
def _safe_emp_id(val):
    """Convert emp_id to string safely — handles int, float, str, username codes."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "—"
    try:
        # Pure number (int or float like 1154.0) → format without decimal
        return str(int(float(str(val).strip())))
    except (ValueError, TypeError):
        # Non-numeric like 'virajd' → return as-is
        cleaned = str(val).strip()
        return cleaned if cleaned else "—"

# ── Build Missing Time Entries ────────────────────────────────────────────────
def build_missing_entries(df, summary, dates, emp_df_info):
    """
    Compare total logged hours vs expected hours per employee.
    Expected hours sourced from:
      1. 'All employee details' sheet if present
      2. Auto-calculated: working days * 8 hrs/day (Mon–Sat, excl Sun)
    """

    # Determine working days & default expected hours
    working_days = 0
    for d in dates:
        try:
            dt = datetime.strptime(d, "%Y-%m-%d")
            if dt.weekday() != 6:   # exclude Sunday
                working_days += 1
        except Exception:
            pass
    default_expected = working_days * 8  # 8 hrs/day standard

    # Build lookup: staff_name → {emp_id, dept, t1_manager, reporting_manager, expected}
    emp_lookup = {}

    if emp_df_info:
        emp_df, col_map, _ = emp_df_info
        for _, row in emp_df.iterrows():
            name_col = col_map.get("emp_name")
            if not name_col or pd.isna(row.get(name_col)):
                continue
            name = str(row[name_col]).strip()
            emp_lookup[name.lower()] = {
                "emp_id":   _safe_emp_id(row.get(col_map["emp_id"]) if col_map.get("emp_id") else None),
                "emp_name": name,
                "dept":     str(row[col_map["dept"]]).strip() if col_map.get("dept") and pd.notna(row.get(col_map["dept"])) else "—",
                "t1_manager": str(row[col_map["t1_manager"]]).strip() if col_map.get("t1_manager") and pd.notna(row.get(col_map["t1_manager"])) else "—",
                "reporting_manager": str(row[col_map["reporting_manager"]]).strip() if col_map.get("reporting_manager") and pd.notna(row.get(col_map["reporting_manager"])) else "—",
                "expected": r1(float(row[col_map["expected"]])) if col_map.get("expected") and pd.notna(row.get(col_map["expected"])) else default_expected,
            }

    # Build actual hours per staff
    staff_hours = {s["name"]: s["total"] for s in summary}

    # Cross-reference
    missing = []
    all_names = set(staff_hours.keys())

    # If emp_df available, use it as master list; else use timesheet staff
    if emp_lookup:
        for key, info in emp_lookup.items():
            name = info["emp_name"]
            logged = r1(staff_hours.get(name, 0))
            # Try fuzzy match if exact not found
            if name not in staff_hours:
                for sn in staff_hours:
                    if sn.lower().replace(" ", "") == name.lower().replace(" ", ""):
                        logged = r1(staff_hours[sn])
                        break
            expected = info["expected"]
            missing_hrs = r1(expected - logged)
            if missing_hrs > 0:
                missing.append({
                    "emp_id":            info["emp_id"],
                    "emp_name":          name,
                    "dept":              info["dept"],
                    "t1_manager":        info["t1_manager"],
                    "reporting_manager": info["reporting_manager"],
                    "total_logged":      logged,
                    "expected":          expected,
                    "missing":           missing_hrs,
                })
    else:
        # Fallback: use timesheet staff only, compare against default expected
        for name, logged in staff_hours.items():
            expected = default_expected
            missing_hrs = r1(expected - logged)
            if missing_hrs > 0:
                missing.append({
                    "emp_id":            "—",
                    "emp_name":          name,
                    "dept":              "—",
                    "t1_manager":        "—",
                    "reporting_manager": "—",
                    "total_logged":      logged,
                    "expected":          expected,
                    "missing":           missing_hrs,
                })

    missing.sort(key=lambda x: -x["missing"])

    # Build department filter list
    depts = sorted(set(m["dept"] for m in missing if m["dept"] != "—"))

    # Totals row
    total_logged   = r1(sum(m["total_logged"] for m in missing))
    total_expected = r1(sum(m["expected"] for m in missing))
    total_missing  = r1(sum(m["missing"] for m in missing))

    return {
        "entries":       missing,
        "count":         len(missing),
        "depts":         depts,
        "total_logged":  total_logged,
        "total_expected":total_expected,
        "total_missing": total_missing,
        "working_days":  working_days,
        "default_expected": default_expected,
    }

# ── Generate HTML ─────────────────────────────────────────────────────────────
def generate_html(data, template_path):
    with open(template_path, "r", encoding="utf-8") as f:
        template = f.read()

    replacements = {
        "{{SUMMARY_JS}}":     json.dumps(data["summary"]),
        "{{DAILY_JS}}":       json.dumps(data["daily"]),
        "{{KPI_BILL_JS}}":    json.dumps(data["kpi_billable"]),
        "{{KPI_LEAVE_JS}}":   json.dumps(data["kpi_leave"]),
        "{{KPI_NONBILL_JS}}": json.dumps(data["kpi_nonbill"]),
        "{{KPI_TOTAL_JS}}":   json.dumps(data["kpi_total"]),
        "{{MISSING_JS}}":     json.dumps(data["missing"]),
        "{{GT_BILLABLE}}":    str(data["totals"]["billable"]),
        "{{GT_LEAVE}}":       str(data["totals"]["leave"]),
        "{{GT_NONBILL}}":     str(data["totals"]["non_billable"]),
        "{{GT_TOTAL}}":       str(data["totals"]["total"]),
        "{{GT_EFFECTIVE}}":   str(data["totals"]["effective"]),
        "{{GT_BE}}":          str(data["totals"]["bill_eff"]),
        "{{GT_NE}}":          str(data["totals"]["nonbill_eff"]),
        "{{GT_TE}}":          str(data["totals"]["total_eff"]),
        "{{DATE_RANGE}}":     data["date_range"],
        "{{RECORD_COUNT}}":   str(data["record_count"]),
        "{{STAFF_COUNT}}":    str(data["staff_count"]),
        "{{LEAVE_COUNT}}":    str(data["leave_count"]),
        "{{MISSING_COUNT}}":  str(data["missing"]["count"]),
        "{{BILL_PCT}}":       str(round(data["totals"]["billable"] / data["totals"]["total"] * 100)) if data["totals"]["total"] > 0 else "0",
    }

    html = template
    for placeholder, value in replacements.items():
        html = html.replace(placeholder, value)
    return html

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    print("=" * 62)
    print("  Smart Accountants — Timesheet Dashboard Generator v2")
    print("=" * 62)
    print()

    excel_file = find_excel()
    if not excel_file:
        print("ERROR: No Excel file (.xlsx) found in this folder.")
        print(f"Place your Time & Expense export in:\n  {BASE_DIR}")
        input("\nPress Enter to exit..."); sys.exit(1)

    print(f"[1/4] Found data file: {excel_file.name}")

    print("[2/4] Parsing timesheet data...")
    try:
        df = parse_timesheet(excel_file)
    except Exception as e:
        print(f"\nERROR reading Time and Expense sheet: {e}")
        input("\nPress Enter to exit..."); sys.exit(1)

    print(f"      {len(df):,} entries · {df['staff'].nunique()} staff members")

    # Try reading employee details sheet
    emp_df_info = None
    try:
        result = parse_employee_details(excel_file)
        if result:
            emp_df_info = result
            print(f"      Employee sheet: '{result[2]}' — {len(result[0])} employees loaded")
    except Exception as e:
        print(f"      NOTE: Could not read employee details sheet ({e})")

    print("[3/4] Calculating summaries & missing time entries...")
    data = build_data(df, emp_df_info)
    m = data["missing"]
    print(f"      Date range: {data['date_range']}  ({m['working_days']} working days)")
    print(f"      Missing entries: {m['count']} employees  |  {m['total_missing']} hrs short")

    print("[4/4] Generating dashboard...")
    if not TEMPLATE.exists():
        print(f"\nERROR: Template file not found: {TEMPLATE.name}")
        input("\nPress Enter to exit..."); sys.exit(1)

    html = generate_html(data, TEMPLATE)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_name = f"SmartAccountants_Dashboard_{timestamp}.html"
    out_path = OUTPUT_DIR / out_name
    latest_path = OUTPUT_DIR / "latest_dashboard.html"

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)
    shutil.copy2(out_path, latest_path)

    print()
    print("=" * 62)
    print("  ✓  Dashboard generated successfully!")
    print("=" * 62)
    print(f"\n  Saved: output\\{out_name}")
    print(f"  Also:  output\\latest_dashboard.html")
    print("\n  Opening in your browser...")
    print()
    webbrowser.open(latest_path.as_uri())
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()
