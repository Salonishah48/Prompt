"""
Smart Accountants — Timesheet Dashboard Generator
Run via: run_dashboard.bat
"""

import os, sys, json, math, glob, webbrowser, shutil
from datetime import datetime
from pathlib import Path

# ── Dependency check ─────────────────────────────────────────────────────────
try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas not installed. Please run install_once.bat first.")
    input("Press Enter to exit...")
    sys.exit(1)

try:
    import openpyxl
except ImportError:
    print("ERROR: openpyxl not installed. Please run install_once.bat first.")
    input("Press Enter to exit...")
    sys.exit(1)

BASE_DIR    = Path(__file__).parent
OUTPUT_DIR  = BASE_DIR / "output"
TEMPLATE    = BASE_DIR / "dashboard_template.html"

OUTPUT_DIR.mkdir(exist_ok=True)

# ── Find Excel file ───────────────────────────────────────────────────────────
def find_excel():
    patterns = ["*.xlsx", "*.xls"]
    candidates = []
    for pat in patterns:
        candidates.extend(BASE_DIR.glob(pat))
    # Exclude temp files
    candidates = [f for f in candidates if not f.name.startswith("~$")]
    if not candidates:
        return None
    # Prefer most recently modified
    candidates.sort(key=lambda f: f.stat().st_mtime, reverse=True)
    return candidates[0]

# ── Rounding helpers ──────────────────────────────────────────────────────────
def r1(v):
    """Round to 1 decimal."""
    return round(float(v), 1)

def r0(v):
    """Round to nearest integer."""
    return round(float(v))

# ── Parse Excel ───────────────────────────────────────────────────────────────
def parse_excel(path):
    print(f"  Reading: {path.name}")
    df_raw = pd.read_excel(path, sheet_name="Time and Expense", header=None)

    records = []
    current_staff = None
    rows = list(df_raw.iterrows())

    i = 0
    while i < len(rows):
        idx, row = rows[i]
        val1 = row[1]
        val2 = str(row[2]) if pd.notna(row[2]) else ""

        if isinstance(val1, str) and "Staff :" in val1:
            current_staff = val1.replace("Staff :", "").strip()
            i += 1
            continue

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
            if i + 1 < len(rows):
                r1v = str(rows[i + 1][1][2]) if pd.notna(rows[i + 1][1][2]) else ""
                if "Invoice Description:" in r1v:
                    rec["inv_desc"] = r1v.replace("Invoice Description:", "").strip()
            if i + 2 < len(rows):
                r2v = str(rows[i + 2][1][2]) if pd.notna(rows[i + 2][1][2]) else ""
                if "Internal Note:" in r2v:
                    rec["int_note"] = r2v.replace("Internal Note:", "").strip()
            records.append(rec)
        i += 1

    return pd.DataFrame(records)

# ── Categorise hours ──────────────────────────────────────────────────────────
LEAVE_KEYWORDS = ["sick", "personal", "holiday", "vacation"]

def categorise(row):
    sl = row["service"].lower()
    for k in LEAVE_KEYWORDS:
        if k in sl:
            return "leave"
    if "Non Billable" in row["client"] or "NONB" in row["client"]:
        return "non_billable"
    return "billable"

# ── Build data ────────────────────────────────────────────────────────────────
def build_data(df):
    df["category"] = df.apply(categorise, axis=1)

    # Summary per staff
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

    # Daily detail per employee (with notes)
    daily = {}
    for staff, g in df.groupby("staff"):
        nm = staff.split("(")[0].strip()
        days = {}
        for date, dg in g.groupby("date"):
            entries = []
            for _, row in dg.iterrows():
                cl = row["client"].split("(")[0].strip()
                if len(cl) > 40: cl = cl[:40] + "…"
                svc = row["service"]
                if "(" in svc: svc = svc[:svc.rfind("(")].strip()
                entries.append({
                    "client":   cl,
                    "service":  svc,
                    "hours":    r1(row["hours"]),
                    "category": row["category"],
                    "status":   row["status"],
                    "inv_desc": row["inv_desc"][:120] if row["inv_desc"] else "",
                    "int_note": row["int_note"][:200] if row["int_note"] else "",
                })
            days[date] = entries
        daily[nm] = days

    # Daily totals (for reference, not shown on landing page)
    daily_tots = {}
    for d, g in df.groupby("date"):
        daily_tots[d] = {
            "hours":        r1(g["hours"].sum()),
            "billable":     r1(g[g.category == "billable"]["hours"].sum()),
            "non_billable": r1(g[g.category == "non_billable"]["hours"].sum()),
            "leave":        r1(g[g.category == "leave"]["hours"].sum()),
        }

    # KPI drill-down: billable
    kpi_billable = []
    for staff, g in df[df.category == "billable"].groupby("staff"):
        nm = staff.split("(")[0].strip()
        by_day = {d: r1(h) for d, h in g.groupby("date")["hours"].sum().items()}
        kpi_billable.append({"name": nm, "total": r1(g["hours"].sum()), "by_day": by_day})
    kpi_billable.sort(key=lambda x: -x["total"])

    # KPI drill-down: leave
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

    # KPI drill-down: non-billable
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

    # KPI drill-down: total
    kpi_total = [{"name": s["name"], "total": s["total"],
                  "billable": s["billable"], "leave": s["leave"],
                  "non_billable": s["non_billable"]} for s in summary]

    # Grand totals
    t_bill  = r1(df[df.category == "billable"]["hours"].sum())
    t_leave = r1(df[df.category == "leave"]["hours"].sum())
    t_nonb  = r1(df[df.category == "non_billable"]["hours"].sum())
    t_total = r1(df["hours"].sum())
    t_eff   = r1(t_total - t_leave)
    gt_be   = r0(t_bill / t_eff * 100) if t_eff > 0 else 0
    gt_ne   = r0(t_nonb / t_eff * 100) if t_eff > 0 else 0

    totals = {
        "total": t_total, "billable": t_bill, "leave": t_leave,
        "non_billable": t_nonb, "effective": t_eff,
        "bill_eff": gt_be, "nonbill_eff": gt_ne, "total_eff": gt_be + gt_ne,
    }

    # Date range label
    dates = sorted(df["date"].unique())
    date_range = f"{dates[0]} to {dates[-1]}" if dates else "Unknown"
    try:
        d1 = datetime.strptime(dates[0], "%Y-%m-%d")
        d2 = datetime.strptime(dates[-1], "%Y-%m-%d")
        date_range = f"{d1.strftime('%b %d')} – {d2.strftime('%b %d, %Y')}"
    except Exception:
        pass

    return {
        "summary":      summary,
        "daily":        daily,
        "daily_tots":   daily_tots,
        "kpi_billable": kpi_billable,
        "kpi_leave":    kpi_leave,
        "kpi_nonbill":  kpi_nonbill,
        "kpi_total":    kpi_total,
        "totals":       totals,
        "date_range":   date_range,
        "record_count": len(df),
        "staff_count":  len(summary),
        "leave_count":  len(kpi_leave),
    }

# ── Generate HTML ─────────────────────────────────────────────────────────────
def generate_html(data, template_path):
    with open(template_path, "r", encoding="utf-8") as f:
        template = f.read()

    # Replace all placeholders
    replacements = {
        "{{SUMMARY_JS}}":     json.dumps(data["summary"]),
        "{{DAILY_JS}}":       json.dumps(data["daily"]),
        "{{DAILY_TOTS_JS}}":  json.dumps(data["daily_tots"]),
        "{{KPI_BILL_JS}}":    json.dumps(data["kpi_billable"]),
        "{{KPI_LEAVE_JS}}":   json.dumps(data["kpi_leave"]),
        "{{KPI_NONBILL_JS}}": json.dumps(data["kpi_nonbill"]),
        "{{KPI_TOTAL_JS}}":   json.dumps(data["kpi_total"]),
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
        "{{BILL_PCT}}":       str(round(data["totals"]["billable"] / data["totals"]["total"] * 100)) if data["totals"]["total"] > 0 else "0",
    }

    html = template
    for placeholder, value in replacements.items():
        html = html.replace(placeholder, value)

    return html

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    print("=" * 60)
    print("  Smart Accountants — Timesheet Dashboard Generator")
    print("=" * 60)
    print()

    # Find Excel file
    excel_file = find_excel()
    if not excel_file:
        print("ERROR: No Excel file (.xlsx) found in this folder.")
        print()
        print("Please place your Time & Expense Excel export in:")
        print(f"  {BASE_DIR}")
        print()
        input("Press Enter to exit...")
        sys.exit(1)

    print(f"[1/4] Found data file: {excel_file.name}")

    # Parse
    print("[2/4] Parsing timesheet data...")
    try:
        df = parse_excel(excel_file)
    except Exception as e:
        print(f"\nERROR reading Excel file: {e}")
        print("Make sure the file is a valid Time & Expense export.")
        input("\nPress Enter to exit...")
        sys.exit(1)

    print(f"      {len(df):,} entries found across {df['staff'].nunique()} staff members")

    # Build data
    print("[3/4] Processing and calculating summaries...")
    data = build_data(df)
    print(f"      Date range: {data['date_range']}")
    print(f"      Total hours: {data['totals']['total']}  |  Billable: {data['totals']['billable']}  |  Leave: {data['totals']['leave']}")

    # Generate HTML
    print("[4/4] Generating dashboard...")
    if not TEMPLATE.exists():
        print(f"\nERROR: Template file not found: {TEMPLATE.name}")
        print("Make sure dashboard_template.html is in the same folder as app.py")
        input("\nPress Enter to exit...")
        sys.exit(1)

    html = generate_html(data, TEMPLATE)

    # Save output
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_name = f"SmartAccountants_Dashboard_{timestamp}.html"
    out_path = OUTPUT_DIR / out_name

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

    # Also save as "latest" for quick access
    latest_path = OUTPUT_DIR / "latest_dashboard.html"
    shutil.copy2(out_path, latest_path)

    print()
    print("=" * 60)
    print("  ✓  Dashboard generated successfully!")
    print("=" * 60)
    print()
    print(f"  File: output\\{out_name}")
    print(f"  Also saved as: output\\latest_dashboard.html")
    print()
    print("  Opening in your browser...")
    print()

    # Open in browser
    webbrowser.open(latest_path.as_uri())

    input("Press Enter to exit...")

if __name__ == "__main__":
    main()
