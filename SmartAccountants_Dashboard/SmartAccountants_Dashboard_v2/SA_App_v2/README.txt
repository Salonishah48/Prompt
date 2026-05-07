============================================================
  Smart Accountants — Timesheet Dashboard v2
  Local Application
============================================================

FIRST TIME SETUP (do this once):
  1. Double-click  install_once.bat
  2. Wait for it to finish
  3. Done!

EVERY WEEK:
  1. Place your Time & Expense Excel export in this folder
  2. Double-click  run_dashboard.bat
  3. Dashboard opens automatically in your browser

============================================================
  MISSING TIME ENTRIES TAB
============================================================

The dashboard includes a "Missing Time Entries" tab that
shows employees who logged fewer hours than expected.

HOW IT WORKS WITHOUT EMPLOYEE DETAILS SHEET:
  The app auto-detects working days from the date range
  (Mon–Sat, excluding Sundays) and uses 9 hrs/day as the
  expected standard. Missing entries are shown but without
  Emp ID, Department, or Manager details.

HOW TO ENABLE FULL MISSING ENTRIES (with Dept & Manager):
  Add a sheet named "List of employees - SA" to your Excel
  file with these columns:

  | Emp ID | Employee Name | Department | T1 Manager | Reporting Manager | Expected Hours |
  |--------|--------------|------------|------------|-------------------|----------------|
  | 1154   | John Smith   | US Tax     | 224 - Ritesh | Devendrakumar G. | 51             |

  Column names are flexible — the app recognises variations:
    Emp ID:            "Emp ID", "Employee ID", "ID"
    Employee Name:     "Employee Name", "Name", "Staff Name"
    Department:        "Department", "Dept"
    T1 Manager:        Any column with "T1" and "Manager"
    Reporting Manager: Any column with "Reporting" and "Manager"
    Expected Hours:    "Expected Hours", "Expected Hrs", "Expected"

  Expected Hours = the target hours for that employee for
  the reporting period. If blank, defaults to working days × 9.

DEPARTMENT COLOUR CODING:
  Marketing    → Orange
  US Accounts  → Blue
  US Audit     → Green
  US Tax       → Purple
  Operations   → Yellow
  Others       → Grey

============================================================
  FILE STRUCTURE
============================================================

  install_once.bat         ← Run ONCE to install requirements
  run_dashboard.bat        ← Run EVERY WEEK
  app.py                   ← Processing engine (do not edit)
  dashboard_template.html  ← Dashboard design template
  output\                  ← Generated dashboards saved here
    latest_dashboard.html     ← Always the most recent
    SmartAccountants_Dashboard_YYYYMMDD_HHMMSS.html

REQUIREMENTS:
  Windows PC · Python 3.8+ · Internet for fonts on first view

============================================================
  Smart Accountants | Confidential Internal Use
============================================================
