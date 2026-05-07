============================================================
  Smart Accountants — Timesheet Dashboard
  Local Application
============================================================

FIRST TIME SETUP (do this once):
  1. Double-click  install_once.bat
  2. Wait for it to finish — it installs Python libraries
  3. You're done!

EVERY WEEK:
  1. Copy your Time & Expense Excel export into this folder
     (same folder as run_dashboard.bat)
  2. Double-click  run_dashboard.bat
  3. The dashboard opens automatically in your browser
  4. The generated file is saved in the  output\  folder

FILE STRUCTURE:
  install_once.bat      — Run once to install requirements
  run_dashboard.bat     — Run every week to generate dashboard
  app.py                — The processing engine (do not edit)
  dashboard_template.html — The dashboard design template
  output\               — Generated dashboards saved here
    latest_dashboard.html  — Always the most recent dashboard
    SmartAccountants_Dashboard_YYYYMMDD_HHMMSS.html — Dated copies

REQUIREMENTS:
  - Windows PC
  - Python 3.8 or later  (download from python.org)
  - Internet connection on first run (to load fonts)

TIPS:
  - You can keep multiple weeks' Excel files in the folder —
    the app always picks the most recently modified one
  - Each run creates a new timestamped file in output\
    so you never lose a previous dashboard
  - The latest_dashboard.html is always overwritten with
    the most recent run

TROUBLESHOOTING:
  - "Python not found"   → Install Python from python.org,
                           check "Add Python to PATH"
  - "No Excel file found" → Make sure your .xlsx file is in
                            the same folder as run_dashboard.bat
  - "Error reading Excel" → Ensure the sheet is named
                            "Time and Expense"

============================================================
  Smart Accountants | Confidential Internal Use
============================================================
