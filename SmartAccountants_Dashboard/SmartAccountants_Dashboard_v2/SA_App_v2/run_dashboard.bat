@echo off
title Smart Accountants - Timesheet Dashboard
color 1F
echo.
echo  ============================================================
echo    Smart Accountants - Timesheet Dashboard Generator v2
echo  ============================================================
echo.
echo  Place your Time and Expense Excel file in this folder,
echo  then press any key to generate the dashboard.
echo.
echo  TIP: Make sure your Excel file includes the sheet named
echo       "List of employees - SA" to enable full Missing Time
echo       Entries with Emp ID, Department, and Manager details.
echo.
pause

cd /d "%~dp0"
python app.py

if %errorlevel% neq 0 (
    echo.
    echo  Something went wrong. Make sure you ran install_once.bat first.
    echo.
    pause
)
