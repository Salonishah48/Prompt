@echo off
title Smart Accountants - Timesheet Dashboard
color 1F
echo.
echo  ============================================================
echo    Smart Accountants - Timesheet Dashboard Generator
echo  ============================================================
echo.
echo  Place your Time and Expense Excel file in this folder,
echo  then press any key to generate the dashboard.
echo.
echo  The dashboard will open automatically in your browser.
echo.
pause

cd /d "%~dp0"
python app.py

if %errorlevel% neq 0 (
    echo.
    echo  Something went wrong. Make sure you have run install_once.bat
    echo.
    pause
)
