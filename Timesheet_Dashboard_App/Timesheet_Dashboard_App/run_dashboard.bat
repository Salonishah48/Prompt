@echo off
echo ============================================
echo  Infinity Globus - Timesheet Dashboard
echo ============================================
echo.
echo Starting the app...
echo.
python app.py
echo.
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Something went wrong. Make sure:
    echo   1. Python is installed and in PATH
    echo   2. You ran install_once.bat first
    echo   3. Your Excel file is placed in this folder
    echo.
)
pause
