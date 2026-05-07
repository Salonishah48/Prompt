@echo off
title Smart Accountants Dashboard - First Time Setup
color 1F
echo.
echo  ============================================================
echo    Smart Accountants - Timesheet Dashboard
echo    First Time Setup
echo  ============================================================
echo.
echo  This will install the required Python libraries.
echo  This only needs to be done ONCE.
echo.
echo  Checking Python installation...
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  ERROR: Python is not installed or not in PATH.
    echo.
    echo  Please install Python from https://www.python.org/downloads/
    echo  During installation, check "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

python --version
echo.
echo  Installing required libraries...
echo.

pip install pandas openpyxl --quiet --upgrade

if %errorlevel% neq 0 (
    echo.
    echo  ERROR: Installation failed.
    echo  Try running this file as Administrator (right-click - Run as administrator)
    echo.
    pause
    exit /b 1
)

echo.
echo  ============================================================
echo    Setup complete! You can now use run_dashboard.bat
echo  ============================================================
echo.
pause
