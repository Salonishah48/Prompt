@echo off
setlocal
set "SD=%~dp0"
set "SD=%SD:~0,-1%"
echo %SD% | findstr /r "^\\\\\\\\" >nul
if %errorlevel%==0 (
    net use Z: /delete >nul 2>&1
    net use Z: "%SD%" >nul 2>&1
    Z: & cd \
) else (
    cd /d "%SD%"
)
echo ============================================
echo  WIP Dashboard Generator - First Time Setup
echo ============================================
echo.
echo Installing Python packages...
pip install pandas openpyxl xlrd
echo.
echo Done! Run run_dashboard.bat to launch.
net use Z: /delete >nul 2>&1
endlocal
pause
