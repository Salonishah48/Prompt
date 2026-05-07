@echo off
setlocal
set "SD=%~dp0"
set "SD=%SD:~0,-1%"
echo %SD% | findstr /r "^\\\\\\\\" >nul
if %errorlevel%==0 (
    echo Mapping network path to Z: ...
    net use Z: /delete >nul 2>&1
    net use Z: "%SD%" >nul 2>&1
    if errorlevel 1 (
        echo ERROR: Could not map network drive.
        echo Please copy the WIP_App folder to a local drive e.g. C:\WIP_App
        pause & exit /b 1
    )
    Z: & cd \
) else (
    cd /d "%SD%"
)
python app.py
if errorlevel 1 (
    py app.py
    if errorlevel 1 (
        echo.
        echo ERROR: Could not launch. Make sure Python is installed and run install_once.bat first.
        pause
    )
)
net use Z: /delete >nul 2>&1
endlocal
