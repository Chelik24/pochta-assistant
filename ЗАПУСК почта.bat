@echo off
rem Only ASCII in this file: cmd.exe cannot read UTF-8 batch files.
rem All Russian messages are printed by fetch_mail.py.
chcp 65001 > nul
cd /d "%~dp0"

where python > nul 2>&1
if errorlevel 1 (
    echo.
    echo Python not found. Install Python 3.10+ from python.org
    echo and tick "Add python.exe to PATH" during setup.
    echo.
    pause
    exit /b 1
)

python fetch_mail.py --days 3 --open
if errorlevel 1 (
    echo.
    pause
    exit /b 1
)
