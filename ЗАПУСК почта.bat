@echo off
rem Only ASCII in this file: cmd.exe cannot read UTF-8 batch files.
rem All Russian messages are printed by fetch_mail.py.
chcp 65001 > nul
cd /d "%~dp0"

call run.cmd fetch_mail.py --days 3 --open
if errorlevel 1 (
    echo.
    pause
    exit /b 1
)
