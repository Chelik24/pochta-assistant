@echo off
rem Only ASCII in this file: cmd.exe cannot read UTF-8 batch files.
rem All Russian messages are printed by mail.py.
rem No labels and no multi-line blocks on purpose: those are the parts
rem that break when the file loses its Windows line endings.
chcp 65001 > nul
cd /d "%~dp0"

call "%~dp0run.cmd" mail.py sync
echo.
call "%~dp0run.cmd" mail.py summary --days 3
echo.
pause
