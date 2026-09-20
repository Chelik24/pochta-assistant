@echo off
rem Only ASCII in this file: cmd.exe cannot read UTF-8 batch files.
rem
rem Finds a working Python and runs it with the given arguments.
rem Needed because "python" on Windows often points to the Microsoft Store
rem stub, which prints "Python" and does nothing. This looks for the real
rem interpreter directly instead of trusting PATH.
rem
rem Usage:
rem   run.cmd fetch_mail.py --days 3
rem   run.cmd -m pytest tests -q

setlocal enabledelayedexpansion
cd /d "%~dp0"

set "PYEXE="

rem 1. Per-user install (what python.org and winget do by default)
for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do (
    if exist "%%~fD\python.exe" set "PYEXE=%%~fD\python.exe"
)

rem 2. All-users install
if not defined PYEXE (
    for /d %%D in ("%ProgramFiles%\Python3*") do (
        if exist "%%~fD\python.exe" set "PYEXE=%%~fD\python.exe"
    )
)

rem 3. Official Windows launcher
if not defined PYEXE (
    py -c "import sys" >nul 2>&1 && set "PYEXE=py"
)

rem 4. Whatever PATH offers, as a last resort
if not defined PYEXE (
    python -c "import sys" >nul 2>&1 && set "PYEXE=python"
)

if not defined PYEXE (
    echo.
    echo Python not found.
    echo Install it from python.org and tick "Add python.exe to PATH".
    echo.
    exit /b 9
)

"%PYEXE%" %*
