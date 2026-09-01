@echo off
rem Copyright (C) 2026 Connor Baird
rem SPDX-License-Identifier: GPL-3.0-or-later
setlocal
title AddressMend
cd /d "%~dp0"

set "PYTHON_CMD="
where py >nul 2>nul
if not errorlevel 1 (
    py -3 -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>nul
    if not errorlevel 1 set "PYTHON_CMD=py -3"
)

if not defined PYTHON_CMD (
    where python >nul 2>nul
    if not errorlevel 1 (
        python -c "import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)" >nul 2>nul
        if not errorlevel 1 set "PYTHON_CMD=python"
    )
)

if not defined PYTHON_CMD (
    echo.
    echo AddressMend could not find Python 3.10 or newer.
    echo Ask your IT support to install or provide a suitable Python version.
    echo Administrator access is not otherwise required by this programme.
    echo.
    pause
    exit /b 1
)

%PYTHON_CMD% "%~dp0addressmend.py"
set "ADDRESSMEND_EXIT=%ERRORLEVEL%"
if not "%ADDRESSMEND_EXIT%"=="0" (
    echo.
    echo The programme reported a problem. Read the message above before closing.
    echo.
    pause
)
exit /b %ADDRESSMEND_EXIT%
