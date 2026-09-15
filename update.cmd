@echo off
rem Copyright (C) 2026 Connor Baird
rem SPDX-License-Identifier: GPL-3.0-or-later
setlocal
title AddressMend updater
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
    echo Python 3.10 or newer was not found. Install Python, then try again.
    pause
    exit /b 1
)
rem Keep launch and exit on one parsed line: the updater may replace this file.
%PYTHON_CMD% "%~dp0update.py" --pause & exit /b
