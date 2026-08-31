@echo off
rem Copyright (C) 2026 Connor Baird
rem SPDX-License-Identifier: GPL-3.0-or-later
setlocal
title AddressMend
cd /d "%~dp0"

where py >nul 2>nul
if not errorlevel 1 (
    py -3 "%~dp0addressmend.py"
) else (
    where python >nul 2>nul
    if errorlevel 1 (
        echo.
        echo AddressMend could not find Python.
        echo Ask your IT support to install or provide Python 3.10 or newer.
        echo Administrator access is not otherwise required by this programme.
        echo.
        pause
        exit /b 1
    )
    python "%~dp0addressmend.py"
)

if errorlevel 1 (
    echo.
    echo The programme reported a problem. Read the message above before closing.
)
echo.
pause
