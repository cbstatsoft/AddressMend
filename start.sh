#!/bin/sh
# Copyright (C) 2026 Connor Baird
# SPDX-License-Identifier: GPL-3.0-or-later

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd) || exit 1
cd "$SCRIPT_DIR" || exit 1

if command -v python3 >/dev/null 2>&1 &&
   python3 -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)'
then
    exec python3 addressmend.py
fi

printf '%s\n' 'Python 3.10 or newer could not be found.'
printf '%s\n' 'Install Python 3.10 or newer, or ask IT support to make it available.'
printf '%s' 'Press Enter to close... '
read answer
exit 1
