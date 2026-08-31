#!/bin/sh
# Copyright (C) 2026 Connor Baird
# SPDX-License-Identifier: GPL-3.0-or-later

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd) || exit 1
cd "$SCRIPT_DIR" || exit 1

if command -v python3 >/dev/null 2>&1; then
    exec python3 addressmend.py
fi

printf '%s\n' 'Python 3.10 or newer could not be found.'
printf '%s\n' 'Ask IT support to make Python 3 available; administrator rights are not needed to run this programme.'
printf '%s' 'Press Enter to close... '
read answer
exit 1
