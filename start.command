#!/bin/sh
# Copyright (C) 2026 Connor Baird
# SPDX-License-Identifier: GPL-3.0-or-later

SCRIPT_DIR=$(CDPATH= cd -- "$(dirname -- "$0")" && pwd) || exit 1
cd "$SCRIPT_DIR" || exit 1

# Finder-launched Terminal windows may not inherit Homebrew or python.org paths.
PATH="/opt/homebrew/bin:/usr/local/bin:/Library/Frameworks/Python.framework/Versions/Current/bin:$PATH"
export PATH

status=0
if command -v python3 >/dev/null 2>&1 &&
   python3 -c 'import sys; raise SystemExit(0 if sys.version_info >= (3, 10) else 1)'
then
    python3 addressmend.py || status=$?
else
    printf '%s\n' 'AddressMend could not find Python 3.10 or newer.'
    printf '%s\n' 'Install Python from https://www.python.org/downloads/macos/ and try again.'
    status=1
fi

printf '\n%s' 'Press Return to close... '
read answer
exit "$status"
