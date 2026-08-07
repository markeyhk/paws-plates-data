#!/bin/bash
# Double-click this to check the FEHD list right now.
# First run sets up a small private Python environment; later runs are quick.
set -u

REPO="$(cd "$(dirname "$0")" && pwd)"
VENV="$HOME/.paws-plates-venv"

popup() {
  osascript -e 'on run argv' \
            -e 'display dialog (item 1 of argv) with title "Paws & Plates" buttons {"OK"} default button "OK" with icon stop' \
            -e 'end run' -- "$1" >/dev/null 2>&1
}

if [ ! -x "$VENV/bin/python" ]; then
  echo "First-time setup — installing a private Python environment…"
  /usr/bin/python3 -m venv "$VENV" \
    || { popup "Could not create the Python environment.

Tell Claude: venv creation failed."; exit 1; }
  "$VENV/bin/pip" install --quiet --upgrade pip
  "$VENV/bin/pip" install --quiet openpyxl requests pypdf \
    || { popup "Could not install the Python packages.

Check your internet connection and try again."; exit 1; }
  echo "Setup done."
fi

exec "$VENV/bin/python" "$REPO/scripts/mac_check.py"
