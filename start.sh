#!/usr/bin/env bash
# One-click start for Mac/Linux. Gets the latest version first, then starts.
# Wrapped in { } so an update that rewrites this file can't derail it:
# bash reads the whole block before running it.
{
  cd "$(dirname "$0")"
  python3 updater.py --auto
  python3 -c "import flask, pandas, openpyxl, waitress, PIL" 2>/dev/null \
    || python3 -m pip install -r requirements.txt
  python3 app.py
  exit
}
