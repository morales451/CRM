#!/usr/bin/env bash
# One-click start for Mac/Linux. First run installs dependencies.
cd "$(dirname "$0")"
python3 -c "import flask, pandas, openpyxl, waitress" 2>/dev/null \
  || python3 -m pip install -r requirements.txt
python3 app.py
