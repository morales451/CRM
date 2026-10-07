@echo off
rem One-click start for Windows. Gets the latest version first, then starts.
rem Your data lives in Documents\RoofCRM and is never touched by an update.
rem
rem Everything is inside one ( ) block on purpose: cmd reads a .bat file
rem line by line while it runs, so an update that rewrites this file would
rem otherwise derail it. A block is read in full before it starts.
(
  cd /d "%~dp0"
  python updater.py --auto
  python -c "import flask, pandas, openpyxl, waitress, PIL" 2>NUL
  if errorlevel 1 python -m pip install -r requirements.txt
  python app.py
  pause
  exit /b
)
