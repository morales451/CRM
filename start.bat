@echo off
rem One-click start for Windows. First run installs dependencies.
cd /d "%~dp0"
python -c "import flask, pandas, openpyxl, waitress" 2>NUL
if errorlevel 1 python -m pip install -r requirements.txt
python app.py
pause
