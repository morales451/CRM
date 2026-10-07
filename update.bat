@echo off
rem Double-click to update Roof CRM now. start.bat already does this every
rem time it starts, so you only need this to update a copy that's running.
rem Your data lives in Documents\RoofCRM and is never touched by this.
(
  cd /d "%~dp0"
  python updater.py
  echo.
  pause
  exit /b
)
