@echo off
rem Double-click to update Roof CRM to the latest version.
rem Your data lives in Documents\RoofCRM and is never touched by this.
cd /d "%~dp0"

where git >NUL 2>&1
if errorlevel 1 (
  echo.
  echo   Git isn't installed on this computer.
  echo   Install it once from https://git-scm.com/download/win
  echo   then run this file again.
  echo.
  pause
  exit /b 1
)

if not exist ".git" (
  echo.
  echo   This folder was unzipped rather than cloned, so it can't update itself.
  echo   One-time fix - in the folder you want the app to live in, run:
  echo.
  echo     git clone https://github.com/morales451/CRM.git RoofCRM-App
  echo.
  echo   After that, double-click update.bat inside RoofCRM-App any time.
  echo   Your data in Documents\RoofCRM is unaffected either way.
  echo.
  pause
  exit /b 1
)

echo Updating Roof CRM...
git pull
if errorlevel 1 (
  echo.
  echo   Update failed - see the message above.
  echo   Your data is safe; nothing was changed.
  echo.
  pause
  exit /b 1
)

echo.
echo   Updated. Installing any new requirements...
python -m pip install -q -r requirements.txt
echo   Done - start the app with start.bat
echo.
pause
