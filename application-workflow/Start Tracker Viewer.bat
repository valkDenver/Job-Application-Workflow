@echo off
setlocal EnableDelayedExpansion
cd /d "%~dp0"
set "TRK=%~1"
if not defined TRK if exist "tracker_path.txt" set /p TRK=<"tracker_path.txt"
if defined TRK goto have
echo Drag your Master Job Tracker .xlsx file onto this window, then press Enter.
set /p "TRK=File: "
:have
rem Clean what a drag or paste gives: "quotes", 'quotes' (PowerShell windows), and leading or trailing spaces
set "TRK=!TRK:"=!"
:trim
if "!TRK:~0,1!"==" " (set "TRK=!TRK:~1!" & goto trim)
if "!TRK:~-1!"==" " (set "TRK=!TRK:~0,-1!" & goto trim)
if "!TRK:~0,1!"=="'" if "!TRK:~-1!"=="'" (set "TRK=!TRK:~1,-1!" & goto trim)
if exist "!TRK!" goto run
echo Could not find: !TRK!
pause
exit /b 1
:run
>"tracker_path.txt" echo(!TRK!
python "Tools\tracker_server.py" "!TRK!"
if not errorlevel 1 goto end
echo.
echo The helper did not start. If Python was not found, install it from python.org (tick "Add to PATH"),
echo then run once:  pip install openpyxl
pause
:end
