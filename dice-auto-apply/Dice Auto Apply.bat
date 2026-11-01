@echo off
rem Windows: double-click. Checks the setup, then asks for the line your AI assistant gave you (or press Enter for a default round).
rem The line is saved to a file and checked by dice_apply.py bridge; it never runs through this command window.
rem PYCMD keeps its own quotes, so a user folder with a space (C:\Users\Jane Doe) works.
setlocal EnableDelayedExpansion
cd /d "%~dp0"
call :findpy
!PYCMD! "Tools\dice_apply.py" doctor
if errorlevel 2 goto nopython
if errorlevel 1 (
  echo.
  set /p "FIX=Some things need fixing. Fix them now? [y/N] "
  if /i "!FIX!"=="y" !PYCMD! "Tools\dice_apply.py" doctor --fix
)
call :findpy
echo.
echo Paste the line from your assistant, then press Enter.
echo (Press Enter on its own for a round with today's default settings.)
set "LINE="
set /p "LINE=> "
>".pasted_line.txt" echo(!LINE!
!PYCMD! "Tools\dice_apply.py" bridge --line-file ".pasted_line.txt"
del ".pasted_line.txt" >nul 2>nul
echo.
pause
goto :eof

:findpy
set "VPY=%USERPROFILE%\.job-workflow\venv\Scripts\python.exe"
if exist "%VPY%" (
  set PYCMD="%VPY%"
  goto :eof
)
where py >nul 2>nul
if not errorlevel 1 (
  set "PYCMD=py -3"
  goto :eof
)
set "PYCMD=python"
goto :eof

:nopython
echo.
echo Python did not start. Install Python 3 from python.org (tick "Add to PATH"), then run this again.
pause
