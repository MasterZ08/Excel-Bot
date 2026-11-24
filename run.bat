@echo off
setlocal
cd /d "%~dp0"
title Random Excel Data Generator

set "PY=python"
where py >nul 2>nul && set "PY=py -3"
%PY% --version >nul 2>nul
if errorlevel 1 goto nopython

if exist ".venv\Scripts\python.exe" goto havevenv
echo Creating virtual environment...
%PY% -m venv .venv
if errorlevel 1 goto fail
:havevenv

".venv\Scripts\python.exe" -c "import customtkinter, openpyxl, faker" >nul 2>nul
if not errorlevel 1 goto launch
echo Installing dependencies - first run only, needs internet...
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto fail

:launch
".venv\Scripts\python.exe" -c "import gui" >nul 2>"%TEMP%\reg_err.txt"
if errorlevel 1 goto starterr
start "" ".venv\Scripts\pythonw.exe" main.py
exit /b 0

:starterr
echo The app failed to start:
type "%TEMP%\reg_err.txt"
pause
exit /b 1

:nopython
echo Python 3.11 or newer was not found.
echo Install it from https://www.python.org/downloads/ and tick "Add python.exe to PATH".
pause
exit /b 1

:fail
echo Setup failed - see the messages above. Check your internet connection.
echo If it keeps failing, delete the .venv folder and run this file again.
pause
exit /b 1
