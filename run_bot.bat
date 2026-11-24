@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" py -3 -m venv .venv
".venv\Scripts\python.exe" -m pip install -r requirements.txt -r requirements-bot.txt
if errorlevel 1 goto fail
start "" ".venv\Scripts\pythonw.exe" typing_bot.py
exit /b 0
:fail
echo Install failed - check your internet connection.
pause
