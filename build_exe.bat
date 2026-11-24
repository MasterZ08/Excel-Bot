@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (py -3 -m venv .venv || python -m venv .venv)
".venv\Scripts\python.exe" -m pip install -r requirements-build.txt || (pause & exit /b 1)
".venv\Scripts\python.exe" -m PyInstaller --noconfirm --clean --onefile --windowed --name RandomExcelDataGenerator --collect-all customtkinter --collect-all faker main.py || (pause & exit /b 1)
echo.
echo Built: dist\RandomExcelDataGenerator.exe  (output\, configs\ and logs\ are created next to the EXE)
pause
