@echo off
REM ===== CAMS Commission Automation System - double-click launcher =====
REM First run: creates .venv and installs libraries. Next runs: opens the app directly.
cd /d "%~dp0"

if exist ".venv\Scripts\pythonw.exe" goto launch

echo First-time setup... (internet chahiye, 1-2 minute lagenge)
where py >nul 2>nul
if %errorlevel%==0 ( py -3 -m venv .venv ) else ( python -m venv .venv )
if not exist ".venv\Scripts\python.exe" (
  echo.
  echo Python 3 nahi mila. python.org se 64-bit Python install karein ^(Add python.exe to PATH tick karein^).
  pause
  exit /b 1
)
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 (
  echo.
  echo Libraries install nahi ho payi. Internet check karke dobara chalayein.
  pause
  exit /b 1
)

:launch
start "" ".venv\Scripts\pythonw.exe" "%~dp0main.py"
