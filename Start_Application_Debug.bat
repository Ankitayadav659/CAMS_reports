@echo off
REM Same as Start_Application.bat but keeps a console open so any error text is visible.
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" call Start_Application.bat
".venv\Scripts\python.exe" "%~dp0main.py"
pause
