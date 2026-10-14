@echo off
REM ===== Desktop par "CAMS Commission Automation" shortcut banata hai (app ke apne icon ke saath) =====
cd /d "%~dp0"

if exist ".venv\Scripts\pythonw.exe" goto make
echo Pehle ek baar Start_Application.bat chalayein (woh libraries install karega), phir yeh file dobara chalayein.
pause
exit /b 1

:make
".venv\Scripts\python.exe" -c "import app_icon; app_icon.ensure_ico()"
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$d = (Get-Location).Path;" ^
  "$s = (New-Object -ComObject WScript.Shell).CreateShortcut([Environment]::GetFolderPath('Desktop') + '\CAMS Commission Automation.lnk');" ^
  "$s.TargetPath = Join-Path $d '.venv\Scripts\pythonw.exe';" ^
  "$s.Arguments = '\"' + (Join-Path $d 'main.py') + '\"';" ^
  "$s.WorkingDirectory = $d;" ^
  "$s.IconLocation = (Join-Path $d 'assets\app.ico') + ',0';" ^
  "$s.Description = 'CAMS Commission Automation System';" ^
  "$s.Save()"
if errorlevel 1 (
  echo Shortcut nahi ban paya.
  pause
  exit /b 1
)
echo Desktop par "CAMS Commission Automation" shortcut ban gaya.
timeout /t 3 >nul
