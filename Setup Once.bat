@echo off
setlocal
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" goto packages
where py >nul 2>nul
if not errorlevel 1 (
  py -3 -m venv .venv
) else (
  python -m venv .venv
)
if errorlevel 1 goto failed
:packages
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto failed
echo Setup complete. Double-click Prepare Drafts.bat.
pause
exit /b 0
:failed
echo Setup failed. Install Python 3.11 or newer from python.org, then retry.
echo If installation is blocked, ask your IT person for help.
pause
exit /b 1
