@echo off
setlocal
cd /d "%~dp0"
".venv\Scripts\python.exe" -c "import sys, openpyxl, win32com.client; raise SystemExit(0 if sys.version_info >= (3, 11) else 1)" >nul 2>nul
if errorlevel 1 (
  echo The app environment is missing or incomplete. Running setup now...
  call "%~dp0Setup Once.bat"
  if errorlevel 1 exit /b 1
)
".venv\Scripts\python.exe" desktop_app.py
if errorlevel 1 pause
