@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo The automation environment is missing. Run Setup Once.bat first.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" -c "import openpyxl, win32com.client" >nul 2>nul
if errorlevel 1 (
  echo The required packages are missing from this folder.
  echo Run Setup Once.bat as administrator, then run Prepare Drafts.bat again.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" desktop_app.py
if errorlevel 1 pause
