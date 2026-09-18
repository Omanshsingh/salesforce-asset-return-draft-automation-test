@echo off
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Please run Setup Once.bat first.
  pause
  exit /b 1
)
".venv\Scripts\python.exe" desktop_app.py
if errorlevel 1 pause
