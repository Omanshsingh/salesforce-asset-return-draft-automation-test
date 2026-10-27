@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoLogo -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%~dp0Setup Once.ps1"
set "SETUP_EXIT=%ERRORLEVEL%"
if not "%SETUP_EXIT%"=="0" (
  echo.
  echo Setup failed. Read the message above, fix the issue, then run Setup Once.bat again.
)
pause
exit /b %SETUP_EXIT%
