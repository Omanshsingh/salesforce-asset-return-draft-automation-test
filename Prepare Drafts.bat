@echo off
setlocal
set "BASE=%~dp0"
set "INPUT=%BASE%Input"
set "OUTPUT=%BASE%Output"
set "FILE="

for /f "delims=" %%F in ('powershell -NoProfile -ExecutionPolicy Bypass -Command "Get-ChildItem -LiteralPath '%INPUT%' -Filter '*.xlsx' -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1 -ExpandProperty FullName"') do set "FILE=%%F"

if not defined FILE (
  echo No Excel file found.
  echo Copy today's Salesforce Excel file into the Input folder, then run this again.
  pause
  exit /b 1
)

echo Preparing unsent email files from:
echo %FILE%
python "%BASE%desktop_draft_automation.py" "%FILE%" --output "%OUTPUT%" --create-eml
echo.
echo Finished. Open the Output\eml folder.
pause
