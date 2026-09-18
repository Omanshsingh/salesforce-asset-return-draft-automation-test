@echo off
setlocal
cd /d "%~dp0"

set "PYTHON_CMD="
where py >nul 2>nul
if not errorlevel 1 set "PYTHON_CMD=py -3"
if not defined PYTHON_CMD (
  where python >nul 2>nul
  if not errorlevel 1 set "PYTHON_CMD=python"
)

if not defined PYTHON_CMD goto no_python

%PYTHON_CMD% -c "import sys; print('Found Python ' + sys.version.split()[0]); raise SystemExit(0 if sys.version_info >= (3, 11) else 1)"
if errorlevel 1 goto old_python

if exist ".venv\Scripts\python.exe" goto packages

%PYTHON_CMD% -m venv .venv
if errorlevel 1 goto venv_failed

:packages
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto packages_failed
echo Setup complete. Double-click Prepare Drafts.bat.
pause
exit /b 0

:no_python
echo Python was not found in PATH.
echo Install Python 3.11 or newer from https://www.python.org/downloads/windows/ and try again.
pause
exit /b 1

:old_python
echo The Python found above is older than 3.11.
echo Install Python 3.11 or newer from https://www.python.org/downloads/windows/ and try again.
pause
exit /b 1

:venv_failed
echo Python is installed, but the local environment could not be created.
echo Check that this folder is writable and that antivirus did not block Python.
pause
exit /b 1

:packages_failed
echo Python is installed, but the required packages could not be installed.
echo Check the internet connection or ask IT to allow Python to install packages.
pause
exit /b 1
