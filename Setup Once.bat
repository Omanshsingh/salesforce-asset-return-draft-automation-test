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
echo Installing the required packages for this Windows user...
%PYTHON_CMD% -m pip install --user -r requirements.txt
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
:packages_failed
echo Python is installed, but the required packages could not be installed.
echo Check the internet connection or ask IT to allow Python package installation.
pause
exit /b 1
