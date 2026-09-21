@echo off
setlocal
cd /d "%~dp0"
set "PYTHON_EXE="
rem Ignore the WindowsApps alias: it can launch Python Manager and hang while
rem attempting a blocked Microsoft Store installation.
for /f "delims=" %%P in ('where python 2^>nul') do (
  echo %%P | findstr /i "\\WindowsApps\\" >nul
  if errorlevel 1 if not defined PYTHON_EXE set "PYTHON_EXE=%%P"
)
if not defined PYTHON_EXE goto no_python
"%PYTHON_EXE%" -c "import sys; print('Found Python ' + sys.version.split()[0]); raise SystemExit(0 if sys.version_info >= (3, 11) else 1)"
if errorlevel 1 goto old_python
echo Installing the required packages for this Windows user...
"%PYTHON_EXE%" -m pip install --user -r requirements.txt
if errorlevel 1 goto packages_failed
echo Setup complete. Double-click Prepare Drafts.bat.
pause
exit /b 0
:no_python
echo A real Python installation was not found.
echo The WindowsApps Python alias is not sufficient and may be blocked by company policy.
echo Ask IT to install Python 3.11 or newer from https://www.python.org/downloads/windows/
echo During installation, enable ^"Add python.exe to PATH^".
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
