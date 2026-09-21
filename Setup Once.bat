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
rem Also search standard installer locations in case PATH was not updated.
if not defined PYTHON_EXE (
  for /d %%D in ("%ProgramFiles%\Python3*") do (
    if exist "%%~D\python.exe" if not defined PYTHON_EXE set "PYTHON_EXE=%%~D\python.exe"
  )
)
if not defined PYTHON_EXE (
  for /d %%D in ("%LocalAppData%\Programs\Python\Python3*") do (
    if exist "%%~D\python.exe" if not defined PYTHON_EXE set "PYTHON_EXE=%%~D\python.exe"
  )
)
if not defined PYTHON_EXE goto no_python
"%PYTHON_EXE%" -c "import sys; print('Found Python ' + sys.version.split()[0]); raise SystemExit(0 if sys.version_info >= (3, 11) else 1)"
if errorlevel 1 goto old_python
if not exist ".venv\Scripts\python.exe" (
  echo Creating the automation environment...
  "%PYTHON_EXE%" -m venv .venv
  if errorlevel 1 goto venv_failed
)
echo Installing all required packages into the automation environment...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto packages_failed
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto packages_failed
".venv\Scripts\python.exe" -c "import openpyxl, win32com.client; print('Required packages verified.')"
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
:venv_failed
echo Python was found, but the local automation environment could not be created.
echo Check that this folder is writable and that antivirus did not block Python.
pause
exit /b 1
:packages_failed
echo Python is installed, but the required packages could not be installed.
echo Ask IT to allow Python to create the local .venv folder and download packages.
pause
exit /b 1
