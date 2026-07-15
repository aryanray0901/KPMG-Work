@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title KPMG Deck Refresh

set "PYTHON_CMD="
py -3 -c "import sys; print(sys.executable)" >nul 2>&1
if not errorlevel 1 set "PYTHON_CMD=py -3"
if defined PYTHON_CMD goto python_found

python -c "import sys; print(sys.executable)" >nul 2>&1
if not errorlevel 1 set "PYTHON_CMD=python"
if defined PYTHON_CMD goto python_found

echo.
echo Python 3 was not found on this computer.
echo Install Python from python.org and select Add Python to PATH.
echo Then run this file again.
echo.
start "" "https://www.python.org/downloads/windows/"
goto launcher_failed

:python_found
echo Python found.

if not exist "venv\Scripts\python.exe" goto create_venv
"venv\Scripts\python.exe" -c "import sys" >nul 2>&1
if errorlevel 1 goto repair_venv
goto check_packages

:repair_venv
echo Repairing the local Python environment...
if exist "venv" rmdir /s /q "venv"

:create_venv
echo Creating the local Python environment...
%PYTHON_CMD% -m venv "venv"
if errorlevel 1 goto launcher_failed

:check_packages
"venv\Scripts\python.exe" -c "import flask, pptx, openpyxl, pandas, rapidfuzz, fitz" >nul 2>&1
if not errorlevel 1 goto run_app

echo Installing required packages...
"venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto launcher_failed
"venv\Scripts\python.exe" -m pip install -r "requirements.txt"
if errorlevel 1 goto launcher_failed

:run_app
echo.
echo Starting KPMG Deck Refresh...
echo Your browser should open to http://127.0.0.1:5050
echo Keep this window open while using the app.
echo.
"venv\Scripts\python.exe" "app.py"
set "APP_EXIT=%ERRORLEVEL%"
echo.
if not "%APP_EXIT%"=="0" echo KPMG Deck Refresh stopped with error code %APP_EXIT%.
if "%APP_EXIT%"=="0" echo KPMG Deck Refresh has stopped.
echo Press any key to close this window.
pause >nul
exit /b %APP_EXIT%

:launcher_failed
echo.
echo Setup failed. Read the error shown above.
echo Press any key to close this window.
pause >nul
exit /b 1
