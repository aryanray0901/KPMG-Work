@echo off
setlocal
cd /d "%~dp0"

echo ==========================================
echo   Recruiting Radar
echo ==========================================
echo.

set "PYCMD="
where python >nul 2>nul
if %errorlevel%==0 (
  set "PYCMD=python"
) else (
  where py >nul 2>nul
  if %errorlevel%==0 (
    set "PYCMD=py"
  )
)

if "%PYCMD%"=="" (
  echo ERROR: Python was not found on this computer.
  echo Install Python 3 from https://www.python.org/downloads/ and try again.
  echo Be sure to check "Add python.exe to PATH" during install.
  echo.
  pause
  exit /b 1
)

if not exist "venv\.setup_complete" (
  echo Setting up ^(first run only, this takes a minute^)...
  if exist venv rmdir /s /q venv
  %PYCMD% -m venv venv
  if errorlevel 1 (
    echo.
    echo ERROR: Failed to create the virtual environment.
    echo.
    pause
    exit /b 1
  )

  call venv\Scripts\pip install --quiet --upgrade pip
  call venv\Scripts\pip install --quiet -r requirements.txt
  if errorlevel 1 (
    echo.
    echo ERROR: Failed to install dependencies. Check your internet connection, then try again.
    echo.
    pause
    exit /b 1
  )

  type nul > venv\.setup_complete
  echo Setup complete.
  echo.
)

echo Starting Recruiting Radar at http://127.0.0.1:5110
echo Leave this window open while you use the app. Close it or press Ctrl+C to stop.
echo.
call venv\Scripts\python app.py

echo.
echo Recruiting Radar has stopped.
pause
