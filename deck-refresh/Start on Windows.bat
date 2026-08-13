@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ==========================================
echo   Deck Refresh
echo ==========================================
echo.

set "PYCMD="
py -3 -c "import sys" >nul 2>nul
if not errorlevel 1 set "PYCMD=py -3"
if not defined PYCMD (
  python -c "import sys; sys.exit(0 if sys.version_info.major == 3 else 1)" >nul 2>nul
  if not errorlevel 1 set "PYCMD=python"
)

if not defined PYCMD (
  echo ERROR: Python was not found on this computer.
  echo Install Python 3 from https://www.python.org/downloads/ and try again.
  echo Be sure to check "Add python.exe to PATH" during install.
  echo.
  pause
  exit /b 1
)

set "VENV_PY=%~dp0venv\Scripts\python.exe"
set "SETUP_MARKER=%~dp0venv\.setup_complete"

if not exist "venv\.setup_complete" (
  echo Setting up ^(first run only, this takes a minute^)...
  if not exist "%VENV_PY%" (
    %PYCMD% -m venv "%~dp0venv"
    if errorlevel 1 (
      echo.
      echo ERROR: Failed to create the virtual environment.
      echo.
      pause
      exit /b 1
    )
  )

  "%VENV_PY%" -m ensurepip --upgrade >nul 2>nul
  "%VENV_PY%" -m pip install --quiet --disable-pip-version-check -r "%~dp0requirements.txt"
  if errorlevel 1 (
    echo.
    echo ERROR: Failed to install dependencies. Check your internet connection, then try again.
    echo.
    pause
    exit /b 1
  )

  type nul > "%SETUP_MARKER%"
  echo Setup complete.
  echo.
)

"%VENV_PY%" -c "import flask, fitz, openai, dotenv, pandas, openpyxl, pptx" >nul 2>nul
if errorlevel 1 (
  echo Installing new Deck Refresh features...
  "%VENV_PY%" -m pip install --quiet --disable-pip-version-check -r "%~dp0requirements.txt"
  if errorlevel 1 (
    if exist "%SETUP_MARKER%" del /q "%SETUP_MARKER%" >nul 2>nul
    echo ERROR: Failed to install dependencies.
    pause
    exit /b 1
  )
)

"%VENV_PY%" -c "from dotenv import load_dotenv; import os,sys; load_dotenv(); sys.exit(0 if os.getenv('OPENAI_API_KEY') else 1)" >nul 2>nul
if errorlevel 1 (
  echo ERROR: The bundled OpenAI API key could not be loaded from .env.
  echo Keep the .env file in the same folder as this launcher.
  pause
  exit /b 1
)
echo OpenAI API key loaded.
echo.

echo Starting Deck Refresh at http://127.0.0.1:5050
echo Leave this window open while you use the app. Close it or press Ctrl+C to stop.
echo.
"%VENV_PY%" "%~dp0app.py"

echo.
echo Deck Refresh has stopped.
pause
