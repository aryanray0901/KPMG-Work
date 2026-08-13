@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "PYCMD="
if exist "%~dp0venv\Scripts\python.exe" set "PYCMD=%~dp0venv\Scripts\python.exe"
if not defined PYCMD (
  py -3 -c "import sys" >nul 2>nul
  if not errorlevel 1 set "PYCMD=py -3"
)
if not defined PYCMD (
  python -c "import sys" >nul 2>nul
  if not errorlevel 1 set "PYCMD=python"
)
if not defined PYCMD (
  echo ERROR: Python 3 was not found.
  pause
  exit /b 1
)

%PYCMD% "%~dp0configure_api_key.py"
if errorlevel 1 (
  echo.
  echo The API key was not changed.
) else (
  echo.
  echo API key updated. Restart Deck Refresh if it is already running.
)
pause
