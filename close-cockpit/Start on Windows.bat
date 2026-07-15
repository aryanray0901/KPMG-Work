@echo off
cd /d "%~dp0"
if not exist venv (
  echo Setting up (first run only, this takes a minute)...
  python -m venv venv
  venv\Scripts\pip install --quiet -r requirements.txt
)
echo Starting Close Cockpit...
venv\Scripts\python app.py
pause
