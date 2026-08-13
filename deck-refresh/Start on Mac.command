#!/bin/bash
cd "$(dirname "$0")"

echo "=========================================="
echo "  Deck Refresh"
echo "=========================================="
echo ""

if ! command -v python3 >/dev/null 2>&1; then
  echo "ERROR: python3 was not found on this Mac."
  echo "Install Python 3 from https://www.python.org/downloads/ and try again."
  echo ""
  read -p "Press Enter to close this window..."
  exit 1
fi

if [ ! -f "venv/.setup_complete" ]; then
  echo "Setting up (first run only, this takes a minute)..."
  if [ ! -x "venv/bin/python3" ]; then
    python3 -m venv venv
    if [ $? -ne 0 ]; then
      echo ""
      echo "ERROR: Failed to create the virtual environment."
      echo "Try running this in Terminal for more detail: python3 -m venv venv"
      echo ""
      read -p "Press Enter to close this window..."
      exit 1
    fi
  fi

  ./venv/bin/python3 -m ensurepip --upgrade >/dev/null 2>&1
  ./venv/bin/python3 -m pip install --quiet --disable-pip-version-check -r requirements.txt
  if [ $? -ne 0 ]; then
    echo ""
    echo "ERROR: Failed to install dependencies. Check your internet connection, then try again."
    echo "This window will show the full error if you run:"
    echo "  ./venv/bin/python3 -m pip install -r requirements.txt"
    echo ""
    read -p "Press Enter to close this window..."
    exit 1
  fi

  touch venv/.setup_complete
  echo "Setup complete."
  echo ""
fi

if ! ./venv/bin/python3 -c "import flask, fitz, openai, dotenv, pandas, openpyxl, pptx" >/dev/null 2>&1; then
  echo "Installing new Deck Refresh features..."
  ./venv/bin/python3 -m pip install --quiet --disable-pip-version-check -r requirements.txt
  if [ $? -ne 0 ]; then
    rm -f venv/.setup_complete
    echo "ERROR: Failed to install dependencies."
    read -p "Press Enter to close this window..."
    exit 1
  fi
fi

if [ -d "/Applications/Microsoft PowerPoint.app" ] || [ -d "$HOME/Applications/Microsoft PowerPoint.app" ]; then
  echo "Mac preview renderer: Microsoft PowerPoint"
  echo "macOS may ask for permission to let Terminal control PowerPoint. Choose Allow."
elif [ -x "/Applications/LibreOffice.app/Contents/MacOS/soffice" ] || [ -x "$HOME/Applications/LibreOffice.app/Contents/MacOS/soffice" ]; then
  echo "Mac preview renderer: LibreOffice"
elif [ -d "/Applications/Keynote.app" ] || [ -d "$HOME/Applications/Keynote.app" ]; then
  echo "Mac preview renderer: Apple Keynote"
  echo "macOS may ask for permission to let Terminal control Keynote. Choose Allow."
else
  echo "Preview note: install Microsoft PowerPoint, Apple Keynote, or LibreOffice for live slide images."
fi

if ! ./venv/bin/python3 -c "from dotenv import load_dotenv; import os,sys; load_dotenv(); sys.exit(0 if os.getenv('OPENAI_API_KEY') else 1)" >/dev/null 2>&1; then
  echo "ERROR: The bundled OpenAI API key could not be loaded from .env."
  echo "Keep the .env file in the same folder as this launcher."
  read -p "Press Enter to close this window..."
  exit 1
fi
echo "OpenAI API key loaded."
echo ""

echo "Starting Deck Refresh at http://127.0.0.1:5050"
echo "Leave this window open while you use the app. Close it or press Ctrl+C to stop."
echo ""
./venv/bin/python3 app.py

echo ""
echo "Deck Refresh has stopped."
read -p "Press Enter to close this window..."
