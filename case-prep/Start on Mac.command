#!/bin/bash
cd "$(dirname "$0")"

echo "=========================================="
echo "  Case Prep"
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
  rm -rf venv
  python3 -m venv venv
  if [ $? -ne 0 ]; then
    echo ""
    echo "ERROR: Failed to create the virtual environment."
    echo "Try running this in Terminal for more detail: python3 -m venv venv"
    echo ""
    read -p "Press Enter to close this window..."
    exit 1
  fi

  ./venv/bin/pip install --quiet --upgrade pip
  ./venv/bin/pip install --quiet -r requirements.txt
  if [ $? -ne 0 ]; then
    echo ""
    echo "ERROR: Failed to install dependencies. Check your internet connection, then try again."
    echo "This window will show the full error if you run:"
    echo "  ./venv/bin/pip install -r requirements.txt"
    echo ""
    read -p "Press Enter to close this window..."
    exit 1
  fi

  touch venv/.setup_complete
  echo "Setup complete."
  echo ""
fi

echo "Starting Case Prep at http://127.0.0.1:5100"
echo "Leave this window open while you use the app. Close it or press Ctrl+C to stop."
echo ""
./venv/bin/python3 app.py

echo ""
echo "Case Prep has stopped."
read -p "Press Enter to close this window..."
