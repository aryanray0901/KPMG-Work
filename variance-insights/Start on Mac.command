#!/bin/bash
# Double-click this file to launch Variance Insights.
cd "$(dirname "$0")"

if [ ! -d "venv" ]; then
  echo "Setting up (first run only, this takes a minute)..."
  python3 -m venv venv
  ./venv/bin/pip install --quiet -r requirements.txt
fi

echo "Starting Variance Insights..."
./venv/bin/python3 app.py
