#!/bin/bash
cd "$(dirname "$0")"

if [ -x "venv/bin/python3" ]; then
  PYTHON_CMD="venv/bin/python3"
elif command -v python3 >/dev/null 2>&1; then
  PYTHON_CMD="python3"
else
  echo "ERROR: Python 3 was not found."
  read -p "Press Enter to close this window..."
  exit 1
fi

"$PYTHON_CMD" configure_api_key.py
STATUS=$?
echo ""
if [ $STATUS -eq 0 ]; then
  echo "API key updated. Restart Deck Refresh if it is already running."
else
  echo "The API key was not changed."
fi
read -p "Press Enter to close this window..."
exit $STATUS
