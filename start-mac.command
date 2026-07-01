#!/bin/bash
# Double-click (or run in Terminal) to launch the Course Descriptor Builder on macOS/Linux.
cd "$(dirname "$0")"
echo "Setting up (first run only may take a minute)..."
python3 -m pip install -r requirements.txt >/dev/null 2>&1
echo ""
echo "Open http://localhost:5000 in your browser."
python3 app.py
