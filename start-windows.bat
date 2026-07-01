@echo off
REM Double-click to launch the Course Descriptor Builder on Windows.
cd /d "%~dp0"
echo Setting up (first run only may take a minute)...
py -m pip install -r requirements.txt >nul 2>&1 || python -m pip install -r requirements.txt
echo.
echo Open http://localhost:5000 in your browser.
py app.py 2>nul || python app.py
pause
