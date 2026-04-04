@echo off
cd /d "%~dp0"

if not exist venv (
    echo ERROR: Virtual environment not found. Please run setup.bat first.
    pause
    exit /b 1
)

if not exist .env (
    echo ERROR: .env file not found. Please run setup.bat first.
    pause
    exit /b 1
)

echo Starting Outlook Inbox Summariser...
echo Browser will open automatically.
echo Close this window to stop the server.
echo.

venv\Scripts\python app.py
