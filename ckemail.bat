@echo off
cd /d "%~dp0"

REM Check setup has been run
if not exist venv (
    echo ERROR: Virtual environment not found.
    echo Please run setup.bat first.
    pause
    exit /b 1
)

if not exist .env (
    echo ERROR: .env file not found.
    echo Please run setup.bat first.
    pause
    exit /b 1
)

REM Run the summarizer
venv\Scripts\python summarize_inbox.py

echo.
pause
