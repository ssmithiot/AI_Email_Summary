@echo off
echo ============================================
echo  Outlook Email Summary - First-Time Setup
echo ============================================
echo.

REM Check Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Install Python from https://python.org
    pause
    exit /b 1
)

REM Create virtual environment if needed
if not exist venv (
    echo Creating virtual environment...
    python -m venv venv
    if errorlevel 1 (
        echo ERROR: Failed to create virtual environment.
        pause
        exit /b 1
    )
) else (
    echo Virtual environment already exists.
)

REM Install dependencies
echo Installing dependencies...
venv\Scripts\python -m pip install --upgrade pip >nul
venv\Scripts\pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install dependencies.
    pause
    exit /b 1
)

REM Create .env if it doesn't exist
if not exist .env (
    copy .env.example .env >nul
    echo.
    echo ============================================
    echo  ACTION REQUIRED: Add your OpenAI API key
    echo ============================================
    echo  1. Visit: https://platform.openai.com/api-keys
    echo  2. Create a new secret key
    echo  3. Open the .env file in this folder
    echo  4. Replace sk-your-key-here with your real key
    echo ============================================
    notepad .env
) else (
    echo .env already exists - skipping.
)

echo.
echo Setup complete!
echo Run run_summary.bat to launch the browser app.
echo Run ckemail.bat for the console version.
echo.
pause
