@echo off
cd /d "%~dp0"

echo ============================================
echo  Outlook Inbox Summarizer - Start Here
echo ============================================
echo.

if not exist venv (
    echo First-time setup has not been run yet.
    echo Launching setup now...
    echo.
    call setup.bat
    if errorlevel 1 exit /b 1
)

if not exist .env (
    echo OpenAI API key setup is still needed.
    echo Launching setup now...
    echo.
    call setup.bat
    if errorlevel 1 exit /b 1
)

findstr /b /c:"OPENAI_API_KEY=sk-your-key-here" .env >nul 2>&1
if not errorlevel 1 (
    echo Your OpenAI API key still needs to be added.
    echo Opening .env so you can paste it in now...
    echo.
    start "" notepad .env
    echo After you save the key, run START_HERE.bat again.
    pause
    exit /b 1
)

echo Launching the app...
call run_summary.bat
