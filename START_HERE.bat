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
    echo AI provider key setup is still needed.
    echo Launching setup now...
    echo.
    call setup.bat
    if errorlevel 1 exit /b 1
)

echo Launching the app...
call run_summary.bat
