@echo off
SETLOCAL
TITLE Transol VMS - Startup

echo =====================================================
echo      Transol VMS - Windows Launcher
echo =====================================================
echo.

REM Navigate to project root (2 levels up)
pushd %~dp0..\..

REM 1. Check for Python
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python is not installed or not found in your PATH.
    echo Please install Python 3.9+ from https://www.python.org/downloads/
    echo IMPORTANT: Check "Add Python to PATH" during installation.
    pause
    popd
    EXIT /B
)

REM 2. Check/Create Virtual Environment
if not exist "venv" (
    echo [INFO] First time setup: Creating virtual environment...
    python -m venv venv
    if not exist "venv" (
        echo [ERROR] Failed to create 'venv'. Check permissions.
        pause
        popd
        EXIT /B
    )
    
    echo [INFO] Activating environment and installing dependencies...
    call venv\Scripts\activate
    
    if exist "requirements.txt" (
        python -m pip install --upgrade pip
        pip install -r requirements.txt
    ) else (
        echo [WARNING] requirements.txt not found! Skipping installation.
    )
) else (
    echo [INFO] Virtual environment found. Activating...
    call venv\Scripts\activate
)

REM 3. Run Migrations (Ensures DB is ready)
echo [INFO] Checking database...
python manage.py migrate

REM 4. Start Server and Browser
echo.
echo [INFO] Starting Server...
echo [INFO] The browser will open automatically in 5 seconds...

REM Start browser in background
timeout /t 5 >nul
start "" "http://127.0.0.1:8000/"

python manage.py runserver

popd
pause
