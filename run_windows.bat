@echo off
echo ==========================================
echo Starting Tech Efficiency Analyzer
echo ==========================================

REM Check if Python is installed
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python is not installed or not in your PATH.
    echo Please install Python from https://www.python.org/downloads/
    pause
    exit /b
)

REM Create virtual environment if it doesn't exist
IF NOT EXIST "venv\Scripts\activate.bat" (
    echo [INFO] First time setup: Creating virtual environment...
    python -m venv venv
    IF %ERRORLEVEL% NEQ 0 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b
    )
)

REM Activate virtual environment
call venv\Scripts\activate.bat

REM Install/Upgrade dependencies
echo [INFO] Checking dependencies...
pip install -r requirements.txt >nul

REM Run the dashboard
echo [INFO] Starting the dashboard...
streamlit run dashboard.py

pause
