@echo off
echo Starting Flask Form Application...
echo.
echo The browser will open automatically in a few seconds.
echo To stop the application, close this window or press Ctrl+C
echo.

REM Change to the directory where this batch file is located
cd /d "%~dp0"

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python from https://python.org
    pause
    exit /b 1
)

REM Check if required packages are installed
echo Checking required packages...
python -c "import flask, openpyxl" >nul 2>&1
if %errorlevel% neq 0 (
    echo Installing required packages...
    pip install flask openpyxl
    if %errorlevel% neq 0 (
        echo ERROR: Failed to install packages
        pause
        exit /b 1
    )
)

REM Run the Flask application
echo Starting the application...
python app.py

REM Keep the window open if there's an error
if %errorlevel% neq 0 (
    echo.
    echo Application stopped with an error.
    pause
)