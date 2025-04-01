@echo off
echo =============================================
echo         RFID BoM Generator Launcher
echo =============================================
echo.

REM Change to the script's directory
cd /d "%~dp0"

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python from https://www.python.org/downloads/
    pause
    exit /b 1
)

REM Check if Streamlit is installed
python -c "import streamlit" >nul 2>&1
if %errorlevel% neq 0 (
    echo Streamlit not found. Attempting to install required packages...
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo ERROR: Failed to install required packages
        pause
        exit /b 1
    )
)

echo Starting RFID BoM Generator...
echo The application will open in your default web browser.
echo.
echo Press Ctrl+C in this window to stop the application when done.
echo.

REM Start the Streamlit application
streamlit run rfid_bom_generator.py

REM If we get here, there was likely an error
if %errorlevel% neq 0 (
    echo.
    echo There was an error starting the application.
    pause
)
