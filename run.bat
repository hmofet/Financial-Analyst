@echo off
REM ============================================
REM Financial Analyst - Trading Report Builder
REM Quick Run Script v1.1.1
REM ============================================

echo.
echo ============================================
echo   Trading Report Builder - Launcher
echo ============================================
echo.

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo.
    echo Please install Python 3.10 or 3.11 from:
    echo   https://www.python.org/downloads/
    echo.
    echo IMPORTANT: Check "Add Python to PATH" during installation!
    echo.
    pause
    exit /b 1
)

REM Display Python version
echo Found Python:
python --version
echo.

REM Check if dependencies are installed
echo Checking dependencies...
python -c "import pandas; import numpy; import openpyxl; import reportlab" >nul 2>&1
if errorlevel 1 (
    echo.
    echo ==========================================
    echo   Installing required packages...
    echo   This only needs to happen once.
    echo ==========================================
    echo.
    pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo ERROR: Failed to install dependencies.
        echo.
        echo Try running manually:
        echo   pip install -r requirements.txt
        echo.
        pause
        exit /b 1
    )
    echo.
    echo Dependencies installed successfully!
    echo.
)

echo All dependencies found. Starting application...
echo.

REM Run the application
python trading_report_builder.py

if errorlevel 1 (
    echo.
    echo ==========================================
    echo ERROR: Application exited with an error.
    echo ==========================================
    echo.
    echo Common fixes:
    echo   1. Reinstall dependencies:
    echo      pip install -r requirements.txt
    echo.
    echo   2. If DLL errors occur, run:
    echo      fix_dll_errors.bat
    echo.
    pause
)
