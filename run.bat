@echo off
REM ============================================
REM Financial Analyst - Trading Report Builder
REM Quick Run Script
REM ============================================

echo Starting Trading Report Builder...
echo.

python trading_report_builder.py

if errorlevel 1 (
    echo.
    echo ==========================================
    echo ERROR: Failed to start application.
    echo ==========================================
    echo.
    echo Make sure Python and dependencies are installed:
    echo   pip install -r requirements.txt
    echo.
    echo If you get DLL errors, run:
    echo   fix_dll_errors.bat
    echo.
    pause
)
