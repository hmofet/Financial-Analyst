@echo off
REM ============================================
REM Trading Report Builder - Windows Build Script
REM WITH DLL COMPATIBILITY FIXES
REM ============================================

echo.
echo ==========================================
echo   Trading Report Builder - Build Script
echo ==========================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo Please install Python 3.8+ from https://python.org
    pause
    exit /b 1
)

echo [1/5] Creating clean virtual environment...
if exist venv rmdir /s /q venv
python -m venv venv
call venv\Scripts\activate.bat

echo.
echo [2/5] Upgrading pip and installing wheel...
python -m pip install --upgrade pip setuptools wheel

echo.
echo [3/5] Installing compatible dependencies...
REM Use specific versions known to work well with PyInstaller
pip install pandas==2.0.3
pip install numpy==1.24.3
pip install openpyxl==3.1.2
pip install matplotlib==3.7.2
pip install reportlab==4.0.4
pip install pyinstaller==6.3.0

echo.
echo [4/5] Building executable with PyInstaller...
REM Use --collect-all for problematic packages
pyinstaller --noconfirm --onefile --windowed ^
    --name "TradingReportBuilder" ^
    --collect-all matplotlib ^
    --collect-all numpy ^
    --hidden-import "PIL._tkinter_finder" ^
    --hidden-import "pkg_resources.py2_warn" ^
    --hidden-import "pkg_resources.markers" ^
    --exclude-module scipy ^
    --exclude-module sklearn ^
    --exclude-module PyQt5 ^
    --exclude-module PyQt6 ^
    trading_report_builder.py

echo.
echo [5/5] Cleaning up...
if exist build rmdir /s /q build
if exist __pycache__ rmdir /s /q __pycache__
if exist *.spec del /q *.spec

echo.
echo ==========================================
echo   BUILD COMPLETE!
echo ==========================================
echo.
echo Executable location: dist\TradingReportBuilder.exe
echo.
echo If you still get DLL errors, try:
echo   1. Install Visual C++ Redistributable 2015-2022
echo   2. Run: pip install --upgrade numpy
echo   3. Use the alternative build script
echo ==========================================
echo.

pause
