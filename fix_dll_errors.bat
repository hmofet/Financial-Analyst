@echo off
REM ============================================
REM Fix DLL Errors - Run this if you get 
REM "ordinal could not be located" errors
REM ============================================

echo.
echo ==========================================
echo   DLL Error Fix Script
echo ==========================================
echo.

echo This script will help fix common DLL errors.
echo.

echo Step 1: Installing Visual C++ Redistributable...
echo Please download and install from:
echo https://aka.ms/vs/17/release/vc_redist.x64.exe
echo.
echo Press any key after installing VC++ Redistributable...
pause > nul

echo.
echo Step 2: Reinstalling numpy with compatible version...
pip uninstall numpy -y
pip install numpy==1.24.3

echo.
echo Step 3: Reinstalling pandas...
pip uninstall pandas -y
pip install pandas==2.0.3

echo.
echo ==========================================
echo   Done! Try running the application again.
echo ==========================================
echo.
echo If errors persist:
echo 1. Use Python 3.10 or 3.11 (not 3.12+)
echo 2. Create a fresh virtual environment
echo 3. Install from requirements.txt
echo.

pause
