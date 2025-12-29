@echo off
REM ============================================================================
REM Clockify Report Converter - Build Script
REM Creates a standalone Windows executable using PyInstaller
REM ============================================================================

echo.
echo ========================================
echo  Clockify Report Converter - Build
echo ========================================
echo.

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python and try again.
    pause
    exit /b 1
)

REM Check if PyInstaller is installed
python -c "import PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo PyInstaller not found. Installing...
    python -m pip install pyinstaller
)

REM Check if required packages are installed
echo Checking dependencies...
python -m pip install -r requirements.txt -q

echo.
echo Building executable...
echo.

REM Build the executable
python -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "ClockifyReportConverter" ^
    --add-data "requirements.txt;." ^
    --clean ^
    --noconfirm ^
    clockify_app.py

echo.
if exist "dist\ClockifyReportConverter.exe" (
    echo ========================================
    echo  BUILD SUCCESSFUL!
    echo ========================================
    echo.
    echo Executable created at:
    echo   dist\ClockifyReportConverter.exe
    echo.
    echo You can now distribute this file to users.
    echo No Python installation required to run it!
) else (
    echo ========================================
    echo  BUILD FAILED
    echo ========================================
    echo Please check the error messages above.
)

echo.
pause

