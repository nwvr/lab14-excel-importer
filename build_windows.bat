@echo off
echo Building LAB14 Excel to SQLite Importer for Windows...
echo.

echo Checking Python installation...
python --version
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8+ from https://python.org
    pause
    exit /b 1
)

echo.
echo Installing/updating required packages...
pip install --upgrade pip
pip install --upgrade pyinstaller pandas openpyxl xlsxwriter flask

echo.
echo Building executable...
pyinstaller importxl.spec

if errorlevel 1 (
    echo ERROR: Build failed!
    pause
    exit /b 1
)

echo.
echo Build completed successfully!
echo.
echo The executable is located in: dist\LAB14_Excel_Importer\
echo.
echo To run the application:
echo 1. Navigate to dist\LAB14_Excel_Importer\
echo 2. Double-click run_importer.bat
echo 3. Open http://localhost:5666 in your web browser
echo.
pause 