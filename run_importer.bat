@echo off
echo Starting LAB14 Excel to SQLite Importer...
echo.
echo This will start the web server on port 5666
echo Open your web browser and go to: http://localhost:5666
echo.
echo Press Ctrl+C to stop the server
echo.
echo Checking if executable exists...
if exist "LAB14_Excel_Importer\LAB14_Excel_Importer.exe" (
    echo Found executable, starting...
    LAB14_Excel_Importer\LAB14_Excel_Importer.exe
) else (
    echo ERROR: LAB14_Excel_Importer\LAB14_Excel_Importer.exe not found!
    echo Current directory: %CD%
    echo Files in current directory:
    dir
    echo.
    echo Please make sure you extracted the zip file completely.
    pause
)
pause 