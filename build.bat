@echo off
echo ============================================
echo  Building Lateness Report Generator EXE
echo ============================================
echo.

REM Install dependencies
pip install pandas openpyxl xlrd pyinstaller

echo.
echo Building EXE...
pyinstaller --onefile --windowed --name "LatenessReportGenerator" --icon=NONE app.py --add-data "report_generator.py;."

echo.
echo ============================================
echo  Build complete!
echo  EXE is at: dist\LatenessReportGenerator.exe
echo ============================================
pause
