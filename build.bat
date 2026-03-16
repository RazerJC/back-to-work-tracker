@echo off
echo ================================================
echo  HR Attendance Analytics System - EXE Builder
echo ================================================
echo.

REM Install requirements
pip install flask openpyxl pywebview pyinstaller

echo.
echo Building EXE...
pyinstaller --name "HR_Analytics" ^
  --onedir ^
  --windowed ^
  --add-data "templates;templates" ^
  --add-data "static;static" ^
  --icon=NUL ^
  main.py

echo.
echo ================================================
echo  Build complete!
echo  EXE is in: dist\HR_Analytics\HR_Analytics.exe
echo ================================================
pause
