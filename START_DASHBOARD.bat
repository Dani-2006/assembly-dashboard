@echo off
echo =====================================================
echo  Department of Case Assembly - Operational Dashboard
echo =====================================================
echo.
echo Starting dashboard server...
echo Dashboard will be available at: http://localhost:5050
echo.
echo The dashboard reads live data from:
echo   Assembly_Daily_Data.xlsx
echo.
echo Modify the Excel file anytime - the dashboard
echo auto-reloads every 3 seconds when file changes.
echo.
echo Press Ctrl+C to stop the server.
echo =====================================================
echo.
python "%~dp0app.py"
pause
