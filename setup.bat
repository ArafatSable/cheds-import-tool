@echo off
REM One-time setup: installs the Python packages this app needs.
REM Run this once, then use run_web_app.bat or run_cli.bat.
cd /d "%~dp0"
python -m pip install -r requirements.txt
echo.
echo Setup done. You can now double-click run_web_app.bat or run_cli.bat.
pause
