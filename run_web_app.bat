@echo off
REM Double-click this to launch the web app in your browser.
cd /d "%~dp0"
python -m streamlit run app_streamlit.py
pause
