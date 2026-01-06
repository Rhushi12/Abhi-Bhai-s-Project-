@echo off
echo Starting Excel Automation Tool...
cd /d "%~dp0"
python -m streamlit run app.py
pause
