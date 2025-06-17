@echo off
cd /d "%~dp0"
echo Running prompt generation...
".venv\Scripts\python.exe" run_prompt.py

echo Launching Streamlit viewer...
".venv\Scripts\streamlit.exe" run view_results.py
pause
