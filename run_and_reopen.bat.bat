@echo off
cd /d "%~dp0"

REM Use local virtual environment if available
set VENV_PATH=%~dp0venv\Scripts\activate.bat

REM Activate venv if exists
if exist "%VENV_PATH%" (
    call "%VENV_PATH%"
) else (
    echo [ERROR] No virtual environment found. Please run setup. & pause & exit /b
)

REM Run the prompt generator
echo Running prompt generation... >> log.txt
python run_prompt.py >> log.txt 2>&1

REM Launch Streamlit in background
start "" cmd /c "streamlit run viewer.py --server.port 8503"

REM Wait and reopen Excel
timeout /t 5 >nul

REM Use Excel from PATH (works on most systems)
start "" excel.exe "%~dp0matrix_template2.xlsm"
