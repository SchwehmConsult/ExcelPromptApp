@echo off
cd /d "%~dp0"
python -m venv .venv
call venv\Scripts\activate
pip install -r requirements.txt
echo Setup complete! You can now open matrix_template2.xlsm
pause
