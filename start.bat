@echo off
setlocal
cd /d "%~dp0retail-roi-python-project"
if not exist .venv (
    echo [.venv] not found. Creating virtual environment...
    python -m venv .venv
)
call .venv\Scripts\activate.bat
echo Installing/Updating dependencies...
pip install -r requirements.txt
pip install -e .
echo Starting Streamlit application...
streamlit run app.py
pause
