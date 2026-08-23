@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

if not exist .venv\Scripts\python.exe (
    py -m venv .venv
)

call .venv\Scripts\activate
python -m pip install --upgrade pip
pip install --upgrade --no-cache-dir -r requirements.txt

flet run main.py
pause
endlocal
