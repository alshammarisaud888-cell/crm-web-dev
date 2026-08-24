@echo off
cd /d "%~dp0"
if exist .venv (
    rmdir /s /q .venv
)
echo The previous Python environment has been removed.
echo Now run run_app.bat.
pause
