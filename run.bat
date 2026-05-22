@echo off
setlocal
cd /d "%~dp0"
set PYTHONPATH=src
python -m plan_to_report
endlocal
