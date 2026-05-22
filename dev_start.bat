@echo off
setlocal
cd /d "%~dp0"
taskkill /F /IM PlanToReport.exe >nul 2>&1
set PYTHONPATH=src
start "" pythonw -m plan_to_report
endlocal
