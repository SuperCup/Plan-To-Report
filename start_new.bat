@echo off
setlocal
cd /d "%~dp0"
taskkill /F /IM PlanToReport.exe >nul 2>&1
call build.bat
if errorlevel 1 (
    echo Build failed, starting from source...
    call dev_start.bat
    exit /b 1
)
start "" "dist_release\PlanToReport\PlanToReport.exe"
endlocal
