@echo off
setlocal
cd /d "%~dp0"
python -m pip install -r requirements.txt -q
taskkill /F /IM PlanToReport.exe >nul 2>&1
ping 127.0.0.1 -n 3 >nul
if exist "build" rmdir /s /q "build"
if exist "dist_release" rmdir /s /q "dist_release"
python -m PyInstaller --noconfirm --windowed --name PlanToReport --distpath dist_release --icon "assets\app.ico" --add-data "templates;templates" --add-data "assets;assets" --paths src launcher.py
if errorlevel 1 (
    echo Build failed.
    exit /b 1
)
xcopy "templates" "dist_release\PlanToReport\templates\" /E /I /Y /Q
xcopy "assets" "dist_release\PlanToReport\assets\" /E /I /Y /Q
if not exist "dist_release\PlanToReport\config" mkdir "dist_release\PlanToReport\config"
copy /Y "config\app_settings.example.json" "dist_release\PlanToReport\config\" >nul
if exist "build" rmdir /s /q "build"
echo Build OK.
echo Run: dist_release\PlanToReport\PlanToReport.exe
echo Distribute: run package_release.ps1, then send releases\PlanToReport-win64-*.zip
endlocal
