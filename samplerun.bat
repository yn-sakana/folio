@echo off
echo === Build + Run folio (sample) ===
powershell -ExecutionPolicy Bypass -File "%~dp0scripts\Build-Addin.ps1" -Sample
if errorlevel 1 (
    echo Build failed.
    pause
    exit /b 1
)
start "" "%~dp0folio.xlsm"
timeout /t 2 /nobreak >nul
start "" "%~dp0sample\folio-sample.xlsx"
echo.
echo Done. Run Alt+F8 ^> Folio_ShowPanel in Excel.
