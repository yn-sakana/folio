@echo off
echo === Opening sample + addin ===
start "" "%~dp0sample\folio-sample.xlsx"
timeout /t 3 /nobreak >nul
start "" "%~dp0folio.xlsm"
echo.
echo Done. Run Alt+F8 ^> Folio_ShowPanel in Excel.
