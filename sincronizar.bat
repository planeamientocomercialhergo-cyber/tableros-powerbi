@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo ==== Sincronizando links de Power BI (a _test) ====
echo.
python sincronizar_powerbi.py %*
echo.
pause
