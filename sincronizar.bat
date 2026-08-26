@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo ==== Sincronizando los links de Power BI ====
echo.
python sincronizar_powerbi.py %*
if errorlevel 1 (
  echo.
  echo *** Hubo un problema. Revisa el mensaje de arriba. ***
)
echo.
echo Si quedo bien, corre actualizar.bat para publicar en Vercel.
echo.
pause
