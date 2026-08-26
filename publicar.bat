@echo off
chcp 65001 >nul
cd /d "%~dp0"
python publicar.py %*
if errorlevel 1 (
  echo.
  echo *** Algo fallo. Revisa el mensaje de arriba. ***
)
echo.
pause
