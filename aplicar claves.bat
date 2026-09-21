@echo off
REM Lee "CLAVES TABLERO.xlsx" (la carpeta de arriba, la que NO se publica) y
REM vuelca las contrasenas al tablero como salt+hash.
REM Cerra el Excel antes de correr esto.
cd /d "%~dp0"
echo.
echo   ==========================================
echo    Aplicar las claves de CLAVES TABLERO.xlsx
echo   ==========================================
echo.
python gestionar_accesos.py --desde-planilla
if errorlevel 1 goto :fin
echo.
python gestionar_accesos.py --listar
:fin
echo.
pause
