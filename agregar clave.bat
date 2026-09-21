@echo off
REM Le suma OTRA clave valida a un usuario, sin tocar la que ya tiene.
REM Sirve para agregar un PIN corto ademas de la clave larga de Power BI:
REM despues entra con cualquiera de las dos, indistintamente.
cd /d "%~dp0"
echo.
echo   ==========================================
echo    Agregar otra clave a un usuario
echo   ==========================================
echo.
python gestionar_accesos.py --listar
echo.
echo   El numero entre parentesis es cuantas claves tiene cada uno.
echo.
set /p USUARIO=Usuario (enter para salir):
if "%USUARIO%"=="" goto :fin
echo.
python gestionar_accesos.py --agregar-clave "%USUARIO%"
echo.
echo   Acordate de publicar: publicar.bat
:fin
echo.
pause
