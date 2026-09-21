@echo off
REM Cambia la contrasena de un usuario del tablero.
REM Se tipea aca y se convierte en hash en el acto: nunca queda escrita
REM en el Excel ni en ningun archivo. Por eso no se puede hacer a mano
REM desde Excel, que se publica entero en Vercel.
cd /d "%~dp0"
echo.
echo   ==========================================
echo    Cambiar la clave de un usuario
echo   ==========================================
echo.
python gestionar_accesos.py --listar
echo.
set /p USUARIO=Usuario a cambiar (enter para salir):
if "%USUARIO%"=="" goto :fin
echo.
python gestionar_accesos.py --set "%USUARIO%"
echo.
echo   Listo. Para que tome efecto en el tablero: publicar.bat
:fin
echo.
pause
