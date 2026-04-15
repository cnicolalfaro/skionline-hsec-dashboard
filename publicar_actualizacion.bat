@echo off
setlocal
cd /d "%~dp0"

call actualizar_dashboard.bat nopause
if errorlevel 1 exit /b 1

set "GITEXE=C:\Program Files\Git\cmd\git.exe"
if not exist "%GITEXE%" set "GITEXE=git"

echo.
echo Subiendo cambios a GitHub Pages...
"%GITEXE%" add -A
"%GITEXE%" commit -m "Actualizar dashboard" >nul 2>&1
"%GITEXE%" push

echo.
echo Publicacion enviada. GitHub Pages puede tardar uno o dos minutos en reflejar el cambio.
pause
exit /b 0
