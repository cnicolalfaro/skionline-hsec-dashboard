@echo off
REM Automatizacion diaria: actualiza el dashboard, genera PDF y envia correo.
REM Programar en el Task Scheduler de Windows a las 08:00 y 17:00.

setlocal
cd /d "%~dp0"

if not exist "logs" mkdir "logs"
set LOGFILE=logs\automatizacion_%date:~-4,4%%date:~-7,2%%date:~-10,2%.log

echo. >> "%LOGFILE%"
echo ================================================== >> "%LOGFILE%"
echo  EJECUCION %date% %time% >> "%LOGFILE%"
echo ================================================== >> "%LOGFILE%"

REM 1) Regenerar Excel + dashboard_data.js
call actualizar_dashboard.bat nopause >> "%LOGFILE%" 2>&1

REM 2) Generar PDF ejecutivo y enviar correo
call .venv\Scripts\activate.bat
python envio_reporte.py >> "%LOGFILE%" 2>&1
set ERR=%ERRORLEVEL%

REM 3) Publicar cambios a GitHub Pages (si hay)
git add -A >> "%LOGFILE%" 2>&1
git commit -m "Actualizacion automatica %date% %time%" >> "%LOGFILE%" 2>&1
git push origin main >> "%LOGFILE%" 2>&1

echo Terminado con codigo %ERR% >> "%LOGFILE%"
endlocal
exit /b %ERR%
