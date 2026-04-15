@echo off
setlocal
cd /d "%~dp0"
echo.
echo ========================================
echo   ACTUALIZANDO DASHBOARD HSEC
echo ========================================
echo.

if not exist ".venv\Scripts\python.exe" (
  echo No se encontro el entorno de Python del proyecto.
  if /I not "%~1"=="nopause" pause
  exit /b 1
)

.venv\Scripts\python.exe process_excel.py
if errorlevel 1 (
  echo.
  echo Ocurrio un error al procesar el Excel.
  if /I not "%~1"=="nopause" pause
  exit /b 1
)

echo.
echo Dashboard actualizado correctamente.
start "" "index.html"

if /I not "%~1"=="nopause" pause
exit /b 0
