@echo off
cd /d "%~dp0"
git add -A
git commit -m "Actualizar dashboard TARJA abril 2026"
git push
echo Listo.
