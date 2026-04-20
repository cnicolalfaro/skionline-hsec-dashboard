@echo off
cd /d "%~dp0"
git add -A
git commit -m "Eliminar script temporal"
git push
echo Listo.
