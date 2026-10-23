@echo off
chcp 65001 > nul
echo ========================================================
echo   CODIAGRO - PRESENTACION EJECUTIVA PRESUPUESTO 2027
echo ========================================================
echo.
echo Iniciando aplicacion web independiente en http://localhost:3035 ...
start "" "http://localhost:3035"
"C:\Users\oscar.ocampo\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" presentacion_presupuesto_2027/server.py
pause
