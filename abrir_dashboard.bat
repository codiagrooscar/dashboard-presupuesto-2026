@echo off
chcp 65001 > nul
echo ========================================================
echo   CODIAGRO - DASHBOARD DE SEGUIMIENTO PRESUPUESTARIO
echo ========================================================
echo.
echo Iniciando servidor local en http://localhost:3025 ...
start "" "http://localhost:3025"
"C:\Users\oscar.ocampo\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" web_dashboard/server.py
pause
