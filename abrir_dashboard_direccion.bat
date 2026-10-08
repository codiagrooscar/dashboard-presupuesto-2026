@echo off
chcp 65001 > nul
echo ========================================================
echo   CODIAGRO - DASHBOARD DIRECCION GENERAL Y FINANZAS
echo   (Comercial, Fabricacion, Tesoreria y Rentabilidad)
echo ========================================================
echo.
echo Iniciando servidor local en http://localhost:3026 ...
start "" "http://localhost:3026"
"C:\Users\oscar.ocampo\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" web_dashboard_direccion/server.py
pause
