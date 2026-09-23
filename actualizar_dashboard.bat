@echo off
chcp 65001 > nul
echo ========================================================
echo   CODIAGRO - ACTUALIZADOR DIARIO DE DASHBOARD WEB
echo ========================================================
echo.
echo 1. Regenerando dataset con nuevos pedidos y previsiones...
"C:\Users\oscar.ocampo\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" build_dataset.py

if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Hubo un problema al procesar los datos.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo 2. Subiendo actualización a GitHub y GitHub Pages...
git add web_dashboard/dashboard_data.json
git commit -m "Actualización diaria de pedidos: %date% %time%"
git push origin main
git subtree push --prefix web_dashboard origin gh-pages

echo.
echo ========================================================
echo   ¡ACTUALIZACIÓN COMPLETADA CON ÉXITO!
echo   Tu web pública se actualizará en unos segundos en:
echo   https://codiagrooscar.github.io/dashboard-presupuesto-2026/
echo ========================================================
pause
