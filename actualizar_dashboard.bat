@echo off
chcp 65001 > nul
echo ========================================================
echo   CODIAGRO - ACTUALIZADOR DIARIO DE DASHBOARD WEB
echo ========================================================
echo.
echo 1. Regenerando dataset y dashboard web...
"C:\Users\oscar.ocampo\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" build_dataset.py

if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Hubo un problema al procesar los datos web.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo 2. Regenerando informe Excel de seguimiento...
"C:\Users\oscar.ocampo\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" generate_excel_dashboard.py

echo.
echo 3. Desplegando en Firebase Hosting...
call firebase deploy --only hosting

echo.
echo 4. Sincronizando con GitHub y GitHub Pages...
git add .
git commit -m "Actualizacion pedidos: %date% %time%"
git push origin main
git subtree push --prefix web_dashboard origin gh-pages

echo.
echo ========================================================
echo   ACTUALIZACION COMPLETADA CON EXITO!
echo   Web Firebase: https://mantenimiento-21758.web.app
echo   Web GitHub Pages: https://codiagrooscar.github.io/dashboard-presupuesto-2026/
echo   Excel: Seguimiento_Presupuesto_Sep_2026.xlsx
echo ========================================================
pause
