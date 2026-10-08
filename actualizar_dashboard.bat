@echo off
chcp 65001 > nul
echo ========================================================
echo   CODIAGRO - ACTUALIZADOR DIARIO DE DASHBOARD WEB
echo ========================================================
echo.
echo 1. Regenerando dataset y dashboard web...
"C:\Users\oscar.ocampo\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" build_dataset.py
"C:\Users\oscar.ocampo\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" build_production_data.py
"C:\Users\oscar.ocampo\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" generar_orden_fabricacion_excel.py
"C:\Users\oscar.ocampo\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" build_treasury_data.py
"C:\Users\oscar.ocampo\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" build_profitability_data.py

REM Asegurar que ningun dato confidencial este en la version publica web_dashboard
del /q "web_dashboard\treasury_data.*" 2>nul
del /q "web_dashboard\profitability_data.*" 2>nul
del /q "web_dashboard\profitability_logic.js" 2>nul
del /q "web_dashboard\Prevision_Tesoreria_CashFlow_2026*.xlsx" 2>nul

if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Hubo un problema al procesar los datos web.
    pause
    exit /b %ERRORLEVEL%
)

echo.
echo 2. Regenerando informe Excel de seguimiento...
"C:\Users\oscar.ocampo\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" generate_excel_dashboard.py

echo.
echo 3. Sincronizando con GitHub y GitHub Pages (solo Comercial y Fabricacion)...
git add .
git commit -m "Actualizacion pedidos: %date% %time%"
git push origin main
git subtree push --prefix web_dashboard origin gh-pages

echo.
echo ========================================================
echo   ACTUALIZACION COMPLETADA CON EXITO!
echo   Web GitHub Pages: https://codiagrooscar.github.io/dashboard-presupuesto-2026/
echo   Excel: Seguimiento_Presupuesto_Sep_2026.xlsx
echo ========================================================
pause
