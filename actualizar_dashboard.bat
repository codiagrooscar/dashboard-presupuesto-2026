@echo off
chcp 65001 > nul
echo ========================================================
echo   CODIAGRO - ACTUALIZADOR DIARIO DE DASHBOARD WEB
echo ========================================================
echo.
setlocal enabledelayedexpansion
set "PY=C:\Users\oscar.ocampo\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

echo 1. Regenerando dataset y dashboard web...
"%PY%" scripts/build_dataset.py
if %ERRORLEVEL% NEQ 0 ( echo [ERROR] Fallo en build_dataset.py & exit /b 1 )

"%PY%" scripts/build_production_data.py
if %ERRORLEVEL% NEQ 0 ( echo [ERROR] Fallo en build_production_data.py & exit /b 1 )

"%PY%" scripts/generar_orden_fabricacion_excel.py
if %ERRORLEVEL% NEQ 0 ( echo [ERROR] Fallo en generar_orden_fabricacion_excel.py & exit /b 1 )

"%PY%" scripts/build_treasury_data.py
if %ERRORLEVEL% NEQ 0 ( echo [ERROR] Fallo en build_treasury_data.py & exit /b 1 )

"%PY%" scripts/build_profitability_data.py
if %ERRORLEVEL% NEQ 0 ( echo [ERROR] Fallo en build_profitability_data.py & exit /b 1 )

"%PY%" scripts/build_bi_modules.py
if %ERRORLEVEL% NEQ 0 ( echo [ERROR] Fallo en build_bi_modules.py & exit /b 1 )

echo 1b. Integrando Full Costing analitico sept/oct...
if exist "scripts\merge_fullcost_drilldown.py" (
    "%PY%" scripts/merge_fullcost_drilldown.py
    if %ERRORLEVEL% NEQ 0 ( echo [ERROR] Fallo en merge_fullcost_drilldown.py & exit /b 1 )
) else if exist "scratch\merge_fullcost_drilldown.py" (
    "%PY%" scratch/merge_fullcost_drilldown.py
    if %ERRORLEVEL% NEQ 0 ( echo [ERROR] Fallo en merge_fullcost_drilldown.py & exit /b 1 )
)

REM Validacion transaccional de integridad de datos generados
if not exist "web_dashboard_direccion\fullcost_data.json" ( echo [ERROR] Archivo fullcost_data.json no encontrado & exit /b 1 )
if not exist "web_dashboard_direccion\dashboard_data.json" ( echo [ERROR] Archivo dashboard_data.json no encontrado & exit /b 1 )

REM Asegurar que ningun dato confidencial este en la version publica web_dashboard
del /q "web_dashboard\treasury_data.*" 2>nul
del /q "web_dashboard\profitability_data.*" 2>nul
del /q "web_dashboard\profitability_logic.js" 2>nul
del /q "web_dashboard\Prevision_Tesoreria_CashFlow_2026*.xlsx" 2>nul

echo.
echo 2. Regenerando informe Excel de seguimiento...
"C:\Users\oscar.ocampo\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" scripts/generate_excel_dashboard.py

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
