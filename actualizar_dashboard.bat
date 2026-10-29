@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

echo ======================================================================
echo   CODIAGRO - ACTUALIZADOR CONTROLADO DE FULL COSTING Y DASHBOARD
echo   (Fases separadas: 1. Generacion - 2. Validacion - 3. Activacion)
echo ======================================================================
echo.

set "PY=C:\Users\oscar.ocampo\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"
set "BACKUP_DIR=scratch\last_valid_backup"
set "STAGING_DIR=scratch\staging"
set "LAST_DATA_DATE=30/09/2026 (Cierre Full Costing Sep-26) / 29/10/2026 (Pedidos y Stock Oct-26)"

REM ----------------------------------------------------------------------
REM PASO 0: Resguardar la ultima version valida antes de regenerar salidas
REM ----------------------------------------------------------------------
echo [PASO 0] Resguardando ultima version valida en %BACKUP_DIR%...
"%PY%" -X utf8 scripts/clean_rollback.py backup
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ======================================================================
    echo   [ERROR CRITICO] Fallo en la creacion de la copia de seguridad.
    echo   Abortando actualizacion para proteger datos activos.
    echo ======================================================================
    exit /b 1
)

echo [PASO 0] Inicializando directorio temporal de staging en %STAGING_DIR%...
"%PY%" -X utf8 scripts/clean_rollback.py prepare_staging
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR CRITICO] Fallo al preparar el entorno temporal de staging.
    exit /b 1
)

REM Activar redirección a staging para que los generadores no toquen carpetas activas
set "CODIAGRO_STAGING_DIR=%STAGING_DIR%"

REM ----------------------------------------------------------------------
REM FASE 1: GENERACION DE TODAS LAS SALIDAS EN CARPETA TEMPORAL (STAGING)
REM ----------------------------------------------------------------------
echo.
echo [FASE 1] Generando salidas analiticas y operativas en staging temporal...

set "FAILED_STEP=build_dataset.py (Dataset comercial)"
"%PY%" -X utf8 scripts/build_dataset.py
if %ERRORLEVEL% NEQ 0 goto :ROLLBACK

if exist "scripts\build_multi_period.py" (
    set "FAILED_STEP=build_multi_period.py (Consolidacion multi-periodo Sep/Oct)"
    "%PY%" -X utf8 scripts/build_multi_period.py
    if !ERRORLEVEL! NEQ 0 goto :ROLLBACK
)

set "FAILED_STEP=build_production_data.py (Planificacion MRP)"
"%PY%" -X utf8 scripts/build_production_data.py
if %ERRORLEVEL% NEQ 0 goto :ROLLBACK

set "FAILED_STEP=generar_orden_fabricacion_excel.py (Orden de Fabricacion Excel)"
"%PY%" -X utf8 scripts/generar_orden_fabricacion_excel.py
if %ERRORLEVEL% NEQ 0 goto :ROLLBACK

set "FAILED_STEP=build_treasury_data.py (Tesoreria y Cash Flow)"
"%PY%" -X utf8 scripts/build_treasury_data.py
if %ERRORLEVEL% NEQ 0 goto :ROLLBACK

set "FAILED_STEP=build_profitability_data.py (Rentabilidad COGS)"
"%PY%" -X utf8 scripts/build_profitability_data.py
if %ERRORLEVEL% NEQ 0 goto :ROLLBACK

set "FAILED_STEP=build_marta_tracking_data.py (Seguimiento Cierre 2026)"
"%PY%" -X utf8 scripts/build_marta_tracking_data.py
if %ERRORLEVEL% NEQ 0 goto :ROLLBACK

set "FAILED_STEP=merge_fullcost_drilldown.py (Modelo Full Costing Septiembre 2026)"
"%PY%" -X utf8 scripts/merge_fullcost_drilldown.py
if %ERRORLEVEL% NEQ 0 goto :ROLLBACK

set "FAILED_STEP=generate_excel_dashboard.py (Excel Seguimiento_Presupuesto_Sep_2026.xlsx)"
"%PY%" -X utf8 scripts/generate_excel_dashboard.py
if %ERRORLEVEL% NEQ 0 goto :ROLLBACK

REM ----------------------------------------------------------------------
REM FASE 2: VALIDACION ESTRICTA DE CALCULOS Y CONSERVACION DE IMPORTES
REM ----------------------------------------------------------------------
echo.
echo [FASE 2] Ejecutando validacion de calculos, redondeos y coherencia en staging...
set "FAILED_STEP=validate_fullcost_update.py (Validacion de integridad y cuadre)"
"%PY%" -X utf8 scripts/validate_fullcost_update.py
if %ERRORLEVEL% NEQ 0 goto :ROLLBACK

REM Comprobar que ningun entregable activo fue modificado durante la generacion en staging
set "FAILED_STEP=clean_rollback.py verify_untouched"
"%PY%" -X utf8 scripts/clean_rollback.py verify_untouched
if %ERRORLEVEL% NEQ 0 goto :ROLLBACK

REM ----------------------------------------------------------------------
REM FASE 3: ACTIVACION LOCAL CONTROLADA Y AISLAMIENTO DE DATOS PUBLICOS
REM ----------------------------------------------------------------------
echo.
echo [FASE 3] Activando nueva version validada desde staging a carpetas activas...
set "FAILED_STEP=clean_rollback.py activate"
"%PY%" -X utf8 scripts/clean_rollback.py activate
if %ERRORLEVEL% NEQ 0 goto :ROLLBACK

REM Comprobar que la activacion conservo y promovio todos los entregables requeridos
set "FAILED_STEP=clean_rollback.py verify_activated"
"%PY%" -X utf8 scripts/clean_rollback.py verify_activated
if %ERRORLEVEL% NEQ 0 goto :ROLLBACK

set "CODIAGRO_STAGING_DIR="

echo.
echo ======================================================================
echo   [OK] ACTUALIZACION LOCAL VALIDADA Y ACTIVADA CON EXITO
echo   Fecha de datos disponibles: %LAST_DATA_DATE%
echo   Salidas actualizadas:
echo     - web_dashboard_direccion\fullcost_data.json / .js
echo     - web_dashboard_direccion\marta_tracking_data.json / .js
echo     - web_dashboard_direccion\dashboard_data.json / .js
echo     - Seguimiento_Presupuesto_Sep_2026.xlsx
echo     - Orden_Fabricacion_MRP_Codiagro.xlsx
echo ======================================================================

REM ----------------------------------------------------------------------
REM FASE 4: PUBLICACION SEPARADA (NO AUTOMATICA SALVO FLAG --publish)
REM ----------------------------------------------------------------------
if /I "%~1"=="--publish" (
    echo.
    echo [FASE 4] Publicando unicamente entregables autorizados de web_dashboard...
    git add web_dashboard/index.html web_dashboard/dashboard_data.json web_dashboard/dashboard_data.js web_dashboard/production_planning_data.json web_dashboard/production_planning_data.js
    git commit -m "Publicacion controlada web_dashboard: %date% %time%"
    git subtree push --prefix web_dashboard origin gh-pages
) else (
    echo.
    echo [INFO] Publicacion automatica desactivada por politica de control.
    echo        Los documentos internos, copias ^(scratch/^) y pruebas quedan fuera de publicacion.
    echo        Para publicar manualmente tras revisar, ejecute: actualizar_dashboard.bat --publish
)
exit /b 0

:ROLLBACK
echo.
echo ======================================================================
echo   [ERROR CRITICO] Fallo detectado en paso: !FAILED_STEP!
echo   Restaurando automaticamente la version previa valida desde %BACKUP_DIR%...
echo ======================================================================
set "CODIAGRO_STAGING_DIR="
"%PY%" -X utf8 scripts/clean_rollback.py restore
if !ERRORLEVEL! NEQ 0 (
    echo.
    echo   [ERROR GRAVE] La restauracion de la copia de seguridad ha fallado.
    exit /b 1
)
echo   [RESTAURADO] Se ha recuperado la version previa valida con exito y purgado huerfanos.
echo   [FECHA DATOS DISPONIBLES] %LAST_DATA_DATE%
echo ======================================================================
exit /b 1
