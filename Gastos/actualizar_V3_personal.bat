@echo off
chcp 65001 > nul
echo ========================================================
echo Actualizando P&L V3 con los nuevos Gastos de Personal
echo ========================================================
copy /y "P&L mensual ene-ago V3 - Personal Actualizado.xlsx" "P&L mensual ene-ago V3.xlsx"
if %ERRORLEVEL% EQU 0 (
    echo.
    echo ¡P&L mensual ene-ago V3.xlsx actualizado con éxito!
) else (
    echo.
    echo Por favor cierra el archivo P&L mensual ene-ago V3.xlsx en Excel y vuelve a ejecutar este script.
)
pause
