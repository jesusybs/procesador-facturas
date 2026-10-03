@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

echo ========================================================
echo   PROCESANDO TODOS LOS ARCHIVOS EN LA CARPETA ENTRADA
echo ========================================================
echo.

python "C:\Facturas_Repuestos\procesador_facturas.py" --carpeta "C:\Facturas_Repuestos\Entrada" "C:\Facturas_Repuestos\Salida"

echo.
echo ========================================================
echo   Proceso finalizado. 
echo   Los resultados se encuentran en: C:\Facturas_Repuestos\Salida\
echo ========================================================
echo.
pause
