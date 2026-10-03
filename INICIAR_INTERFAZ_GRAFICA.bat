@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

echo ===================================================================
echo   INICIANDO INTERFAZ GRÁFICA DE FACTURAS Y LISTAS DE EMPAQUE
echo ===================================================================
echo.
echo Abriendo la aplicación en tu navegador web en: http://127.0.0.1:5000
echo.
echo Para cerrar la aplicación en cualquier momento, cierra esta ventana.
echo ===================================================================
echo.

python "C:\Facturas_Repuestos\app\app.py"
pause
