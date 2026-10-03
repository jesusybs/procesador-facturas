@echo off
chcp 65001 > nul
setlocal enabledelayedexpansion

if "%~1"=="" (
    echo ========================================================
    echo   PROCESADOR DETERMINISTA DE FACTURAS Y LISTAS DE EMPAQUE
    echo ========================================================
    echo.
    echo Arrastra un archivo PDF o Excel (.xlsx) sobre este archivo .bat
    echo o escribe la ruta del archivo a procesar:
    echo.
    set /p "ARCHIVO=Ruta del archivo: "
) else (
    set "ARCHIVO=%~1"
)

if not exist "!ARCHIVO!" (
    echo [ERROR] El archivo no existe: "!ARCHIVO!"
    pause
    exit /b 1
)

echo.
echo Procesando: "!ARCHIVO!" ...
echo.

python "C:\Facturas_Repuestos\procesador_facturas.py" "!ARCHIVO!" > "%~dpn1_resultado.md"
type "%~dpn1_resultado.md"

echo.
echo ========================================================
echo   Listo! Resultado guardado en:
echo   "%~dpn1_resultado.md"
echo ========================================================
echo.
pause
