@echo off
title NetserGroup - Actualizando Dashboard...
cd /d "%~dp0"
chcp 65001 >nul 2>&1

echo ========================================
echo   NetserGroup Dashboard - Actualizando
echo ========================================
echo.

echo [1/4] Generando dashboard desde Excel...
python generar_dashboard.py
if %errorlevel% neq 0 (
    echo ERROR: Fallo al generar dashboard. Verifica que Python y openpyxl esten instalados.
    pause
    exit /b 1
)
echo.

echo [2/4] Preparando Git...
git status >nul 2>&1
if %errorlevel% neq 0 (
    echo Inicializando repositorio Git...
    git init
    git remote add origin https://github.com/henriquezbh5-cpu/netsergroup-dashboard.git
)
echo.

echo [3/4] Agregando archivos y creando commit...
git add index.html generar_dashboard.py ACTUALIZAR_DASHBOARD.bat
git commit -m "Dashboard actualizado %date% %time%"
echo.

echo [4/4] Subiendo a GitHub...
git pull --rebase origin main 2>nul
git push -u origin main 2>&1

if %errorlevel% neq 0 (
    echo.
    echo PUSH FALLO - Intentando con force...
    git push -u origin main --force 2>&1
)

echo.
echo Abriendo dashboard...
start "" "https://henriquezbh5-cpu.github.io/netsergroup-dashboard/"
echo.
echo ========================================
echo   Listo! Espera 30-60 segundos y
echo   recarga con Ctrl+F5
echo ========================================
echo.
pause
