@echo off
title NetserGroup - Actualizando Dashboard...
cd /d "%~dp0"
chcp 65001 >nul 2>&1

echo ========================================
echo   NetserGroup Dashboard - Actualizando
echo ========================================
echo.

echo [1/4] Generando dashboard y widget desde Excel...
python generar_dashboard.py
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
git add index.html widget.html generar_dashboard.py ACTUALIZAR_DASHBOARD.bat manifest.json sw.js
git commit -m "Dashboard y widget actualizado %date% %time%"
echo.

echo [4/4] Subiendo a GitHub...
git branch -M main
git push -u origin main 2>&1
echo.

if %errorlevel% neq 0 (
    echo.
    echo PUSH FALLO - Intentando con force...
    git push -u origin main --force 2>&1
)

echo.
echo Abriendo dashboard...
start "" "https://henriquezbh5-cpu.github.io/netsergroup-dashboard/"
echo.
echo Widget movil disponible en:
echo https://henriquezbh5-cpu.github.io/netsergroup-dashboard/widget.html
echo.
echo ========================================
echo   Listo! Espera 30-60 segundos y
echo   recarga con Ctrl+F5
echo ========================================
echo.
echo Para instalar el widget en tu celular:
echo   1. Abre el link del widget en Chrome
echo   2. Toca los 3 puntos (menu)
echo   3. Selecciona "Agregar a pantalla de inicio"
echo.
pause
