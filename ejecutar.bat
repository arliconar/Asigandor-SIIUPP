@echo off
cd /d "%~dp0"
title Asignador de Calificaciones SIIUPP

echo ===================================================
echo     Iniciando Asignador de Calificaciones SIIUPP
echo ===================================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo [!] Configurando entorno virtual por primera vez...
    python -m venv .venv
    .venv\Scripts\python.exe -m pip install -r requirements.txt
)

echo Abriendo aplicacion...
".venv\Scripts\python.exe" app.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [!] Hubo un detalle al cerrar la aplicacion.
    pause
)
