@echo off
title Asignador de Calificaciones SIIUPP
echo ===================================================
echo     Iniciando Asignador de Calificaciones SIIUPP
echo ===================================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo [!] No se encontro el entorno virtual en .venv.
    echo Creando entorno virtual e instalando librerias...
    python -m venv .venv
    .venv\Scripts\python.exe -m pip install -r requirements.txt
)

echo Iniciando aplicacion...
start "" ".venv\Scripts\pythonw.exe" app.py
exit

