@echo off
chcp 65001 > nul
title Instalador de Requisitos - Automatización CXP

echo ======================================================
echo    INSTALADOR DE REQUISITOS (AUTOMATIZACION CXP)
echo ======================================================
echo.

cd /d "%~dp0"

echo [1/3] Comprobando Python en este equipo...
where python >nul 2>nul
if %errorlevel% equ 0 (
    set "PYTHON_CMD=python"
    goto :python_detectado
)

where py >nul 2>nul
if %errorlevel% equ 0 (
    set "PYTHON_CMD=py"
    goto :python_detectado
)

echo.
echo [ERROR] No se encontró Python en este equipo.
echo.
echo Por favor instala Python descargándolo desde:
echo https://www.python.org/downloads/
echo.
echo IMPORTANTE AL INSTALAR: Marca la casilla:
echo "Add python.exe to PATH"
echo.
pause
exit /b 1

:python_detectado
%PYTHON_CMD% --version
echo [OK] Python detectado correctamente.
echo.

echo [2/3] Verificando / creando entorno virtual local (.venv)...
if not exist ".venv\Scripts\python.exe" (
    echo Creando entorno virtual con %PYTHON_CMD%...
    %PYTHON_CMD% -m venv .venv
    if not exist ".venv\Scripts\python.exe" (
        echo.
        echo [ERROR] No se pudo crear el entorno virtual.
        pause
        exit /b 1
    )
    echo [OK] Entorno virtual creado exitosamente.
) else (
    echo [OK] El entorno virtual ya existe.
)
echo.

echo [3/3] Instalando dependencias (openpyxl)...
".venv\Scripts\python.exe" -m pip install --upgrade pip >nul 2>nul
if exist "requirements.txt" (
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
) else (
    ".venv\Scripts\python.exe" -m pip install openpyxl
)

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Ocurrió un problema instalando las librerías.
    pause
    exit /b 1
)

echo [OK] Librerías instaladas correctamente.
echo.
echo ======================================================
echo    ¡TODO LISTO PARA USAR!
echo ======================================================
echo Ya puedes usar "ejecutar_automatizacion.bat" en esta PC.
echo.
pause