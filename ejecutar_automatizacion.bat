@echo off
cd /d "%~dp0"
chcp 65001 >nul 2>&1
title Organizador Automatico de CXP - Excel

echo ======================================================
echo    ORGANIZADOR AUTOMATICO DE ARCHIVOS CXP (EXCEL)
echo ======================================================
echo.

REM 1. Si existe Organizador_CXP.exe y no hay entorno .venv, ejecutar el .exe directo
if not exist ".venv\Scripts\python.exe" (
    if exist "Organizador_CXP.exe" (
        if "%~1"=="" (
            "Organizador_CXP.exe"
        ) else (
            "Organizador_CXP.exe" "%~1"
        )
        goto fin
    )
)

REM 2. Si existe .venv, verificar si es funcional o si fue copiado de otra PC
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" -c "import sys; exit(0)" >nul 2>&1
    if %errorlevel% neq 0 (
        echo [AVISO] El entorno virtual de esta carpeta necesita configurarse para este equipo.
        call instalar_requisitos.bat
    )
) else (
    echo [AVISO] Configurando el entorno por primera vez...
    call instalar_requisitos.bat
)

if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] No se pudo iniciar el entorno de Python.
    pause
    exit /b 1
)

REM 3. Ejecutar script Python
if "%~1"=="" (
    echo Procesando archivos Excel en esta carpeta...
    ".venv\Scripts\python.exe" organizar_cxp.py
) else (
    echo Procesando archivo: "%~1"
    ".venv\Scripts\python.exe" organizar_cxp.py "%~1"
)

:fin
echo.
echo ======================================================
echo    Proceso finalizado.
echo ======================================================
echo.
pause