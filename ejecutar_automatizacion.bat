@echo off
chcp 65001 > nul
title Organizador Automático de CXP - Excel

echo ======================================================
echo    ORGANIZADOR AUTOMATICO DE ARCHIVOS CXP (EXCEL)
echo ======================================================
echo.

cd /d "%~dp0"

REM Si no existe el entorno virtual, intentar crearlo con el instalador
if exist ".venv\Scripts\python.exe" goto venv_ok

echo [AVISO] No se encontró el entorno virtual local.
echo Configurando el entorno por primera vez...
call instalar_requisitos.bat
if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] No se pudo preparar el entorno virtual. Abortando.
    pause
    exit /b 1
)

:venv_ok
if "%~1"=="" goto procesar_todos

echo Procesando archivo: "%~1"
".venv\Scripts\python.exe" organizar_cxp.py "%~1"
goto fin

:procesar_todos
echo Procesando archivos Excel en esta carpeta...
".venv\Scripts\python.exe" organizar_cxp.py

:fin
echo.
echo ======================================================
echo    Proceso finalizado.
echo ======================================================
echo.
pause