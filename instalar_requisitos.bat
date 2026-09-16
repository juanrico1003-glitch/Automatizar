@echo off
cd /d "%~dp0"
chcp 65001 >nul 2>&1
title Instalador de Requisitos - Automatizacion CXP

echo ======================================================
echo    INSTALADOR DE REQUISITOS (AUTOMATIZACION CXP)
echo ======================================================
echo.

echo [1/3] Detectando instalacion de Python en este equipo...

set "PYTHON_CMD="

REM 1. Probar lanzador oficial 'py -3' (instalado por defecto en C:\Windows\py.exe)
py -3 -c "import sys; exit(0)" >nul 2>&1
if %errorlevel% equ 0 (
    set "PYTHON_CMD=py -3"
    goto :python_encontrado
)

REM 2. Probar ruta de instalacion estandar de Python 3.14 (usuario)
if exist "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" (
    "%LOCALAPPDATA%\Programs\Python\Python314\python.exe" -c "import sys; exit(0)" >nul 2>&1
    if %errorlevel% equ 0 (
        set "PYTHON_CMD="%LOCALAPPDATA%\Programs\Python\Python314\python.exe""
        goto :python_encontrado
    )
)

REM 3. Probar 'python' en PATH (verificando que funcione y NO sea el alias vacio de Microsoft Store)
python -c "import sys; exit(0)" >nul 2>&1
if %errorlevel% equ 0 (
    set "PYTHON_CMD=python"
    goto :python_encontrado
)

REM 4. Probar otras versiones comunes de Python en AppData del usuario
for %%V in (Python313 Python312 Python311 Python310) do (
    if exist "%LOCALAPPDATA%\Programs\Python\%%V\python.exe" (
        "%LOCALAPPDATA%\Programs\Python\%%V\python.exe" -c "import sys; exit(0)" >nul 2>&1
        if %errorlevel% equ 0 (
            set "PYTHON_CMD="%LOCALAPPDATA%\Programs\Python\%%V\python.exe""
            goto :python_encontrado
        )
    )
)

REM 5. Probar rutas de Program Files (instalaciones para todos los usuarios)
for %%V in (Python314 Python313 Python312 Python311 Python310) do (
    if exist "%ProgramFiles%\Python%%V\python.exe" (
        "%ProgramFiles%\Python%%V\python.exe" -c "import sys; exit(0)" >nul 2>&1
        if %errorlevel% equ 0 (
            set "PYTHON_CMD="%ProgramFiles%\Python%%V\python.exe""
            goto :python_encontrado
        )
    )
    if exist "%SystemDrive%\Python%%V\python.exe" (
        "%SystemDrive%\Python%%V\python.exe" -c "import sys; exit(0)" >nul 2>&1
        if %errorlevel% equ 0 (
            set "PYTHON_CMD="%SystemDrive%\Python%%V\python.exe""
            goto :python_encontrado
        )
    )
)

REM 6. Probar comando generico 'py'
py -c "import sys; exit(0)" >nul 2>&1
if %errorlevel% equ 0 (
    set "PYTHON_CMD=py"
    goto :python_encontrado
)

REM Si llegamos aqui, no se encontro ningun interprete funcional de Python
echo.
echo [AVISO] No se encontro una instalacion funcional de Python en el sistema.
echo (O la instalacion actual no agrego Python al PATH ni al iniciador).
echo.
echo Opciones disponibles:
echo  1) Usar el archivo "Organizador_CXP.exe" incluido en esta carpeta.
echo     (NO requiere instalar Python ni configurar nada, funciona directo).
echo.
echo  2) Deseas que intentemos instalar Python automaticamente ahora mismo con winget?
set /p INSTALAR_SN="   Instalar Python automaticamente? (S/N): "
if /i "%INSTALAR_SN%"=="S" (
    echo Instalando Python via winget...
    winget install Python.Python.3.14 --accept-package-agreements --accept-source-agreements
    echo.
    echo Instalacion finalizada. Por favor cierra y vuelve a abrir este instalador.
    pause
    exit /b 0
)

echo.
echo Si prefieres instalarlo manualmente:
echo Descargalo de: https://www.python.org/downloads/
echo RECUERDA MARCAR: "Add python.exe to PATH"
echo.
pause
exit /b 1

:python_encontrado
echo [OK] Python detectado: %PYTHON_CMD%
%PYTHON_CMD% --version
echo.

echo [2/3] Verificando entorno virtual (.venv)...

REM Si existe la carpeta .venv pero no funciona (ej. copiada de otra PC), eliminarla
if exist ".venv" (
    ".venv\Scripts\python.exe" -c "import sys; exit(0)" >nul 2>&1
    if %errorlevel% neq 0 (
        echo [AVISO] Se detecto un entorno virtual antiguo o copiado de otra PC.
        echo Limpiando y regenerando para este equipo...
        rmdir /s /q ".venv" >nul 2>&1
    )
)

REM Crear entorno virtual si no existe
if not exist ".venv\Scripts\python.exe" (
    echo Creando entorno virtual local...
    %PYTHON_CMD% -m venv .venv
    if not exist ".venv\Scripts\python.exe" (
        echo.
        echo [ERROR] No se pudo crear el entorno virtual con %PYTHON_CMD%.
        pause
        exit /b 1
    )
    echo [OK] Entorno virtual creado exitosamente.
) else (
    echo [OK] Entorno virtual verificado y funcional.
)
echo.

echo [3/3] Instalando / verificando librerias necesarias...
".venv\Scripts\python.exe" -m pip install --upgrade pip >nul 2>&1
if exist "requirements.txt" (
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
) else (
    ".venv\Scripts\python.exe" -m pip install openpyxl
)

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Hubo un problema instalando las librerias con pip.
    pause
    exit /b 1
)

echo [OK] Librerias instaladas correctamente.
echo.

REM Preguntar si desea crear acceso directo en el Escritorio
echo Deseas crear un acceso directo en el Escritorio de esta PC?
set /p CREAR_ACCESO="Crear acceso directo? (S/N): "
if /i "%CREAR_ACCESO%"=="S" (
    powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws = New-Object -ComObject WScript.Shell; $d = [Environment]::GetFolderPath('Desktop'); $s = $ws.CreateShortcut([System.IO.Path]::Combine($d, 'Organizador CXP.lnk')); $s.TargetPath = '%~dp0ejecutar_automatizacion.bat'; $s.WorkingDirectory = '%~dp0'; $s.Save()" >nul 2>&1
    echo [OK] Acceso directo 'Organizador CXP' creado en el Escritorio.
)

echo.
echo ======================================================
echo    ¡INSTALACION Y CONFIGURACION COMPLETADA!
echo ======================================================
echo Ya puedes procesar archivos ejecutando "ejecutar_automatizacion.bat"
echo o simplemente usando "Organizador_CXP.exe".
echo.
pause