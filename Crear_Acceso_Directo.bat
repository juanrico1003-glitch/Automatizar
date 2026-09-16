@echo off
cd /d "%~dp0"
chcp 65001 >nul 2>&1
title Crear Acceso Directo - Organizador CXP

echo ======================================================
echo    CREAR ACCESO DIRECTO EN EL ESCRITORIO
echo ======================================================
echo.

set "TARGET=%~dp0ejecutar_automatizacion.bat"
if exist "%~dp0Organizador_CXP.exe" (
    set "TARGET=%~dp0Organizador_CXP.exe"
)

powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws = New-Object -ComObject WScript.Shell; $d = [Environment]::GetFolderPath('Desktop'); $s = $ws.CreateShortcut([System.IO.Path]::Combine($d, 'Organizador CXP.lnk')); $s.TargetPath = '%TARGET%'; $s.WorkingDirectory = '%~dp0'; $s.Save()"

if %errorlevel% equ 0 (
    echo [OK] ¡Acceso directo creado con exito en tu Escritorio!
) else (
    echo [ERROR] No se pudo crear el acceso directo automaticamente.
)

echo.
pause
