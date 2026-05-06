@echo off
:: ============================================================
::  Activa IT — Descargador de Cartas Glosa Previsora
::  Script de instalación y ejecución (Windows)
:: ============================================================

setlocal
set "SCRIPT_DIR=%~dp0"
set "VENV_DIR=%SCRIPT_DIR%venv"

echo.
echo ============================================================
echo   Activa IT -- Descargador de Cartas Glosa Previsora
echo ============================================================
echo.

:: Verificar Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python no encontrado. Descargalo de https://python.org
    pause
    exit /b 1
)

:: Crear entorno virtual
if not exist "%VENV_DIR%" (
    echo Creando entorno virtual Python...
    python -m venv "%VENV_DIR%"
)

:: Activar entorno virtual
call "%VENV_DIR%\Scripts\activate.bat"

:: Instalar dependencias
echo Instalando dependencias...
pip install --quiet flask playwright

:: Instalar Chromium
echo Instalando Playwright Chromium...
python -m playwright install chromium

:: Crear carpeta descargas
if not exist "%SCRIPT_DIR%downloads" mkdir "%SCRIPT_DIR%downloads"

:: Abrir navegador
echo.
echo Iniciando servidor...
start "" "http://localhost:5050"

:: Ejecutar app
python "%SCRIPT_DIR%app.py"

pause
