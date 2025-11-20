@echo off
title Extractor de Correos - Actualizado y Fácil de Usar
color 0A

echo ============================================
echo      INICIANDO EXTRACTOR DE CORREOS
echo ============================================
echo.

REM --- Ir a la carpeta donde está el BAT ---
cd /d "%~dp0"

echo Verificando actualizaciones del programa...
git pull
echo.

REM --- Verificar si Python está instalado ---
echo Verificando instalacion de Python...
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Python no está instalado en esta PC.
    echo Por favor instala Python desde: https://www.python.org/downloads/
    pause
    exit /b
)

echo Python detectado correctamente.
echo.

REM --- Instalar dependencias ---
IF EXIST requirements.txt (
    echo Instalando librerías necesarias...
    python -m pip install --upgrade pip >nul
    python -m pip install -r requirements.txt
    echo.
) ELSE (
    echo No se encontro requirements.txt, continuando...
)

REM --- Ejecutar el script principal ---
echo Ejecutando el programa...
echo ============================================
echo.
@echo off
python extractor_de_correos.py

echo.
echo ============================================
echo Tarea finalizada. Puede cerrar esta ventana.
@echo off
pause

