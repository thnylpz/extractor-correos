@echo off

REM --- Ir a la carpeta donde está el BAT ---
cd /d "%~dp0\..\src"

REM --- Actualizar desde Github ---
git pull
cls

REM --- Verificar si Python está instalado ---
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo ERROR: Python no esta instalado en esta PC.
    echo Por favor instala Python desde: https://www.python.org/downloads/
    pause
    exit /b
)

echo Python detectado correctamente.
echo.
cls

REM --- Instalar dependencias ---
echo Instalando dependencias
echo.
IF EXIST requirements.txt (
    python -m pip install --upgrade pip >nul
    python -m pip install -r requirements.txt
)
cls

REM --- Ejecutar el script principal ---
python extractor_de_correos.py

pause
