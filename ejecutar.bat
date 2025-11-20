@echo off
cd /d "%~dp0"
git pull
python extractor_de_correos.py
pause