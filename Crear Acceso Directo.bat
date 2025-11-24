@echo off

REM ====== Rutas ======
set TARGET=%~dp0\scripts\ejecutar.bat
set ICON=%~dp0\icons\icono_mail.ico
set SHORTCUT_PATH=%USERPROFILE%\Desktop\Extractor de correos.lnk
@REM set VBS=%TEMP%\shortcut.vbs

@REM Convertir rutas a formato PowerShell seguro
for %%A in ("%TARGET%") do set "PTARGET=%%~fA"
for %%A in ("%ICON%") do set "PICON=%%~fA"
for %%A in ("%SHORTCUT_PATH%") do set "PSHORTCUT=%%~fA"

@REM @echo Creando acceso directo...

powershell -NoProfile -Command ^
    "$W = New-Object -ComObject WScript.Shell; " ^
    "$S = $W.CreateShortcut('%PSHORTCUT%'); " ^
    "$S.TargetPath = '%PTARGET%'; " ^
    "$S.WorkingDirectory = (Split-Path '%PTARGET%'); " ^
    "$S.IconLocation = '%PICON%'; " ^
    "$S.Save()"

@REM cls
@echo Acceso directo creado correctamente.
@echo.

@REM echo Set oWS = CreateObject("WScript.Shell") > "%VBS%"
@REM echo sLinkFile = "%USERPROFILE%\Desktop\Extractor de correos.lnk" >> "%VBS%"
@REM echo Set oLink = oWS.CreateShortcut(sLinkFile) >> "%VBS%"
@REM echo oLink.TargetPath = "%CD%\scripts\ejecutar.bat" >> "%VBS%"
@REM echo oLink.WorkingDirectory = "%CD%" >> "%VBS%"
@REM echo oLink.IconLocation = "%CD%\icons\icono_mail.ico" >> "%VBS%"
@REM echo oLink.Save >> "%VBS%"

@REM cscript //nologo "%VBS%"
@REM del "%VBS%"

pause