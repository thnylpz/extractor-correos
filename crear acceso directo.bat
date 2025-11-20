@echo off

REM ====== Ruta del ejecutable ======
set TARGET=%~dp0ejecutar.bat

REM ====== Rutas ======
set TARGET=%~dp0ejecutar.bat
set ICON=%~dp0icono_mail.ico
set SHORTCUT_PATH=%USERPROFILE%\Desktop\extractor de correos.lnk

(
echo Set oWS = WScript.CreateObject("WScript.Shell")
echo Set oLink = oWS.CreateShortcut("%SHORTCUT_PATH%")
echo oLink.TargetPath = "%TARGET%"
echo oLink.WorkingDirectory = "%~dp0"
echo oLink.IconLocation = "%ICON%"
echo oLink.Save
) > "%TEMP%\shortcut.vbs"

cscript //nologo "%TEMP%\shortcut.vbs" >nul
del "%TEMP%\shortcut.vbs"

cls
