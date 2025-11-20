@echo off

REM ====== Rutas ======
set TARGET=%~dp0ejecutar.bat
set ICON=%~dp0icono_mail.ico
set SHORTCUT_PATH=%USERPROFILE%\Desktop\Extractor de correos.lnk
set VBS=%TEMP%\shortcut.vbs

echo Set oWS = CreateObject("WScript.Shell") > "%VBS%"
echo sLinkFile = "%USERPROFILE%\Desktop\Extractor de correos.lnk" >> "%VBS%"
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> "%VBS%"
echo oLink.TargetPath = "%CD%\ejecutar.bat" >> "%VBS%"
echo oLink.WorkingDirectory = "%CD%" >> "%VBS%"
echo oLink.IconLocation = "%CD%\icon\icono_mail.ico" >> "%VBS%"
echo oLink.Save >> "%VBS%"

cscript //nologo "%VBS%"
del "%VBS%"