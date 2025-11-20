@echo off

REM ====== Rutas ======
set TARGET=%~dp0\scripts\ejecutar.bat
set ICON=%~dp0\icons\icono_mail.ico
set SHORTCUT_PATH=%USERPROFILE%\Desktop\Extractor de correos.lnk
set VBS=%TEMP%\shortcut.vbs

echo Set oWS = CreateObject("WScript.Shell") > "%VBS%"
echo sLinkFile = "%USERPROFILE%\Desktop\Extractor de correos.lnk" >> "%VBS%"
echo Set oLink = oWS.CreateShortcut(sLinkFile) >> "%VBS%"
echo oLink.TargetPath = "%CD%\scripts\ejecutar.bat" >> "%VBS%"
echo oLink.WorkingDirectory = "%CD%" >> "%VBS%"
echo oLink.IconLocation = "%CD%\icons\icono_mail.ico" >> "%VBS%"
echo oLink.Save >> "%VBS%"

cscript //nologo "%VBS%"
del "%VBS%"
