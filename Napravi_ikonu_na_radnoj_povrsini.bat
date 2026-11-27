@echo off
rem Precac "OMR Ankete" na radnoj povrsini: otvara program na posluzitelju u pregledniku.
chcp 65001 >nul
set "DIR=%USERPROFILE%\OMR Ankete"
if not exist "%DIR%" mkdir "%DIR%"
copy /Y "%~dp0omr-ankete.ico" "%DIR%\omr-ankete.ico" >nul
for /f "usebackq delims=" %%D in (`powershell -NoProfile -Command "[Environment]::GetFolderPath('Desktop')"`) do set "DESK=%%D"
(
echo [InternetShortcut]
echo URL=https://ankete.92-5-74-253.sslip.io/
echo IconFile=%DIR%\omr-ankete.ico
echo IconIndex=0
) > "%DESK%\OMR Ankete.url"
echo.
echo Gotovo: na radnoj povrsini je ikona "OMR Ankete".
echo.
pause
