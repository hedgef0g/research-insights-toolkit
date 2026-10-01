@echo off
setlocal
set "SCRIPT=%~dp0install-windows.ps1"
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%"
echo.
if errorlevel 1 (
  echo Установщик завершился с ошибкой. Журнал: "%USERPROFILE%\Desktop\ResearchSignal-Installer.log"
) else (
  echo Установщик завершён. Журнал: "%USERPROFILE%\Desktop\ResearchSignal-Installer.log"
)
pause
