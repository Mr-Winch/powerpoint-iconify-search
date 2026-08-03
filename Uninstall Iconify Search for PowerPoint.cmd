@echo off
setlocal EnableExtensions
title Uninstall Iconify Search for PowerPoint

set "ADDIN_ID=4d5bf4c8-7851-4f6f-b595-66fc244898a4"
set "TARGET_DIR=%LOCALAPPDATA%\Iconify Search for PowerPoint"
set "DEV_KEY=HKCU\SOFTWARE\Microsoft\Office\16.0\Wef\Developer"

echo.
echo  Uninstall Iconify Search for PowerPoint
echo  ---------------------------------------
echo.

tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>nul | find /I "POWERPNT.EXE" >nul
if not errorlevel 1 (
  echo  Close every PowerPoint window, then run this uninstaller again.
  echo.
  pause
  exit /b 1
)

reg delete "%DEV_KEY%" /v "%ADDIN_ID%" /f >nul 2>nul
reg add "%DEV_KEY%" /v "RefreshAddins" /t REG_DWORD /d 1 /f >nul
if exist "%TARGET_DIR%\manifest.xml" del /Q "%TARGET_DIR%\manifest.xml"
if exist "%TARGET_DIR%" rd "%TARGET_DIR%" 2>nul

echo  Uninstallation complete.
echo.
pause
exit /b 0
