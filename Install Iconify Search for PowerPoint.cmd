@echo off
setlocal EnableExtensions
title Install Iconify Search for PowerPoint

set "ADDIN_ID=4d5bf4c8-7851-4f6f-b595-66fc244898a4"
set "SOURCE=%~dp0manifest.xml"
set "TARGET_DIR=%LOCALAPPDATA%\Iconify Search for PowerPoint"
set "TARGET=%TARGET_DIR%\manifest.xml"
set "DEV_KEY=HKCU\SOFTWARE\Microsoft\Office\16.0\Wef\Developer"
set "OFFICE_EXE="

echo.
echo  Iconify Search for PowerPoint
echo  --------------------------------
echo.

if not exist "%SOURCE%" goto missing

tasklist /FI "IMAGENAME eq POWERPNT.EXE" 2>nul | find /I "POWERPNT.EXE" >nul
if not errorlevel 1 goto running

for /f "skip=2 tokens=2,*" %%A in ('reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE" /ve 2^>nul') do if /I "%%A"=="REG_SZ" set "OFFICE_EXE=%%B"
if not defined OFFICE_EXE for /f "skip=2 tokens=2,*" %%A in ('reg query "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE" /ve 2^>nul') do if /I "%%A"=="REG_SZ" set "OFFICE_EXE=%%B"
if not defined OFFICE_EXE if exist "%ProgramFiles%\Microsoft Office\root\Office16\POWERPNT.EXE" set "OFFICE_EXE=%ProgramFiles%\Microsoft Office\root\Office16\POWERPNT.EXE"
if not defined OFFICE_EXE if exist "%ProgramFiles(x86)%\Microsoft Office\root\Office16\POWERPNT.EXE" set "OFFICE_EXE=%ProgramFiles(x86)%\Microsoft Office\root\Office16\POWERPNT.EXE"
if not defined OFFICE_EXE goto nooffice

if not exist "%ProgramFiles%\Microsoft\EdgeWebView\Application" if not exist "%ProgramFiles(x86)%\Microsoft\EdgeWebView\Application" if not exist "%LOCALAPPDATA%\Microsoft\EdgeWebView\Application" goto webview

if not exist "%TARGET_DIR%" mkdir "%TARGET_DIR%"
if errorlevel 1 goto failed
copy /Y "%SOURCE%" "%TARGET%" >nul
if errorlevel 1 goto failed

reg add "%DEV_KEY%" /v "%ADDIN_ID%" /t REG_SZ /d "%TARGET%" /f >nul
if errorlevel 1 goto failed
reg add "%DEV_KEY%" /v "RefreshAddins" /t REG_DWORD /d 1 /f >nul
if errorlevel 1 goto failed

echo  Installation complete!
echo.
echo  This installer will not open PowerPoint automatically.
echo.
echo  1. Close this installer.
echo  2. Open PowerPoint yourself.
echo  3. Choose Home ^> Iconify ^> Search Icons.
echo.
echo  Installed in:
echo  %TARGET_DIR%
echo.
pause
exit /b 0

:running
echo  INSTALLATION PAUSED
echo.
echo  Close every PowerPoint window, then run this installer again.
goto end_error

:missing
echo  INSTALLATION COULD NOT START
echo.
echo  Extract the complete ZIP and keep manifest.xml beside this installer.
goto end_error

:nooffice
echo  MICROSOFT POWERPOINT WAS NOT FOUND
echo.
echo  Install Microsoft 365 or PowerPoint, then run this installer again.
goto end_error

:webview
echo  MICROSOFT EDGE WEBVIEW2 IS REQUIRED
echo.
echo  Download it directly from Microsoft, install it, then rerun this installer:
echo  https://developer.microsoft.com/microsoft-edge/webview2/
goto end_error

:failed
echo  INSTALLATION FAILED
echo.
echo  Windows could not copy or register the add-in for this user.

:end_error
echo.
pause
exit /b 1
