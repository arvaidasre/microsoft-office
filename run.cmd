@echo off
:: Check for admin rights and elevate if necessary
NET SESSION >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    powershell -Command "Start-Process -FilePath '%~dpnx0' -Verb RunAs" >nul 2>&1
    exit /b
)
:: Continue with admin privileges silently
echo Running Office installation script...
powershell.exe -ExecutionPolicy Bypass -File "%~dp0install_office_auto.ps1"
echo Installation complete.
