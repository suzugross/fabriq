@echo off
rem ========================================
rem  Fabriq AD lab builder - launcher
rem  Run this ON the server that becomes the lab domain controller.
rem  LAB ONLY. Never run against production Active Directory.
rem ========================================
setlocal
cd /d "%~dp0"

net session >nul 2>&1
if errorlevel 1 (
    echo Administrator privileges are required - relaunching elevated...
    powershell.exe -NoProfile -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0ad_lab.ps1" %*
set RC=%ERRORLEVEL%

echo.
echo Exit code: %RC%
pause
exit /b %RC%
