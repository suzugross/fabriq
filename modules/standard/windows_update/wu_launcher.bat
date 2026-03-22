@echo off
setlocal enabledelayedexpansion

REM ===== Windows Update Auto-Resume Launcher =====
REM This batch file is registered in RunOnce to re-launch
REM the Windows Update module after a reboot.

REM ===== administrator check =====
net session >nul 2>&1
if errorlevel 1 (
    powershell -Command "Start-Process cmd -ArgumentList '/c \"%~f0\"' -Verb RunAs"
    exit /b
)

cd /d %~dp0\..\..\..

start "" conhost.exe powershell.exe -NoProfile -ExecutionPolicy Unrestricted -Command ". .\kernel\common.ps1; $null = & '.\modules\standard\windows_update\windows_update.ps1'"

exit /b
