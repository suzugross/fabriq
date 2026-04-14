@echo off
setlocal enabledelayedexpansion

REM ===== Windows Update Auto-Resume Launcher (DEPRECATED) =====
REM This file is no longer used for WU resume.
REM WU now uses Register-FabriqRunOnce (Fabriq.exe) and main.ps1
REM detects wu_state.json to auto-resume the update loop.
REM Kept for reference only.

REM ===== administrator check =====
net session >nul 2>&1
if errorlevel 1 (
    powershell -Command "Start-Process cmd -ArgumentList '/c \"%~f0\"' -Verb RunAs"
    exit /b
)

cd /d %~dp0\..\..\..

start "" conhost.exe powershell.exe -NoProfile -ExecutionPolicy Unrestricted -Command ". .\kernel\common.ps1; $null = & '.\modules\standard\windows_update\windows_update.ps1'"

exit /b
