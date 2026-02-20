@echo off
setlocal enabledelayedexpansion

REM ===== administrator check =====
echo [*] administrator check...

net session >nul 2>&1
if errorlevel 1 (
    echo [!] need to administrator, elevating...
    powershell -Command "Start-Process cmd -ArgumentList '/c \"%~f0\"' -Verb RunAs"
    exit /b
)

echo [+] administrator mode is already

REM Set CWD to fabriq root (2 levels up from profiles/<name>/)
cd /d "%~dp0..\.."

start "" conhost.exe powershell.exe -NoProfile -ExecutionPolicy Unrestricted -File "%~dp0easyprofile.ps1"

exit /b
