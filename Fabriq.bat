@echo off
setlocal enabledelayedexpansion

REM ===== administrator check =====
echo [*] administrator check...

REM administrator check
net session >nul 2>&1
if errorlevel 1 (
    echo [!] need to administrator, update ...

    powershell -Command "Start-Process cmd -ArgumentList '/c \"%~f0\"' -Verb RunAs"
    exit /b
)

echo [+] administrator mode is alrady

cd /d %~dp0

REM Launch in conhost for reliable window size control (Windows Terminal ignores mode con)
start "" conhost.exe powershell.exe -NoProfile -ExecutionPolicy Unrestricted .\kernel\main.ps1

exit /b
