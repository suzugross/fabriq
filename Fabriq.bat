@echo off
REM ============================================================
REM Fabriq Launcher (batch variant)
REM Parallel alternative to Fabriq.exe. Provided so the framework
REM can still be launched when antivirus heuristics quarantine the
REM unsigned EXE launcher. Behaviour mirrors Fabriq.exe:
REM   - self-elevates to Administrator (the EXE used a
REM     requireAdministrator manifest)
REM   - pins the working directory to the fabriq root
REM   - starts kernel\main.ps1 via conhost + powershell
REM This file does NOT replace Fabriq.exe; both can coexist.
REM ============================================================
setlocal

REM ----- Administrator check / self-elevation -----
net session >nul 2>&1
if errorlevel 1 (
    echo [*] Elevation required. Requesting Administrator...
    powershell -NoProfile -Command "Start-Process cmd -ArgumentList '/c \"%~f0\"' -Verb RunAs"
    exit /b
)

echo [+] Running with Administrator privileges.

REM ----- Pin working directory to the fabriq root (where this .bat lives).
REM       main.ps1 resolves relative paths like .\kernel\common.ps1. -----
cd /d "%~dp0"

if not exist ".\kernel\main.ps1" (
    echo [!] kernel\main.ps1 was not found under "%~dp0".
    echo     Run this launcher from the fabriq root directory.
    pause
    exit /b 2
)

REM ----- Launch the kernel (conhost gives reliable console window sizing;
REM       Windows Terminal ignores "mode con"). Matches Fabriq.exe. -----
start "" conhost.exe powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\kernel\main.ps1"

exit /b
