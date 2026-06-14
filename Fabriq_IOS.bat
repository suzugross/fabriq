@echo off
REM ============================================================
REM Fabriq IOS Launcher (batch variant)
REM Parallel alternative to Fabriq_IOS.exe. Provided so the joke
REM shell can still be launched when antivirus heuristics quarantine
REM the unsigned EXE launcher. Behaviour mirrors Fabriq_IOS.exe:
REM   - self-elevates to Administrator (the EXE used a
REM     requireAdministrator manifest)
REM   - pins the working directory to the fabriq root
REM   - starts apps\fabriq_ios\fabriq_ios.ps1 as a MINIMIZED bootstrap
REM     console. fabriq_ios.ps1 self-spawns the interactive child shell
REM     (FABRIQ_IOS_SUBPROCESS sentinel) in its own normal, focused
REM     window, then the bootstrap just blocks - so it is minimized to
REM     avoid leaving a blank idle window beside the shell.
REM This file does NOT replace Fabriq_IOS.exe; both can coexist.
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

REM ----- Pin working directory to the fabriq root (where this .bat lives). -----
cd /d "%~dp0"

if not exist ".\apps\fabriq_ios\fabriq_ios.ps1" (
    echo [!] apps\fabriq_ios\fabriq_ios.ps1 was not found under "%~dp0".
    echo     Run this launcher from the fabriq root directory.
    pause
    exit /b 2
)

REM ----- Launch the bootstrap minimized (it only self-spawns the child
REM       shell and then blocks); the child shell opens visible/focused. -----
start "" /min conhost.exe powershell.exe -NoProfile -ExecutionPolicy Bypass -File ".\apps\fabriq_ios\fabriq_ios.ps1"

exit /b
