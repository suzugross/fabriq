@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul 2>&1

REM ========================================
REM Fabriq Deploy Script
REM USBメモリからPCへFabriqフォルダを展開する
REM ========================================

echo.
echo ========================================
echo  Fabriq Deploy Tool
echo ========================================
echo.

REM ===== Administrator Check =====
net session >nul 2>&1
if errorlevel 1 (
    echo [!] Administrator privileges required. Elevating...
    powershell -Command "Start-Process cmd -ArgumentList '/c \"%~f0\"' -Verb RunAs"
    exit /b
)
echo [+] Running as Administrator
echo.

REM ===== Determine source drive (where this script resides) =====
set "SOURCE_DRIVE=%~d0"
set "SOURCE_DIR=%~dp0"
set "FABRIQ_SRC=%SOURCE_DIR%fabriq"

REM If Deploy.bat is inside the fabriq folder itself, use parent
if not exist "%FABRIQ_SRC%\kernel\main.ps1" (
    REM Check if we're inside the fabriq folder
    if exist "%SOURCE_DIR%kernel\main.ps1" (
        set "FABRIQ_SRC=%SOURCE_DIR:~0,-1%"
    ) else (
        echo [ERROR] fabriq folder not found.
        echo   Expected: %FABRIQ_SRC%
        echo   Or Deploy.bat should be inside the fabriq folder.
        pause
        exit /b 1
    )
)

echo [INFO] Source: %FABRIQ_SRC%
echo [INFO] Source Drive: %SOURCE_DRIVE%
echo.

REM ===== Get Volume Serial Number of source drive =====
set "VOL_SERIAL="
for /f "tokens=*" %%a in ('vol %SOURCE_DRIVE% 2^>nul') do (
    set "VOL_LINE=%%a"
)
REM Extract serial number (format: "XXXX-XXXX" at end of second line)
for /f "tokens=*" %%a in ('vol %SOURCE_DRIVE% 2^>nul ^| findstr /R "[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]-[0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f][0-9A-Fa-f]"') do (
    for %%b in (%%a) do set "VOL_SERIAL=%%b"
)

if "%VOL_SERIAL%"=="" (
    REM Fallback: use WMI to get volume serial
    for /f "skip=1 tokens=*" %%a in ('wmic logicaldisk where "DeviceID='%SOURCE_DRIVE%'" get VolumeSerialNumber 2^>nul') do (
        if not "%%a"=="" (
            set "VOL_SERIAL=%%a"
            goto :got_serial
        )
    )
)
:got_serial

REM Trim whitespace
for /f "tokens=*" %%a in ("%VOL_SERIAL%") do set "VOL_SERIAL=%%a"

if "%VOL_SERIAL%"=="" (
    echo [WARNING] Could not retrieve volume serial number.
    set "VOL_SERIAL=UNKNOWN"
)

echo [INFO] Volume Serial: %VOL_SERIAL%
echo.

REM ===== Save volume serial to source_media.id =====
echo %VOL_SERIAL%> "%FABRIQ_SRC%\kernel\source_media.id"
echo [SUCCESS] Saved source_media.id: %VOL_SERIAL%
echo.

REM ===== Select destination =====
set "DEFAULT_DEST=C:\Windows\work\fabriq"

echo Destination folder:
echo   [1] %DEFAULT_DEST% (Default)
echo   [2] Custom path
echo.
set /p "DEST_CHOICE=Select [1]: "
if "%DEST_CHOICE%"=="" set "DEST_CHOICE=1"

if "%DEST_CHOICE%"=="1" (
    set "DEST_DIR=%DEFAULT_DEST%"
) else if "%DEST_CHOICE%"=="2" (
    set /p "DEST_DIR=Enter destination path: "
) else (
    set "DEST_DIR=%DEFAULT_DEST%"
)

echo.
echo ========================================
echo  Deploy Summary
echo ========================================
echo   Source:      %FABRIQ_SRC%
echo   Destination: %DEST_DIR%
echo   Media ID:    %VOL_SERIAL%
echo ========================================
echo.

set /p "CONFIRM=Proceed with deployment? (Y/N): "
if /i not "%CONFIRM%"=="Y" (
    echo [INFO] Deployment cancelled.
    pause
    exit /b 0
)

echo.

REM ===== Copy fabriq folder to destination =====
if exist "%DEST_DIR%" (
    echo [INFO] Destination exists - refreshing framework/config from source.
    echo [INFO] Preserved: evidence, logs, and in-progress resume/WU state.
    echo [INFO] Removed: stale files no longer present in the source.
) else (
    echo [INFO] Creating destination folder...
    mkdir "%DEST_DIR%"
)

echo [INFO] Copying files...
REM /MIR mirrors the source and DELETES dest files absent from the source.
REM Runtime artifacts are generated on the target PC and are NOT in the source,
REM so a re-deploy to an already-kitted PC would otherwise purge them. Exclude
REM them so re-deploying never destroys delivery evidence or in-progress state:
REM   /XD evidence logs : preserve evidence\ (delivery data) + logs\ (logs,
REM                       telemetry, execution_history). Source paths are used
REM                       because robocopy matches /XD source-rooted; a dest
REM                       absolute path does NOT protect against the /MIR purge.
REM   /XF <state files> : preserve mid-kitting resume / Windows Update state and
REM                       the per-environment telemetry salt.
robocopy "%FABRIQ_SRC%" "%DEST_DIR%" /MIR /XD "%FABRIQ_SRC%\evidence" "%FABRIQ_SRC%\logs" /XF resume_state.json wu_state.json wu_completed.json telemetry_salt.txt /NJH /NJS /NDL /NP /R:2 /W:1

if errorlevel 8 (
    echo.
    echo [ERROR] File copy failed with errors.
    pause
    exit /b 1
)

echo.
echo [SUCCESS] Deployment complete!
echo   Destination: %DEST_DIR%
echo   Media ID:    %VOL_SERIAL%
echo.
echo You can now run Fabriq from:
echo   %DEST_DIR%
echo.

set /p "RUN_FABRIQ=続けて Fabriq を実行しますか？ (Y/N): "
if /i "%RUN_FABRIQ%"=="Y" (
    cd /d "%DEST_DIR%"
    if exist "%DEST_DIR%\Fabriq.exe" (
        echo [INFO] Fabriq.exe を起動します...
        start "" "%DEST_DIR%\Fabriq.exe"
    ) else (
        echo [ERROR] Fabriq.exe が見つかりません: %DEST_DIR%\Fabriq.exe
    )
    exit /b 0
)

echo.
pause
exit /b 0
