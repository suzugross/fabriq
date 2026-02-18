@echo off
setlocal enabledelayedexpansion

REM ===== 管理者権限チェック・昇格処理 =====
echo [*] 管理者権限を確認中...

REM 管理者権限チェック（net sessionコマンド使用）
net session >nul 2>&1
if errorlevel 1 (
    echo [!] 管理者権限が必要です。昇格して再実行します...
    
    REM PowerShellでUAC昇格して自分自身を再実行
    powershell -Command "Start-Process cmd -ArgumentList '/c \"%~f0\"' -Verb RunAs"
    exit /b
)

echo [+] 管理者権限で実行中です。

REM ===== ここから本来の処理 =====

cd /d %~dp0


setup.exe /download CDN.xml


pause


exit /b
