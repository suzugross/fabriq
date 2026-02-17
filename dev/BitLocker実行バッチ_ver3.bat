@echo off
setlocal enabledelayedexpansion

REM ========================================================
REM [修正版 v2] BitLocker 統合有効化スクリプト
REM （PC名フォルダ作成・格納版）
REM 
REM [Cドライブ]
REM  - 全体暗号化 / 新モード(XTS-AES 128) / ハードウェアテストスキップ
REM 
REM [Dドライブ]
REM  - 全体暗号化 / 新モード(XTS-AES 128)
REM  - 自動ロック解除(AutoUnlock)
REM 
REM [保存先]
REM  - 実行場所\%COMPUTERNAME%\%COMPUTERNAME%-C.txt
REM  - 実行場所\%COMPUTERNAME%\%COMPUTERNAME%-D.txt
REM ========================================================

REM ===== 管理者権限チェック・昇格処理 =====
echo [*] 管理者権限を確認中...

net session >nul 2>&1
if errorlevel 1 (
    echo [!] 管理者権限が必要です。昇格して再実行します...
    powershell -Command "Start-Process cmd -ArgumentList '/c \"%~f0\"' -Verb RunAs"
    exit /b
)

echo [+] 管理者権限で実行中です。

REM ===== 保存先フォルダの準備 =====
echo.
echo [準備] 保存先フォルダを確認しています...

REM バッチファイルの場所に「PC名」のフォルダを作成
set "KEY_DIR=%~dp0%COMPUTERNAME%"

if not exist "%KEY_DIR%" (
    mkdir "%KEY_DIR%"
    echo [作成] フォルダを作成しました: %KEY_DIR%
) else (
    echo [確認] フォルダは既に存在します: %KEY_DIR%
)

REM ===== ここから本来の処理 =====

echo ========================================================
echo BitLockerの設定を開始します...
echo ========================================================
echo.

REM --- 2. Cドライブの設定 ---
REM パスをフォルダ内に指定
set "LOG_C=%KEY_DIR%\%COMPUTERNAME%-C.txt"

echo [処理中] Cドライブの暗号化設定を行っています...
echo 出力先: %LOG_C%

REM Cドライブ有効化
manage-bde -on C: -RecoveryPassword -EncryptionMethod xts_aes128 -SkipHardwareTest > "%LOG_C%" 2>&1

if %errorlevel% equ 0 (
    echo [OK] Cドライブの設定コマンドが成功しました。
) else (
    echo [ERROR] Cドライブの設定に失敗しました。ログを確認してください。
)

echo.
echo --------------------------------------------------------
echo.

REM --- 3. Dドライブの設定 ---
REM パスをフォルダ内に指定
set "LOG_D=%KEY_DIR%\%COMPUTERNAME%-D.txt"

if exist D:\ (
    echo [処理中] Dドライブの暗号化設定を行っています...
    echo 出力先: %LOG_D%

    REM 3-1. Dドライブ有効化 (回復パスワード発行)
    manage-bde -on D: -RecoveryPassword -EncryptionMethod xts_aes128 > "%LOG_D%" 2>&1
    
    if %errorlevel% equ 0 (
        echo [OK] Dドライブの暗号化を開始しました。
        
        echo.
        echo [処理中] Dドライブの自動ロック解除を設定しています...
        
        REM 3-2. 自動ロック解除の有効化 (ログは追記 >> を使用)
        manage-bde -autounlock -enable D: >> "%LOG_D%" 2>&1
        
        if %errorlevel% equ 0 (
             echo [OK] 自動ロック解除を有効にしました。
        ) else (
             echo [WARNING] 自動ロック解除の設定に失敗しました。
             echo Cドライブの暗号化が完了してから手動で設定が必要な場合があります。
        )
        
    ) else (
        echo [ERROR] Dドライブの暗号化設定に失敗しました。ログを確認してください。
    )

) else (
    echo [SKIP] Dドライブが見つからないため、処理をスキップしました。
)

echo.
echo ========================================================
echo 全ての処理が完了しました。
echo 以下のフォルダに回復キー(テキストファイル)が保存されています。
echo 保存先: %KEY_DIR%
echo.
echo 設定を反映するため、PCを【再起動】してください。
echo ========================================================

pause

cls

REM ===== 生成されたテキストファイルの内容を表示 =====
echo.
echo ========================================================
echo 生成されたファイルの内容を表示します
echo ========================================================
echo.

REM Cドライブのログファイルを表示
if exist "%LOG_C%" (
    echo ┌────────────────────────────────────────────────────────
    echo │ ファイル: %LOG_C%
    echo └────────────────────────────────────────────────────────
    type "%LOG_C%"
    echo.
) else (
    echo [注意] Cドライブのログファイルが見つかりません。
    echo.
)

pause
cls

REM Dドライブのログファイルを表示
if exist "%LOG_D%" (
    echo ┌────────────────────────────────────────────────────────
    echo │ ファイル: %LOG_D%
    echo └────────────────────────────────────────────────────────
    type "%LOG_D%"
    echo.
) else (
    echo [注意] Dドライブのログファイルが見つかりません。
    echo.
)

echo ========================================================
pause