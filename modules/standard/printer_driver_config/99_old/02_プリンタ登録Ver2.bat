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

:: 設定
set "PRN_DIR=C:\Windows\System32\Printing_Admin_Scripts\ja-JP"
set "CSV_FILE=kitting_list.csv"

if not exist "%CSV_FILE%" (
    echo [エラー] "%CSV_FILE%" が見つかりません。
    pause
    exit /b
)

:: ---------------------------------------------------------
:: メニュー表示
:: ---------------------------------------------------------
:MENU
cls
echo ========================================================
echo   個別キッティング用 プリンタ一括インストーラー
echo ========================================================
echo   現在のPC名: %COMPUTERNAME%
echo.
echo   [No]  [対象PC名]
echo   --------------------------------------------------------

set count=0

:: CSV読み込み
for /f "usebackq tokens=1* delims=," %%a in ("%CSV_FILE%") do (
    set /a count+=1
    set "PcName[!count!]=%%a"
    set "PrinterData[!count!]=%%b"
    
    REM --- プリンタ名抽出 ---
    set "P1="
    set "P2="
    if not "%%b"=="" (
        for /f "tokens=1,5 delims=," %%A in ("%%b") do (
            set "P1=%%A"
            set "P2=%%B"
        )
    )
    
    REM --- 表示ステータス作成 ---
    set "StatusMark= "
    if /i "%%a"=="%COMPUTERNAME%" (
        set "StatusMark=★ THIS PC"
    )

    REM 表示用テキスト作成
    set "DispText=!P1!"
    if not "!P2!"=="" set "DispText=!P1!, !P2!..."
    
    REM 表示実行
    echo   [!count!]  %%a  [!StatusMark!] : [!DispText!]
)

echo.
if !count! equ 0 (
    echo [エラー] CSVファイルが読み込めませんでした。
    pause
    exit /b
)

:: ---------------------------------------------------------
:: PC選択
:: ---------------------------------------------------------
:INPUT
set /p "Choice=PC番号を入力 (1-!count!): "

if "%Choice%"=="" goto INPUT
if %Choice% LSS 1 goto INPUT
if %Choice% GTR !count! goto INPUT

set "TargetPC=!PcName[%Choice%]!"
set "RawData=!PrinterData[%Choice%]!"

cls
echo ========================================================
echo   対象PC: %TargetPC%
echo   インストールを開始します...
echo ========================================================

:: ---------------------------------------------------------
:: インストールループ
:: ---------------------------------------------------------
:PROCESS_LOOP
if "!RawData!"=="" goto FINISH

:: データを4つ取り出し、残りをRawDataに戻す
for /f "tokens=1,2,3,4* delims=," %%a in ("!RawData!") do (
    set "P_Name=%%a"
    set "P_Model=%%b"
    set "P_Port=%%c"
    set "P_IP=%%d"
    set "RawData=%%e"
)

:: 必須項目のチェック
if "%P_IP%"=="" goto FINISH

echo.
echo [設定中] %P_Name%
echo   - ドライバ: %P_Model%
echo   - IPポート: %P_IP%

:: ポート作成
:: エラー回避のため echo内のカッコを削除しました
cscript "%PRN_DIR%\prnport.vbs" -a -r "%P_Port%" -h "%P_IP%" -o raw -n 9100 >nul 2>&1
if %errorlevel% equ 0 (
    echo     [OK] ポート作成
) else (
    echo     [Info] ポート作成スキップ-エラー
)

:: プリンタ作成
cscript "%PRN_DIR%\prnmngr.vbs" -a -p "%P_Name%" -m "%P_Model%" -r "%P_Port%" >nul 2>&1
if %errorlevel% equ 0 (
    echo     [OK] プリンタ追加成功
) else (
    echo     [NG] プリンタ追加失敗
)

goto PROCESS_LOOP

:FINISH
echo.
echo --------------------------------------------------------
echo   完了しました。
pause