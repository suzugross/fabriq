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

:: ---------------------------------------------------------
:: 設定
:: ---------------------------------------------------------
set "PRN_DIR=C:\Windows\System32\Printing_Admin_Scripts\ja-JP"
set "CSV_FILE=kitting_list.csv"

cd /d %~dp0

if not exist "%CSV_FILE%" (
    echo [エラー] "%CSV_FILE%" が見つかりません。
    pause
    exit /b
)

:: ---------------------------------------------------------
:: メニュー表示：CSVからPC一覧を読み込み
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

:: CSVを読み込み、PC名(1列目)と、残りのデータ(2列目以降)を配列に保存
:: tokens=1* : 1列目を %%a に、2列目以降すべてを %%b に入れる
for /f "usebackq tokens=1* delims=," %%a in ("%CSV_FILE%") do (
    set /a count+=1
    set "PcName[!count!]=%%a"
    set "PrinterData[!count!]=%%b"
    
    :: 現在のPC名と一致したら「★」をつける演出
    if /i "%%a"=="%COMPUTERNAME%" (
        echo   [!count!]  %%a  [★ 現在のPC]
    ) else (
        echo   [!count!]  %%a
    )
)

echo.
if !count! equ 0 (
    echo [エラー] CSVファイルが空か、正しく読み込めません。
    pause
    exit /b
)

:: ---------------------------------------------------------
:: PC選択
:: ---------------------------------------------------------
:INPUT
set /p "Choice=設定を行うPCの番号を入力してください (1-!count!): "

if "%Choice%"=="" goto INPUT
if %Choice% LSS 1 goto INPUT
if %Choice% GTR !count! goto INPUT

set "TargetPC=!PcName[%Choice%]!"
set "RawData=!PrinterData[%Choice%]!"

cls
echo ========================================================
echo   対象PC: %TargetPC%
echo   プリンタのインストールを開始します...
echo ========================================================
echo.

:: ---------------------------------------------------------
:: データ解析とインストールループ
:: ---------------------------------------------------------
:: RawDataには "P1名,P1モデル,P1ポート,P1IP,P2名..." が入っている
:: これを先頭から4つずつ切り出して処理する

:PROCESS_LOOP
:: データが空になったら終了
if "!RawData!"=="" goto FINISH

:: 文字列解析
:: tokens=1,2,3,4* : 先頭4つを変数へ、残りを %%e へ
for /f "tokens=1,2,3,4* delims=," %%a in ("!RawData!") do (
    set "P_Name=%%a"
    set "P_Model=%%b"
    set "P_Port=%%c"
    set "P_IP=%%d"
    
    :: 残りのデータを更新（次のループ用）
    set "RawData=%%e"
)

:: 必須項目が欠けていないか簡易チェック
if "%P_IP%"=="" (
    echo [終了] 全データの処理が完了、またはデータ形式が不正です。
    goto FINISH
)

echo --------------------------------------------------------
echo [処理中] %P_Name%
echo --------------------------------------------------------
echo   モデル: %P_Model%
echo   ポート: %P_Port% (IP: %P_IP%)

:: 1. ポート作成
echo   - ポートを作成しています...
cscript "%PRN_DIR%\prnport.vbs" -a -r "%P_Port%" -h "%P_IP%" -o raw -n 9100 >nul
if %errorlevel% equ 0 ( echo     [OK] ポート作成 ) else ( echo     [Info] ポート作成スキップ/エラー )

:: 2. プリンタ作成
echo   - プリンタキューを作成しています...
cscript "%PRN_DIR%\prnmngr.vbs" -a -p "%P_Name%" -m "%P_Model%" -r "%P_Port%" >nul
if %errorlevel% equ 0 (
    echo     [OK] プリンタ追加完了
) else (
    echo     [NG] プリンタ追加失敗
    echo          ※ドライバ "%P_Model%" がインストールされているか確認してください
)

echo.
:: ループ先頭に戻る
goto PROCESS_LOOP


:FINISH
echo.
echo ========================================================
echo   全ての処理が完了しました。
echo ========================================================
pause