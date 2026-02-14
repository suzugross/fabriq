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
:: 設定: プリンター管理スクリプトの場所
:: ---------------------------------------------------------
set "PRN_SCRIPT=C:\Windows\System32\Printing_Admin_Scripts\ja-JP\prndrvr.vbs"

cd /d %~dp0

if not exist "INF" (
    echo [エラー] "INF" フォルダが見つかりません。
    pause
    exit /b
)

:: =========================================================
:: 【STEP 1】 プリンターモデル（-m）の選択
:: =========================================================
:MENU_MODEL
cls
echo ========================================================
echo   [Step 1/3] プリンターモデル（フォルダ）を選択
echo   ※これが -m (Model Name) になります
echo ========================================================
echo.

set modelCount=0
for /D %%D in ("INF\*") do (
    set /a modelCount+=1
    set "ModelName[!modelCount!]=%%~nxD"
    set "ModelRootPath[!modelCount!]=%%~fD"
    echo   [!modelCount!] %%~nxD
)

if !modelCount! equ 0 (
    echo [エラー] INFフォルダ内にモデルフォルダがありません。
    pause
    exit /b
)

echo.
:INPUT_MODEL
set /p "ModelChoice=番号を入力 (1-!modelCount!): "
if "%ModelChoice%"=="" goto INPUT_MODEL
if %ModelChoice% LSS 1 goto INPUT_MODEL
if %ModelChoice% GTR !modelCount! goto INPUT_MODEL

set "TargetModelName=!ModelName[%ModelChoice%]!"
set "TargetModelRoot=!ModelRootPath[%ModelChoice%]!"

:: =========================================================
:: 【STEP 2】 INFファイル（-i）の選択
:: =========================================================
:MENU_INF
cls
echo ========================================================
echo   [Step 2/3] インストール定義ファイル (.inf) を選択
echo   ※これが -i (Inf File) になります
echo ========================================================
echo.
echo   検索中...

set infCount=0
:: INFファイルを検索してリストアップ
for /f "delims=" %%F in ('dir /s /b "%TargetModelRoot%\*.inf" 2^>nul') do (
    set /a infCount+=1
    set "InfFullPath[!infCount!]=%%F"
    set "InfFileName[!infCount!]=%%~nxF"
    echo   [!infCount!] %%~nxF
    echo        (パス: %%~dpF)
    echo.
)

if !infCount! equ 0 (
    echo [エラー] .inf ファイルが見つかりませんでした。
    pause
    goto MENU_MODEL
)

:INPUT_INF
set /p "InfChoice=番号を入力 (1-!infCount!): "
if "%InfChoice%"=="" goto INPUT_INF
if %InfChoice% LSS 1 goto INPUT_INF
if %InfChoice% GTR !infCount! goto INPUT_INF

set "FinalInfPath=!InfFullPath[%InfChoice%]!"

:: =========================================================
:: 【STEP 3】 DLL格納フォルダ（-h）の選択
:: =========================================================
:MENU_DLL
cls
echo ========================================================
echo   [Step 3/3] ドライバー本体 (.dll) があるフォルダを選択
echo   ※これが -h (Path) になります
echo ========================================================
echo.
echo   DLLを検索してフォルダ一覧を作成中...

set dllFolderCount=0
set "PrevDir="

:: DLLファイルを検索し、ディレクトリ単位で重複を排除してリスト化
:: sort機能を使って同じフォルダのファイルを連続させ、重複判定を行う
for /f "delims=" %%A in ('dir /s /b /o:n "%TargetModelRoot%\*.dll" 2^>nul') do (
    set "CurrDir=%%~dpA"
    
    :: 直前のループと同じフォルダならスキップ（重複排除）
    if "!CurrDir!" neq "!PrevDir!" (
        set /a dllFolderCount+=1
        set "DllFolderPath[!dllFolderCount!]=!CurrDir!"
        
        echo   [!dllFolderCount!] !CurrDir!
        set "PrevDir=!CurrDir!"
    )
)

if !dllFolderCount! equ 0 (
    echo [警告] .dll ファイルが見つかりませんでした。
    echo INFファイルのあるフォルダをそのまま使用しますか？
    echo.
    echo   [1] はい (INFと同じ場所を使用)
    echo   [0] いいえ (中断)
    set /p "DllWarn=選択: "
    if "!DllWarn!"=="1" (
        :: INFの親フォルダを取得してセット
        for %%I in ("!FinalInfPath!") do set "FinalDllDir=%%~dpI"
        goto EXECUTE
    )
    goto MENU_MODEL
)

echo.
:INPUT_DLL
set /p "DllChoice=番号を入力 (1-!dllFolderCount!): "
if "%DllChoice%"=="" goto INPUT_DLL
if %DllChoice% LSS 1 goto INPUT_DLL
if %DllChoice% GTR !dllFolderCount! goto INPUT_DLL

set "FinalDllDir=!DllFolderPath[%DllChoice%]!"
:: 末尾の\を削除
if "!FinalDllDir:~-1!"=="\" set "FinalDllDir=!FinalDllDir:~0,-1!"

:: =========================================================
:: 【実行】 インストール
:: =========================================================
:EXECUTE
cls
echo --------------------------------------------------------
echo   インストールを開始します
echo --------------------------------------------------------
echo   モデル (-m): %TargetModelName%
echo   ＩＮＦ (-i): %FinalInfPath%
echo   パス   (-h): %FinalDllDir%
echo --------------------------------------------------------
echo.

cscript "%PRN_SCRIPT%" -a -m "%TargetModelName%" -v 3 -e "windows x64" -h "%FinalDllDir%" -i "%FinalInfPath%"

if %errorlevel% neq 0 (
    echo.
    echo [失敗] エラーが発生しました。
    echo 管理者権限や、選択したパス構成が正しいか確認してください。
) else (
    echo.
    echo [成功] ドライバーが追加されました。
)

echo.
pause