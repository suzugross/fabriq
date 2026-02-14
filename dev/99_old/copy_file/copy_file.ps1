<#
.SYNOPSIS
初期キッティング用ファイル配布スクリプト（上書き制御機能付き）
#>

# ログを見やすくするための設定
$Host.UI.RawUI.WindowTitle = "キッティング用ファイル配布ツール"
[Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding("utf-8")

# カレントディレクトリをスクリプトの場所に固定
Set-Location $PSScriptRoot

$csvFile = "source.csv"
$sourceDir = "source"

# --------------------------------------------------
# 事前チェック
# --------------------------------------------------
if (!(Test-Path $csvFile)) {
    Write-Error "[Critical] 設定ファイル($csvFile)が見つかりません。"
    Read-Host "Enterキーを押して終了してください"
    exit
}

if (!(Test-Path $sourceDir)) {
    Write-Error "[Critical] 配布元フォルダ($sourceDir)が見つかりません。"
    Read-Host "Enterキーを押して終了してください"
    exit
}

# CSV読み込み
try {
    $taskList = Import-Csv $csvFile -Encoding Default
}
catch {
    Write-Error "[Critical] CSVの読み込みに失敗しました。フォーマットを確認してください。"
    Read-Host "Enterキーを押して終了してください"
    exit
}

Write-Host "処理を開始します... 対象件数: $($taskList.Count)" -ForegroundColor Cyan
Write-Host "--------------------------------------------------"

# --------------------------------------------------
# メイン処理
# --------------------------------------------------
foreach ($row in $taskList) {
    $id       = $row.adminID
    $fileName = $row.filename
    $destDir  = $row.path
    $enable   = $row.enable
    
    # overwriteカラムの値を取得（1またはTRUEなら $true、それ以外は $false）
    $isOverwrite = ($row.overwrite -eq "1" -or $row.overwrite -eq "TRUE")

    # 1. 有効化チェック
    if ($enable -ne "1") {
        Write-Host "[$id] SKIP : 無効化されています ($fileName)" -ForegroundColor Gray
        continue
    }

    # パス生成
    $srcPath = Join-Path $sourceDir $fileName
    $destPath = Join-Path $destDir $fileName  # 最終的な配置パス

    # 2. ソース確認
    if (!(Test-Path $srcPath)) {
        Write-Host "[$id] ERROR: 配布元ファイルなし ($fileName)" -ForegroundColor Red
        continue
    }

    # 3. 親フォルダ作成
    if (!(Test-Path $destDir)) {
        try {
            New-Item -ItemType Directory -Path $destDir -Force | Out-Null
        }
        catch {
            Write-Host "[$id] ERROR: フォルダ作成失敗 ($destDir)" -ForegroundColor Red
            continue
        }
    }

    # 4. コピー実行判定（ここが重要）
    if (Test-Path $destPath) {
        # 既にファイルが存在する場合
        if ($isOverwrite) {
            # 上書き許可なら実行
            try {
                Copy-Item -Path $srcPath -Destination $destDir -Recurse -Force -ErrorAction Stop
                Write-Host "[$id] UPDATE: 上書きしました ($fileName)" -ForegroundColor Yellow
            }
            catch {
                Write-Host "[$id] ERROR : 上書き失敗 ($fileName) $_" -ForegroundColor Red
            }
        }
        else {
            # 上書き禁止ならスキップ
            Write-Host "[$id] SKIP : 既存ファイルがあるためスキップ ($fileName)" -ForegroundColor DarkCyan
        }
    }
    else {
        # ファイルが存在しない場合（新規コピー）
        try {
            Copy-Item -Path $srcPath -Destination $destDir -Recurse -Force -ErrorAction Stop
            Write-Host "[$id] OK   : コピー成功 ($fileName)" -ForegroundColor Green
        }
        catch {
            Write-Host "[$id] ERROR : コピー失敗 ($fileName) $_" -ForegroundColor Red
        }
    }
}

Write-Host "--------------------------------------------------"
Write-Host "すべての処理が完了しました。"
# 実行結果を確認できるよう停止
Read-Host "Enterキーを押して終了してください"