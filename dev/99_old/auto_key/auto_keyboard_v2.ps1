<#
    汎用キーボード自動化ツール (Event-Driven Edition)
    使い方: 同階層にある 'recipe.csv' を読み込み、順次実行します。
#>

# --- 設定エリア ---
$csvPath = ".\recipe.csv"
# CSVの文字コード (Excelで保存した場合は Default(Shift-JIS)、VSCodeなら UTF8)
$csvEncoding = "Default" 
# ------------------

# .NETのアセンブリをロード (SendKeys利用のため)
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.VisualBasic

# WScript.Shell (ウィンドウアクティブ化用)
$wsShell = New-Object -ComObject WScript.Shell

# CSVファイルの存在チェック
if (-not (Test-Path $csvPath)) {
    Write-Error "設定ファイル '$csvPath' が見つかりません。"
    exit
}

# CSV読み込み
try {
    $commands = Import-Csv $csvPath -Encoding $csvEncoding
}
catch {
    Write-Error "CSVの読み込みに失敗しました。文字コードやフォーマットを確認してください。"
    exit
}

Write-Host "=== 自動化を開始します ===" -ForegroundColor Cyan

foreach ($row in $commands) {
    # 空行スキップ
    if ([string]::IsNullOrWhiteSpace($row.Action)) { continue }

    Write-Host "[Step $($row.Step)] $($row.Note) ($($row.Action))" -NoNewline

    switch ($row.Action) {
        "Open" {
            # アプリ起動 (引数あり/なし自動判定)
            $val = $row.Value
            if ($val -match " ") {
                $parts = $val -split " ", 2
                $prog = $parts[0]
                $argsList = $parts[1]
                try {
                    Start-Process -FilePath $prog -ArgumentList $argsList -ErrorAction Stop
                } catch {
                    # 失敗時はシェル経由で起動 (URLなどを確実に開くため)
                    Start-Process -FilePath "cmd" -ArgumentList "/c start $prog $argsList" -WindowStyle Hidden
                }
            } else {
                Start-Process $val
            }
            Write-Host " -> 起動完了" -ForegroundColor Green
        }

        "WaitWin" {
            # ウィンドウが現れるまで待機 (イベント駆動)
            $targetTitle = $row.Value
            $timeout = [int]$row.Wait
            $elapsed = 0
            $found = $false
            
            Write-Host "`n   Waiting: '$targetTitle' を探しています (最大 $($timeout/1000)秒)..." -NoNewline

            while ($elapsed -lt $timeout) {
                # タイトル部分一致検索
                $proc = Get-Process | Where-Object { $_.MainWindowTitle -like "*$targetTitle*" } | Select-Object -First 1
                
                if ($proc) {
                    $found = $true
                    Write-Host " OK!" -ForegroundColor Green
                    
                    # 見つけたウィンドウをアクティブ化
                    try {
                        $wsShell.AppActivate($proc.Id)
                        Start-Sleep -Milliseconds 500 # アクティブ化の安定待ち
                    } catch {}
                    break
                }
                Start-Sleep -Milliseconds 500
                $elapsed += 500
                Write-Host "." -NoNewline
            }

            if (-not $found) {
                Write-Warning "`n   [Timeout] 指定時間内にウィンドウが見つかりませんでした。"
                # タイムアウト時に停止したい場合は下の行の # を外す
                # break 
            }
        }

        "AppFocus" {
            # 既存のウィンドウに切り替え
            $success = $wsShell.AppActivate($row.Value)
            if ($success) { Write-Host " -> Focus OK" -ForegroundColor Green }
            else { Write-Warning " -> Focus Failed (Window not found)" }
        }

        "Type" {
            # 文字列入力
            [System.Windows.Forms.SendKeys]::SendWait($row.Value)
            Write-Host " -> 入力完了" -ForegroundColor Green
        }

        "Key" {
            # 特殊キー送信
            [System.Windows.Forms.SendKeys]::SendWait($row.Value)
            Write-Host " -> Key送信" -ForegroundColor Green
        }

        "Wait" {
            # 単純待機
            $ms = [int]$row.Value
            Start-Sleep -Milliseconds $ms
            Write-Host " -> $($ms)ms 待機" -ForegroundColor Gray
        }

        Default {
            Write-Warning " -> 未定義のアクション: $($row.Action)"
        }
    }

    # 各ステップ後の固定ウェイト (Waitカラムが0より大きい場合のみ)
    if ([int]$row.Wait -gt 0 -and $row.Action -ne "WaitWin") {
        Start-Sleep -Milliseconds ([int]$row.Wait)
    }
    
    Write-Host "" # 改行
}

Write-Host "=== 全ての処理が完了しました ===" -ForegroundColor Cyan
# 実行完了を確認できるように一時停止 (不要なら削除可)
Read-Host "Enterキーを押して終了してください"