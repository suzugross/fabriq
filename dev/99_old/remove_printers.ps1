# プリンタ削除スクリプト（一括削除版）
# CSVファイルのエンコーディングを指定して読み込み
$csvPath = "list.csv"
$encoding = [System.Text.Encoding]::GetEncoding("Shift_JIS")
$printerData = Get-Content $csvPath -Encoding Default | ConvertFrom-Csv -Header "PrinterName"

# プリンタ名の一覧を取得
$printersToDelete = $printerData | Select-Object -ExpandProperty PrinterName

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  プリンタ削除スクリプト" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "削除対象のプリンタ一覧:" -ForegroundColor Yellow
Write-Host ""
$printersToDelete | ForEach-Object { Write-Host "  - $_" }
Write-Host ""
Write-Host "合計 $($printersToDelete.Count) 個のプリンタを削除します。" -ForegroundColor Yellow
Write-Host ""

# 確認
$confirm = Read-Host "削除を実行しますか? (Y/N)"

if ($confirm -ne "Y" -and $confirm -ne "y") {
    Write-Host "削除をキャンセルしました。" -ForegroundColor Yellow
    exit
}

Write-Host ""
Write-Host "プリンタの削除を開始します..." -ForegroundColor Cyan
Write-Host ""

# プリンタを削除
$successCount = 0
$failCount = 0
$notFoundCount = 0

foreach ($printerName in $printersToDelete) {
    try {
        # プリンタの存在確認
        $printer = Get-Printer -Name $printerName -ErrorAction SilentlyContinue
        
        if ($printer) {
            # プリンタを削除
            Remove-Printer -Name $printerName -ErrorAction Stop
            Write-Host "[成功] $printerName を削除しました。" -ForegroundColor Green
            $successCount++
        } else {
            Write-Host "[スキップ] $printerName は存在しません。" -ForegroundColor Yellow
            $notFoundCount++
        }
    } catch {
        Write-Host "[エラー] $printerName の削除に失敗しました: $($_.Exception.Message)" -ForegroundColor Red
        $failCount++
    }
}

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  削除処理が完了しました" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "成功: $successCount 個" -ForegroundColor Green
Write-Host "スキップ(未インストール): $notFoundCount 個" -ForegroundColor Yellow
Write-Host "失敗: $failCount 個" -ForegroundColor Red
Write-Host ""

# 終了前に待機
Read-Host "Enterキーを押して終了してください"