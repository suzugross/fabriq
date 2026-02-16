# フィルタリングプリンタ表示スクリプト

# CSVファイルから削除対象プリンタを読み込む
$csvPath = "list.csv"

if (-not (Test-Path $csvPath)) {
    Write-Host "エラー: list.csv が見つかりません。" -ForegroundColor Red
    Write-Host "このスクリプトと同じフォルダに list.csv を配置してください。" -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Enterキーを押して終了してください"
    exit
}

# フィルタリスト（表示対象のプリンタ）
$filterList = @(
    '【給水】LBP8710',
    '【経理】LBP8710',
    '【管路保全】LBP8710',
    '【管路保全】Apeos C3570（FAX）',
    '【管路保全】Apeos C3570',
    '【水走】LBP8710',
    '【水走】Apeos3570（FAX）',
    '【水走】Apeos3570',
    '【分室】Apeos3570（FAX）',
    '【分室】Apeos3570',
    'LPC240005',
    'LPC240004',
    'LPC240003',
    'LPC240002',
    'LPC240001',
    'LBP842C',
    'DocuCentre-V 5080N（FAX）',
    'DocuCentre-V 5080N',
    'Apeos7580',
    'Apeos5580',
    'Apeos4570（FAX）',
    'Apeos4570',
    'Apeos1860',
    'Apeos C2571'
)

# 現在のPC名を取得
$currentPCName = $env:COMPUTERNAME

# CSVから現在のPC名に対応する削除対象プリンタ名を取得
$printerData = Get-Content $csvPath -Encoding Default | ConvertFrom-Csv -Header "PCName","PrinterName"
$currentPCDeleteTargets = $printerData | Where-Object { $_.PCName -eq $currentPCName } | Select-Object -ExpandProperty PrinterName

Write-Host "================================================" -ForegroundColor Cyan
Write-Host "  インストールされているプリンタ一覧" -ForegroundColor Cyan
Write-Host "  (フィルタリング表示)" -ForegroundColor Cyan
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "PC名: " -NoNewline
Write-Host "$currentPCName" -ForegroundColor Green
Write-Host ""

# インストールされているプリンタを取得
Write-Host "プリンタ情報を取得中..." -ForegroundColor Yellow
$installedPrinters = Get-Printer | Select-Object -ExpandProperty Name

# フィルタリストに含まれるプリンタのみ抽出
$filteredPrinters = $installedPrinters | Where-Object { $filterList -contains $_ }

Write-Host ""
Write-Host "================================================" -ForegroundColor Cyan
Write-Host "凡例: " -NoNewline
Write-Host "■ 削除対象 " -ForegroundColor Red -NoNewline
Write-Host "/ " -NoNewline
Write-Host "■ 削除対象外" -ForegroundColor Blue
Write-Host "================================================" -ForegroundColor Cyan
Write-Host ""

if ($filteredPrinters.Count -eq 0) {
    Write-Host "フィルタ対象のプリンタは見つかりませんでした。" -ForegroundColor Green
    Write-Host ""
    Write-Host "このPCにはフィルタリストに該当するプリンタがインストールされていません。" -ForegroundColor Green
} else {
    $deleteTargetCount = 0
    $nonTargetCount = 0
    
    foreach ($printer in $filteredPrinters) {
        # 現在のPCの削除対象に含まれているか判定
        if ($currentPCDeleteTargets -contains $printer) {
            Write-Host "  ● " -NoNewline -ForegroundColor Red
            Write-Host "$printer" -ForegroundColor Red
            $deleteTargetCount++
        } else {
            Write-Host "  ● " -NoNewline -ForegroundColor Blue
            Write-Host "$printer" -ForegroundColor Blue
            $nonTargetCount++
        }
    }
    
    Write-Host ""
    Write-Host "================================================" -ForegroundColor Cyan
    Write-Host "フィルタ該当: $($filteredPrinters.Count) 個のプリンタ" -ForegroundColor White
    Write-Host "  削除対象: " -NoNewline
    Write-Host "$deleteTargetCount 個" -ForegroundColor Red
    Write-Host "  削除対象外: " -NoNewline
    Write-Host "$nonTargetCount 個" -ForegroundColor Blue
    Write-Host "================================================" -ForegroundColor Cyan
    
    if ($deleteTargetCount -gt 0) {
        Write-Host ""
        Write-Host "※削除対象プリンタが見つかりました。" -ForegroundColor Yellow
        Write-Host "  remove_printers.ps1 を実行して削除できます。" -ForegroundColor Yellow
    }
}

Write-Host ""
Read-Host "Enterキーを押して終了してください"
