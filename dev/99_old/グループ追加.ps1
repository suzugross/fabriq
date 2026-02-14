# ドメイン設定（固定）
$domain = "city.higashiosaka.osaka.jp"

# 設定リスト（ローカルグループ = 追加したいドメイングループ）
$groups = @{
    "Administrators" = "Client Administrators"
    "Power Users"    = "Client Power Users"
}

Write-Host "--- グループ追加処理を開始します ---"

# ループ処理
foreach ($localGroup in $groups.Keys) {
    $targetMember = "$domain\$($groups[$localGroup])"
    
    Write-Host "[$localGroup] に [$targetMember] を追加中..."
    
    # グループに追加（エラーが発生しても無視して次へ進む設定）
    Add-LocalGroupMember -Group $localGroup -Member $targetMember -ErrorAction SilentlyContinue
}

Write-Host "--- 処理完了 ---"
Read-Host "[Enter]キーを押して終了してください"