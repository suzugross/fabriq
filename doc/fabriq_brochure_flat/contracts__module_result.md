# ModuleResult 契約

`KERNEL_API.md §5` で公式宣言。fabriq モジュールがカーネルへ実行結果を返すための統一スキーマ。**すべてのモジュールはこの契約に従って結果を返却する**ことが義務付けられている。

---

## 戻り値オブジェクトのフィールド

| フィールド | 型 | 必須 | 意味 |
|---|---|---|---|
| `_IsModuleResult` | `bool` | 必須（常に `$true`） | カーネル側のフィルタが pipeline から ModuleResult を取り出す際の識別マーカー |
| `Status` | `string` | 必須 | `Success` / `Error` / `Cancelled` / `Skipped` / `Partial` のいずれか（ValidateSet） |
| `Message` | `string` | 必須（空文字 OK） | 結果の人間可読なメッセージ。実行履歴 CSV / HTML チェックリストに表示される |
| `Details` | `array` | 任意 | 任意の詳細情報。実行履歴には入らないが pipeline 経由でテストハーネスが消費可能 |
| `Verified` | `Nullable[bool]` | 任意 | Post-Apply Verification 結果。`$null`=未検証 / `$true`=PASS / `$false`=FAIL |
| `Timestamp` | `DateTime` | 必須（自動付与） | `Get-Date` で生成時刻を打鍵 |

---

## 標準呼び出しパターン

### 単純な完了

```powershell
return (New-ModuleResult -Status "Success" -Message "Done")
```

### 検証付き完了

```powershell
return (New-ModuleResult -Status "Success" -Message "Hostname renamed to NEW-PC-01 (pending reboot)" -Verified $true)
```

### バッチ集計

```powershell
return (New-BatchResult -Success 3 -Skip 1 -Fail 0 `
                        -Title "Registry Configuration" `
                        -MessageSuffix "(verified by readback)" `
                        -Verified $true)
```

`New-BatchResult` は内部で：
1. `Show-Separator` + `$Title` + 集計を画面表示
2. `$Status` を自動判定（Fail=0 + Success>0 → `Success` / Success>0 + Fail>0 → `Partial` / Skip 全件 → `Skipped` / Fail のみ → `Error`）
3. Message を `"Success: 3, Skip: 1, Fail: 0 (verified by readback)"` に組み立て
4. `New-ModuleResult` を呼んで返す

### キャンセル（Y/N で N 押下）

```powershell
$cancelled = Confirm-ModuleExecution -Message "本当に実行しますか？"
if ($cancelled) { return $cancelled }   # 即 return（既に Cancelled の ModuleResult が入っている）
```

### スキップ（対象なし）

```powershell
$rows = Import-ModuleCsv -Path $listCsv -FilterEnabled -RequiredColumns @("Enabled","Path")
if ($null -eq $rows -or $rows.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled entries")
}
```

### エラー（例外発生）

```powershell
try {
    # ... 処理 ...
    return (New-ModuleResult -Status "Success" -Message "Applied $($results.Count) settings" -Verified $true)
}
catch {
    Show-Error "Configuration failed: $_"
    return (New-ModuleResult -Status "Error" -Message "Exception: $($_.Exception.Message)")
}
```

---

## Status セマンティクス

| Status | 意味 | HTML 表示色 | beep 通知 |
|---|---|---|---|
| `Success` | 全件正常完了 | 緑 | なし |
| `Partial` | 一部成功・一部失敗（New-BatchResult が Success>0 + Fail>0 で自動付与） | 黄 | 2-tone beep |
| `Skipped` | 全件スキップ（対象が無い、または明示スキップ） | 灰 | なし |
| `Cancelled` | ユーザーが Y/N で N 押下 | 黄 | なし |
| `Error` | エラー発生（例外 or 失敗） | 赤 | 3-tone beep |

`Invoke-ErrorNotification` は `Error` / `Partial` のときコンソールを foreground に持ってきて beep する（operator が見落とさないように）。

---

## Verified セマンティクス（Post-Apply Verification）

| Verified | 意味 | 推奨ケース |
|---|---|---|
| `$null` | 未検証 / 検証不可 | 検証ロジックを実装していないモジュール、または再起動後にしか確認できない（sysprep, hostname pending reboot 等） |
| `$true` | PASS（適用後の OS 状態が期待値と完全一致） | 設定読み返しが冪等性チェックと同じパターンで成立する場合 |
| `$false` | FAIL（適用は受理されたが乖離あり） | 適用 API が成功を返したのに読み返すと値が違う場合（外部要因や OS バグの兆候） |

### 実装パターン

```powershell
# Step 5.5 (Post-Apply Verification): 適用後に状態を読み返して照合
$verifiedCount = 0
$failedCount = 0
foreach ($entry in $rows) {
    if (Test-RegistryValueMatch -Path $entry.Path -Name $entry.Name -Expected $entry.Value) {
        $verifiedCount++
    } else {
        $failedCount++
    }
}
$allVerified = ($failedCount -eq 0)
return (New-BatchResult -Success $appliedCount -Skip $skipCount -Fail 0 `
                        -Title "Registry HKLM Results" -Verified $allVerified)
```

### 検証除外モジュール（feedback memory `project_verification_exclusions`）

以下は「**誤 PASS リスク回避**」の観点で意図的に検証を実装しない：

- `acl_config`: ACL ツリーを完全に読み返すと膨大、サブセット検証だと false PASS の可能性
- `spi_config`: Default Profile への hive load 経由のためログイン後にしか反映されず、検証時点では確定できない
- `copyfile_config`: ファイルが存在することと内容が正しいことは別、ハッシュ検証は重い

これらは `-Verified` 引数を省略（`$null`）。Guide.txt にも検証非実装の理由が明示される。

---

## Pipeline Capture と Fallback

カーネルは `Invoke-SafeCommand` 内で 2 段階で ModuleResult を捕捉：

```powershell
$global:_LastModuleResult = $null    # クリア
$output = & $ScriptBlock              # モジュール実行
foreach ($item in @($output)) {
    if ($item -is [PSCustomObject] -and $item._IsModuleResult -eq $true) {
        $moduleResult = $item
    }
}
if (-not $moduleResult -and $null -ne $global:_LastModuleResult) {
    $moduleResult = $global:_LastModuleResult   # フォールバック
}
```

**Pipeline capture 失敗例**: モジュールが内部で `Write-Output` を雑に呼び出して pipeline が混線、PowerShell が暗黙の implicit return で複数オブジェクトを return した場合に発生。`New-ModuleResult` 自体が内部で `$global:_LastModuleResult = $resultObj` を行うため、フォールバックで救済される（防御的二重化）。

---

## モジュールスクリプトの典型構造（7 step）

`dev/template/_template_script.ps1` がベースとなる正典スケルトン。すべてのモジュールはこの構造を踏襲する：

```powershell
# Step 1: Initialization
. "$PSScriptRoot\..\..\..\kernel\common.ps1"
$ModuleName = "Module Name"

# Step 2: Header (Show-CategorySeparator + Show-Info)
Show-CategorySeparator $ModuleName
Show-Info "..."

# Step 3: Confirmation
$cancelled = Confirm-ModuleExecution
if ($cancelled) { return $cancelled }

# Step 4: Load CSV
$csvPath = Join-Path $PSScriptRoot "<name>_list.csv"
$rows = Import-ModuleCsv -Path $csvPath -FilterEnabled -RequiredColumns @(...)
if ($null -eq $rows) { return (New-ModuleResult -Status "Error" -Message "CSV load failed") }
if ($rows.Count -eq 0) { return (New-ModuleResult -Status "Skipped" -Message "No enabled rows") }

# Step 5: Apply (loop)
$successCount = 0; $skipCount = 0; $failCount = 0
foreach ($row in $rows) {
    # idempotency check (skip if already applied)
    # apply
    # increment counter
}

# Step 5.5: Post-Apply Verification (optional but recommended)
$verifiedAll = $true
foreach ($row in $rows) { if (-not (verify $row)) { $verifiedAll = $false } }

# Step 6: Display summary (handled by New-BatchResult)

# Step 7: Return result
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount -Verified $verifiedAll)
```

カーネルが Step 6 / Step 7 の表示を内部で完結させる設計のため、モジュール側は集計値だけ用意すればよい。
