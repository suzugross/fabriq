# 非同期モジュール実行（__ASYNC__ + Runspace + Skip）

`__ASYNC__` マーカー以降のモジュールを別 PowerShell Runspace で実行し、Status Monitor からの **Skip ボタン** または **タイムアウト** で強制中断できる仕組み。kernel 2.1.0 で追加。

---

## なぜ async 実行が必要か

通常の `Invoke-SafeCommand` は同期実行で、モジュールが内部でハングすると fabriq 全体がブロックされる。長時間処理（domain join, winget install など）でこのリスクが顕在化したため、選択的に Runspace 実行に切り替える機構を導入した。

ただし「全モジュール async」は Runspace 起動コストとデバッグ性の問題があるため、**プロファイルで明示的に opt-in** する設計。

---

## kernel/json/async_config.json

```json
{
  "Enabled": true,
  "DefaultTimeoutSec": 0,
  "PollIntervalMs": 500,
  "SkipFlagPath": ".\\kernel\\json\\skip_request.flag"
}
```

| キー | 用途 |
|---|---|
| `Enabled` | `false` で `__ASYNC__` マーカーを silent ignore（kill switch） |
| `DefaultTimeoutSec` | `0`=無制限。`> 0` なら全 async モジュールにこの timeout を適用 |
| `PollIntervalMs` | Skip flag / completion / timeout チェックの polling 間隔（最小 50ms に clamp） |
| `SkipFlagPath` | Status Monitor が touch する flag ファイル絶対パス |

Kill switch を kernel 側に置いた理由: 緊急時に運用環境でも `Enabled=false` に書き換えるだけで全モジュールが従来の同期実行へ自動 fallback できる。

---

## Resolve-ProfileModules での解釈

```
profile CSV:
   Order, ScriptPath
   10,    standard/hostname_config/hostname_config.ps1
   20,    __ASYNC__                         ← この行は ValidModules に入らない
   30,    standard/winget_install/winget_install.ps1   ← _IsAsync=$true
   40,    standard/app_config/app_config.ps1            ← _IsAsync=$true（sticky）
   50,    __RESTART__
   60,    standard/reg_hklm_config/reg_hklm_config.ps1  ← _IsAsync=$true（profile 末尾まで継続）
```

`__ASYNC__` はプロファイル末尾までの sticky フラグ。途中で sync に戻すマーカーは無い（必要なら別プロファイルに分割）。

---

## Invoke-SafeCommandAsync の実装

`common.ps1` L1008、約 220 行。

### Runspace 立ち上げ

```powershell
$runspace = [runspacefactory]::CreateRunspace($Host)
$runspace.ApartmentState = "STA"           # WinForms 互換のため
$runspace.ThreadOptions  = "ReuseThread"
$runspace.Open()
$runspace.SessionStateProxy.Path.SetLocation($fabriqRoot)  # 相対パス整合
```

### Globals 注入

```powershell
$inject = @{
    FabriqMasterPassphrase = $global:FabriqMasterPassphrase
    AutoPilotMode          = $global:AutoPilotMode
    AutoPilotWaitSec       = $global:AutoPilotWaitSec
    FabriqTranscriptPath   = $global:FabriqTranscriptPath
    FabriqUniqueId         = $global:FabriqUniqueId
    FabriqSessionTimestamp = $global:FabriqSessionTimestamp
    FabriqEvidenceBasePath = $global:FabriqEvidenceBasePath
    FabriqEvidenceRootPath = $global:FabriqEvidenceRootPath
}
```

env vars は Process スコープなので Runspace に自動継承される。`$global:*` だけ明示注入。

### Wrapper Script Block

```powershell
$ps.AddScript({
    param($CommonPath, $ModuleScript, $Inject, $FabriqRoot)
    Set-Location -Path $FabriqRoot
    . $CommonPath                          # Show-Info etc を再ロード
    foreach ($key in $Inject.Keys) {
        Set-Variable -Name $key -Value $Inject[$key] -Scope Global -Force
    }
    $global:_LastModuleResult = $null
    $output = & $ModuleScript
    [PSCustomObject]@{
        _AsyncWrapper = $true
        Output        = $output
        LastResult    = $global:_LastModuleResult
    }
})
```

`_LastModuleResult` も pipeline と並列で回収するため、wrapper PSCustomObject 経由で親 Runspace に運ぶ。

### Monitor Loop

```powershell
$asyncHandle = $ps.BeginInvoke()
while (-not $asyncHandle.IsCompleted) {
    Start-Sleep -Milliseconds $pollMs

    if (Test-Path $skipFlagPath) {
        Remove-Item $skipFlagPath -Force
        $interrupted = $true
        $interruptReason = "Skip"
        try { $ps.Stop() } catch { }       # Runspace 強制停止
        break
    }

    if ($TimeoutSec -gt 0 -and ((Get-Date) - $startTime).TotalSeconds -ge $TimeoutSec) {
        $interrupted = $true
        $interruptReason = "Timeout"
        try { $ps.Stop() } catch { }
        break
    }
}
```

中断時の result.Message:

- Skip: `"Module skipped by operator (async runspace stopped; system state may be incomplete)"`
- Timeout: `"Module exceeded timeout of {N}s (async runspace stopped; system state may be incomplete)"`

「system state may be incomplete」を文字列に含めることで、Skip 後の状態が部分適用である可能性を operator に明示する。

### 完走時の捕捉ロジック

```powershell
$wrapper = $wrappedOutput | Where { $_._AsyncWrapper -eq $true } | Select -First 1
$moduleResult = $wrapper.Output | Where { $_._IsModuleResult -eq $true } | Select -First 1
if (-not $moduleResult -and $null -ne $wrapper.LastResult) {
    $moduleResult = $wrapper.LastResult     # フォールバック
}
```

`Invoke-SafeCommand` と同じセマンティクスを保つ。

### クリーンアップ

```powershell
finally {
    $result.Duration = (Get-Date) - $startTime
    if ($null -ne $ps)       { try { $ps.Dispose() } catch { } }
    if ($null -ne $runspace) {
        try { $runspace.Close() } catch { }
        try { $runspace.Dispose() } catch { }
    }
}
```

Skip / Timeout / 例外いずれの経路でも Runspace は確実に解放される。

---

## Skip Flag Path Normalization

`SkipFlagPath` を絶対パスに正規化してから polling する：

```powershell
if (-not [System.IO.Path]::IsPathRooted($skipFlagPath)) {
    $skipFlagPath = Join-Path (Get-Location).Path $skipFlagPath
}
$skipFlagPath = [System.IO.Path]::GetFullPath($skipFlagPath)
```

理由: Status Monitor は別プロセスで cwd が異なる可能性があり、絶対パスに揃えないと Skip 要求を取りこぼす。長時間 polling の途中で fabriq 側 cwd が変わるケースもあるため、loop 開始前に固定する。

---

## Skip フローの全体像

```
[ Status Monitor process ]
    operator が [Skip] ボタンを押す
       ↓
    New-Item -Path .\kernel\json\skip_request.flag -Force
       ↓
[ Main process / Invoke-SafeCommandAsync ]
    polling loop が Test-Path で検出
       ↓
    Remove-Item で flag を消す（次の async 実行のために）
       ↓
    $ps.Stop() で Runspace を強制終了
       ↓
    result.Status = "Error" + 上記の Message
       ↓
    Add-ExecutionResult / Write-ExecutionHistory で記録
       ↓
    AutoPilot 中なら ErrorMode 分岐（skip / retry / ask）
```

---

## 設計上の注意

### 1. 同期実行を default にした理由

- Runspace 起動オーバーヘッド（数十〜数百 ms / module）
- デバッグ性低下（Runspace 内例外のスタックトレースが分かりにくい）
- 互換性のリスク（一部モジュールが `$Host.UI.PromptForChoice` などホスト依存 API を使うと STA Runspace で崩れる）

→ 「ハングが懸念される一部のモジュールに対して opt-in」が現実解。

### 2. AutoPilot skip/timeout 機能の rejected 経緯

過去に「AutoPilot 全体に skip/timeout を適用する」案が検討されたが、Runspace refactor が広範すぎる vs ハング occurrence が稀という比較で **rejected**（2026-04-22、project memory `project_autopilot_skip_rejected`）。`__ASYNC__` の選択的適用が現行の落とし所。

### 3. Skip した状態の不完全性

Skip された Runspace は強制停止のため、モジュールが書きかけたレジストリ・ファイル・ネットワーク設定が中途半端に残る可能性がある。fabriq はこれを「Error として記録、operator に通知、続行は ErrorMode 次第」と扱う。後続モジュールが影響を受ける場合は profile 設計時に skip 不可能な順序を組む（domain_join の前に hostname_config を絶対に通すなど）。
