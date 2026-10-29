# fabriq アーキテクチャ調査 — Android 一括キッティングツールの「お手本」としての移植可否分析

調査対象: ローカルリポジトリ `/home/user/fabriq`（HEAD `928a4c9`、`kernel/KERNEL_VERSION` = `3.6.2`、調査日 2026-10-29）。リンク先はすべてリポジトリ相対パス（`path:行番号`）。Web 情報は使っていない。

**分類の凡例**（各節の Inferences で使う）:
- **(A) そのまま移植できる概念** — データ契約・規約・プロセスで、OS に依存しない。
- **(B) 作り直せば移植できる** — 考え方は活かせるが、「ホスト PC 1 台が USB/ADB で多数の Android 端末を並列に操作する」1 対多モデルに合わせて設計し直す必要があるもの。
- **(C) Windows 専用** — Android に相当するものが無い、または意味を持たない。

前提: Android 版は**ホスト PC（Windows）で動き、ADB 経由で N 台を並列操作する**ものとして評価した。fabriq は**対象 PC 自身の上で動く**（1 対 1）。端末上にエージェントアプリを置く案の差分は §11 にまとめた。Android 側のコマンド名・可否はすべて**推論（要検証）**で、このリポジトリからは確認できない。

---

## 1. カーネル: main.ps1 のセッションライフサイクルと common.ps1 の主要機構

### Takeaway
`kernel/main.ps1`（2,296 行）は、1 台の PC の上で 1 つの「セッション」を動かすプログラムである。流れは「transcript 開始 → 再開（resume）判定 → パスフレーズ・作業者・ホスト選択 → モジュール探索 → GUI メインループ → プロファイル実行エンジン（`Invoke-BatchExecution`）」。実際の機能は `kernel/common.ps1`（5,500 行・114 関数）に集まっている。中核の契約（結果オブジェクト・CSV ローダ・ログ）は OS に依存しない。一方、セッション状態の多くは**プロセス全体で 1 つのグローバル変数／環境変数**として持たれている。そのため、1 プロセスで N 台を並列に動かすモデルにはそのまま載せられない。

### Cited Findings
**起動とセッション開始（main.ps1）**
- 起動処理: `common.ps1`、`manifesto.ps1`、`apps/fabriq_operator/fabriq_operator.ps1` を dot-source する。GUI は必須で、`-NoGui` は廃止済み（GUI が読めなければ exit 1）。その後 `Enable-SleepSuppression` / `Disable-QuickEditMode` / `Set-ConsoleSize` を呼ぶ — [kernel/main.ps1:7-45](kernel/main.ps1)
- 定数: `$HOSTLIST_CSV = .\kernel\csv\hostlist.csv`、`$script:AutoPilotMaxRetry = 5`、パスフレーズ検証トークン `kernel\txt\passphrase_verify.txt`（Studio が生成）— [kernel/main.ps1:50-58](kernel/main.ps1)
- Transcript のファイル名は `logs\{yyyy_MM_dd_HHmmss}_{HardwareUniqueId}_{COMPUTERNAME}.log`。`$global:FabriqUniqueId` と `$global:FabriqSessionTimestamp` をここで確定（bake）する — [kernel/main.ps1:1397-1409](kernel/main.ps1)
- **再開判定はパスフレーズ入力より前に行う**。①`modules/standard/windows_update/wu_state.json` があれば Windows Update ループを即再開する（パスフレーズ・作業者・ホスト選択は不要）。②`resume_state.json` があればプロファイル再開。`schemaVersion>=2` かつ `ExecutionMode='Flex'` のときは Flex 再開になる — [kernel/main.ps1:1420-1450](kernel/main.ps1)
- AutoPilot 状態で再開するときは、`Wait-SystemReady` のあと `Invoke-AutoResumeCountdown -Seconds 60`（Enter で即時、Esc で中止）を挟む。AutoPilot でなければ Y/N で確認する。再開時は `Restore-HostEnvironment` と DPAPI で、パスフレーズと PIN を復元する — [kernel/main.ps1:1468-1545](kernel/main.ps1)、[kernel/common.ps1:5242](kernel/common.ps1)
- 新規セッション: 検証トークンが無ければ exit 1。`hostlist.csv` と `workers.csv` を読み込み、`Show-SessionSetupForm`（作業者・ホスト・パスフレーズを入力する WinForms 画面）を出す。そのあと `session.json`（WorkerID/WorkerName/MediaSerial=USB ボリュームシリアル/StartTime/WindowsUser/ComputerName）を保存し、`Set-SelectedHostEnvironment` と `Initialize-EvidenceBasePath` を呼ぶ — [kernel/main.ps1:1593-1664](kernel/main.ps1)
- 新規起動時は `logs\history\execution_history.csv` を削除する（再開時は残す）。その後 `Initialize-ExecutionHistory` → `Restore-ExecutionHistory` → `Initialize-ModuleSystem`（`module.csv` を走査）→ verbose 捕捉 ON → Execution Toolbar（同一プロセス内の STA Runspace）を表示 — [kernel/main.ps1:1666-1749](kernel/main.ps1)
- GUI メインループ: `Show-OperatorDashboard` が返す Action で分岐する。Action は `FlexProfile` / `ExecuteProfile`（Linear）/ `ExecuteModules` / `NewSession` / `OpenCsvEditor` / `SystemLauncher` / `AppsMode` / `LaunchApp` / `HistoryExport` / `RegenerateChecklist` / `WindowsUpdate` / `Restart` / `Refabriq` / `Manifesto` / `Quit` — [kernel/main.ps1:1979-2284](kernel/main.ps1)
- main.ps1 内の関数は 9 個: `Load-HostList`(L63)、`Set-SelectedHostEnvironment`(L86)、`Invoke-KittingScript`(L216)、`Invoke-BatchExecution`(L304)、`Invoke-FlexProfileLoop`(L831)、`Set-/Clear-WindowsUpdateAutoLogon`(L1102/L1142)、`Invoke-WindowsUpdateLoop`(L1161)、`Show-WindowsUpdateSummary`(L1345) — [kernel/main.ps1](kernel/main.ps1)

**実行エンジン（Invoke-BatchExecution, main.ps1:304-811）**
- 1 モジュールごとの流れ: `__GATE__` 判定(L483) → `__RESTART__`(L504) → `__REEXPLORER__`(L562) → 通常モジュール。AutoPilot 時はモジュール間で `WaitSec` 秒待つ。`_AutoLogonUser` は `$env:FABRIQ_AUTOLOGON_USER` へ、`_Segment` は `$env:FABRIQ_SEGMENT` へ入れて実行し、終わったら解除する。async なら `Invoke-SafeCommandAsync`(L640)、そうでなければ `Invoke-SafeCommand` で実行する。Error/Partial のときは ErrorMode 分岐(L666-690)。そのあと `Add-ExecutionResult` → `Write-ExecutionHistory` → `Capture-ScreenEvidence`(L704) — [kernel/main.ps1:470-725](kernel/main.ps1)
- 完走時: `Remove-ResumeState`、テレメトリ `profile.end`、`Show-ExecutionSummary` を実行し、プロファイル実行なら `Complete-ProfileExecution -Mode Auto`（履歴 export → HTML チェックリスト → log_uploader → ビューア）。`finally` で AutoPilot フラグ・データオーバーレイ文脈・テレメトリ文脈を必ずリセットする — [kernel/main.ps1:730-811](kernel/main.ps1)
- `Invoke-KittingScript`（単発実行経路）は fail-closed: ModuleResult が返らなければ Error として記録する — [kernel/main.ps1:216-300](kernel/main.ps1)

**common.ps1 の主要機構（行番号は `^function` の位置）**
- Show-* 系(L861-900): `Get-FabriqMaskedText` で秘密値をマスク → `Write-Host "[INFO]…"` → `Write-ArtPulse`（`kernel\json\art_pulse.txt` のカウンタを +1。生存シグナル）→ `_TrackShowEvent`（テレメトリ）。`Show-Separator`(L812) / `Show-CategorySeparator`(L816) も同系統 — [kernel/common.ps1:812-900](kernel/common.ps1)
- 結果契約: `New-ModuleResult`(L905) は `{_IsModuleResult, Status(Success/Error/Cancelled/Skipped/Partial), Message, Details, Verified(Nullable[bool]), Timestamp}` を返し、同じものを `$global:_LastModuleResult` にも退避する。`New-BatchResult`(L935) は件数を集計表示してから Status を自動判定する（Fail=0かつSuccess>0→Success / 両方>0→Partial / Skipのみ→Skipped / Failのみ→Error / それ以外→Success）。`Confirm-ModuleExecution`(L981) は N のとき Cancelled を返す — [kernel/common.ps1:905-996](kernel/common.ps1)
- `Invoke-SafeCommand`(L1877): テレメトリ envelope の開始と終了、verbose ストリーム（`4>&1`）のテレメトリ化、パイプライン出力から `_IsModuleResult` を探す処理を担う。見つからなければ `$global:_LastModuleResult` で救済し、それも無ければ Error として記録する（fail-closed）— [kernel/common.ps1:1877-2029](kernel/common.ps1)
- `Invoke-SafeCommandAsync`(L2061): STA Runspace の中で common.ps1 を読み直し、親プロセスのグローバル変数（パスフレーズ・AutoPilot・エビデンスパス・テレメトリ envelope）を注入する。監視ループは `skip_request.flag` とタイムアウトを見ていて、Skip でもタイムアウトでも `Status=Error`（「system state may be incomplete」）になる。`async_config.json` の現状値は `Enabled=true, DefaultAsync=true, DefaultTimeoutSec=0` — [kernel/common.ps1:2031-2317](kernel/common.ps1)、[kernel/json/async_config.json](kernel/json/async_config.json)
- CSV ローダ: `Import-CsvSafe`(L1779) / `Test-CsvColumns`(L1804) / `Import-ModuleCsv`(L1435)。`Import-ModuleCsv` の処理順は、オーバーレイ解決 → `ENC:` の透過復号（復号した値は秘密値として登録）→ 必須列の検証 → Enabled="1" フィルタ → Segment 厳密一致 → テレメトリ `csv.load`。戻り値は 3 通り: ロード失敗は `$null`、対象ゼロは `,@()`（空配列）、それ以外は常に配列 — [kernel/common.ps1:1435-1546](kernel/common.ps1)
- 暗号: `Unprotect-FabriqValue`(L1042) は AES-256-CBC/PKCS7。鍵と IV は PBKDF2-HMAC-SHA256（10 万回）で導出し、**ソルトは固定の `"fabriq-fixed-salt-2024"`**。`Test-MasterPassphrase`(L1105) は、トークンを復号して `"surkitinisme"` と一致すれば正しいパスフレーズと判定する。再開用のパスフレーズと PIN は DPAPI（LocalMachine）で保護する(L1000-1040)。秘密値レジストリ(L1157-) は Show-* / テレメトリ / 実行履歴の 3 箇所の出力すべてで `***` マスクする — [kernel/common.ps1:1000-1200](kernel/common.ps1)
- 再起動跨ぎ: `Save-ResumeState`(L3641) が JSON を保存する。中身は ProfilePath/ProfileName/ProfileDataDir/AutoPilot/WaitSec/SessionID/HardwareUniqueId/ResumeAfterOrder/CompletedModules/HostEnvironment(SELECTED_* 全部)/EvidenceBasePath/ProfileStartTime/ProtectedPassphrase/ProtectedPin、Flex のときは schemaVersion=2 と SelectedOrders・ModuleStates。`Load-ResumeState`(L3764) は **HardwareUniqueId が違えば拒否する**（共有 USB で別 PC の状態を誤って使わないため。ファイルは消さない）。`Register-FabriqRunOnce`(L5048) は `HKLM\…\RunOnce\FabriqAutoStart = Fabriq.exe` を書く — [kernel/common.ps1:3641-3810, 5048-5074](kernel/common.ps1)
- `__RESTART__` で RunOnce の登録に失敗したら、再起動せずに残りを中止して Error を記録する（fail-closed、TM t-0045）— [kernel/main.ps1:536-548](kernel/main.ps1)
- ハードウェア ID: `Get-HardwareUniqueId`(L1676) は BIOS シリアルを優先する（"To be filled by O.E.M." 等の無効値リストあり）。使えなければ Up 状態の物理 NIC の MAC、それも無ければ `"UNKNOWN"` — [kernel/common.ps1:1676-1720](kernel/common.ps1)
- 「次の顧客」用のリセット: `Reset-FabriqState`(L3810) は同じプロセスのまま新しいセッションを始めるための関数（transcript の張り替え等。エビデンスは消さない）— [kernel/common.ps1:3810](kernel/common.ps1)
- そのほか: `Get-FabriqGateBarrier`(L4335)、`Test-FabriqProtectedPath`(L4553)、`Write-StatusFile`(L4807)、`Initialize-ModuleSystem`(L4965。categories.csv と各 `module.csv` を走査し `Enabled=0` は除外)、`Capture-ScreenEvidence`(L5340)、`Save-Screenshot`(L5424)、`Invoke-CountdownRestart`(L5226)、`Show-AutoPilotErrorDialog`(L239)、テレメトリ群(L338-811) — [kernel/common.ps1](kernel/common.ps1)
- セッション関連のパス: `$script:SessionID = yyyyMMdd_HHmmss`、`HistoryPath = .\logs\history\execution_history.csv`、`ResumeStatePath = .\kernel\json\resume_state.json`、`WorkersCsvPath = .\kernel\csv\workers.csv` — [kernel/common.ps1:72-80](kernel/common.ps1)

### Inferences
- **(A)** `New-ModuleResult` / `New-BatchResult` の 5 値 Status と `Verified` の三値、件数から Status を決めるルール、pipeline 優先・global 救済・無返却は Error という fail-closed の結果回収。どれも純粋なデータ契約なので、bash でも C# でも Kotlin でもそのまま再実装できる（fabriq_mac の OK/NG/SKIP を書いた結果ファイルは、これを簡略化したもの）。
- **(A)** `Import-ModuleCsv` の「ロード失敗は `$null` → Error、対象ゼロは空配列 → Skipped」の区別。PowerShell 特有の配列 unroll 対策を除けば言語非依存で、移植先でも単体テストで固定すべき仕様。
- **(B)** グローバル状態がプロセスに 1 つしか無い問題。`$env:SELECTED_*` / `$env:FABRIQ_SEGMENT` / `$global:AutoPilotMode` / `$global:FabriqEvidenceBasePath` / `$global:_LastModuleResult` はすべて「1 プロセス = 1 台」を前提にしている（main.ps1 L620-660 で実行のたびに設定・解除）。ホスト側で N 台を並列に動かすには、(a) **端末ごとにワーカプロセスを分ける**（環境変数による契約を保ったまま並列化できる。fabriq_mac 流の `HOST_<column>` をそのまま使える）か、(b) モジュールに端末コンテキストを明示的に渡すか、のどちらかが必要。(a) が fabriq 流に最も近い。
- **(B)** async Runspace と Skip フラグとタイムアウト（`Invoke-SafeCommandAsync`）は、1 対多モデルでは**常に有効な端末単位のキャンセル・タイムアウト**として作り直すのが自然。`DefaultTimeoutSec=0`（無制限）のままでは、1 台が止まると監視も止まりかねない。`adb` 呼び出しごとに上限時間を設けるべき（無期限待ちは fabriq でも既知の痛点。§12）。
- **(B)** 再開判定・DPAPI・RunOnce・AutoLogon・Windows Update ループは、「対象機の上で動いていて、自分が再起動で死ぬ」ことへの対策である。ホスト PC モデルでは**ホストは死なない**ので、端末の再起動は「`adb reboot` → 再接続を待つ → 起動完了をポーリング」という単なる待ち状態になる（要検証: `sys.boot_completed`）。resume_state 相当が必要なのは、ホスト自身が落ちた場合の再開だけ。HardwareUniqueId を照合して別機の状態を拒否する考え方は、**端末シリアルとの照合**として (A) で残すべき。
- **(C)** `Enable-SleepSuppression`、QuickEdit 無効化、コンソールサイズ、`Remove-ZoneIdentifier`、Explorer 再起動はすべて Windows 固有。ホストの Windows で GUI ツールを作るなら、ホスト側の事情（スリープ抑止など）としては一部残る。

### Gaps
- `Show-OperatorDashboard` / `Show-SessionSetupForm` の UI の詳細（549 行と 400 行）は、主な分岐を確認しただけで、全部は読んでいない。
- テレメトリ（`Start-ModuleTelemetry` 等、L338-811）の内部実装は、`dev/TELEMETRY_INTERNAL.md` の章立てと保存先の確認までに留めた。

---

## 2. カーネル公開 API（KERNEL_API.md）の目的別一覧と導入版

### Takeaway
公開 API は**関数 32 個・公開グローバル変数 6 個・公開環境変数・Profile CSV スキーマ・特殊マーカー 6 種・ModuleResult 契約**で、「ここに書かれたものだけにモジュールが依存してよい」という形で厳密に宣言されている。導入版は表（§8）で逆引きできる。Android 版でも「モジュールから見えるカーネル面」を最初に明文化して固定する、という方式そのものが最も移植価値の高い資産である。

### Cited Findings
- 方針: 公開 API として宣言された要素だけがモジュール依存の対象。それ以外は内部扱いで、PATCH でも予告なく変えてよい。公開 API を変えるときは、同じコミットで KERNEL_VERSION を MINOR 以上に上げる — [kernel/KERNEL_API.md:5-7](kernel/KERNEL_API.md)
- **§1.1 表示・通知（2.0.0）**: `Show-Info` / `Show-Success` / `Show-Warning` / `Show-Error` / `Show-Skip` / `Show-Separator` / `Show-CategorySeparator` — [kernel/KERNEL_API.md:13-22](kernel/KERNEL_API.md)
- **§1.2 CSV 読み込み**: `Import-ModuleCsv -Path [-FilterEnabled] [-RequiredColumns] [-Segment]`（2.0.0）、`Resolve-ModuleDataPath -Path [-ForWrite]`（3.7.0 [Unreleased]）、`Get-ModuleDataFiles -Directory -Filter`（3.7.0 [Unreleased]）。契約: Segment は `$env:FABRIQ_SEGMENT` と厳密に一致させる（空同士も一致）。戻り値は失敗=`$null`、対象ゼロ=空配列 — [kernel/KERNEL_API.md:24-35](kernel/KERNEL_API.md)
- **§1.3 結果返却（2.0.0）**: `New-ModuleResult` / `New-BatchResult` / `Confirm-ModuleExecution`。`Verified` は**範囲付き（scoped）**の意味を持つ。「1 件も検証していないのに `$true`」は禁止（その場合は `$null`）。対象が検証不能な形で欠けているときは `$false` に倒すことを推奨（fail-closed）。`Fail>0` と `Verified=$true` は両立し得るので、利用側は Status と Verified を必ず両方見る — [kernel/KERNEL_API.md:37-48](kernel/KERNEL_API.md)
- **§1.4 確認・待機（2.0.0）**: `Confirm-Execution`（AutoPilot 中は自動で Y）/ `Wait-KeyPress` / `Wait-NetworkReady` — [kernel/KERNEL_API.md:50-55](kernel/KERNEL_API.md)
- **§1.5 権限・環境（2.0.0）**: `Test-AdminPrivilege`、`Unprotect-FabriqValue` — [kernel/KERNEL_API.md:57-61](kernel/KERNEL_API.md)
- **§1.6 破壊的削除ガード（3.5.0）**: `Test-FabriqProtectedPath`（保護ルート・その親・3 セグメント未満・`\\?\`・UNC を fail-closed で拒否）、`Test-FabriqSafePathComponent`。ガードは確認ゲートの**外側**（AutoPilot でも効く位置）に置く — [kernel/KERNEL_API.md:63-69](kernel/KERNEL_API.md)
- **§1.7–1.10（2.0.0 として追認宣言）**: `Expand-UserEnvironmentVariables`、`Resolve-HkcuRoot`、`Get-HardwareUniqueId`、`Wait-SystemReady`（タイムアウトしても必ず戻る）、`Register-FabriqActiveSetup`、`Deploy-FabriqUserSetupLauncher`、`Deploy-FabriqStartupTrigger`、`Invoke-CountdownRestart`（戻らない）、`Invoke-CountdownSignout`、`Remove-ZoneIdentifier`、`Capture-ScreenEvidence`。いずれも 2026-07-02 の監査で「実際には使われているのに未宣言」と判明し、追認した関数 — [kernel/KERNEL_API.md:71-104](kernel/KERNEL_API.md)
- **§1.11 UI（3.7.0 [Unreleased]）**: `New-UiFont`（DPI モードに依存しない WinForms フォント）— [kernel/KERNEL_API.md:106-110](kernel/KERNEL_API.md)
- **§2 公開グローバル変数**: `$global:FabriqMasterPassphrase` / `AutoPilotMode` / `AutoPilotWaitSec` / `FabriqEvidenceBasePath`（2.0.0）、`FabriqUniqueId`（2.0.0 として追認）、`AutoConfirmMode`（3.1.0）。書き込みが許されるのは `$global:_LastModuleResult` だけ — [kernel/KERNEL_API.md:114-125](kernel/KERNEL_API.md)
- **§3 公開環境変数**: `SELECTED_KANRI_NO/OLD_PCNAME/NEW_PCNAME/ETH_{IP,SUBNET,GATEWAY}/WIFI_{…}/DNS1-4/PIN`、`SELECTED_PRINTER_<1..10>_{NAME,DRIVER,PORT}`、`FABRIQ_SEGMENT` / `FABRIQ_AUTOLOGON_USER` / `FABRIQ_WORKER_NAME` / `FABRIQ_EVIDENCE_BASE`（2.0.0）、`FABRIQ_PROFILE_DATA_DIR`（3.7.0）。`__SELF__` トークン（3.5.0）は hostlist のセルを実行中 PC の実際の値に解決し、入室時に 1 回だけ確定（bake）する — [kernel/KERNEL_API.md:129-146](kernel/KERNEL_API.md)
- **§4 Profile CSV / マーカー**（§4 節を参照）、**§5 ModuleResult**（2.0.0 で全フィールド）— [kernel/KERNEL_API.md:150-193](kernel/KERNEL_API.md)
- **§6 内部実装（API ではない）**: `Invoke-SafeCommand(Async)` / `Invoke-BatchExecution` / `Resolve-ProfileModules` / Save・Load-ResumeState / 実行履歴・HTML 系 / Execution Toolbar / 状態ファイル（`resume_state.json`, `status.json`, `session.json`, `art_pulse.txt`, `async_config.json`, `skip_request.flag`）— [kernel/KERNEL_API.md:197-218](kernel/KERNEL_API.md)
- **§8 版の歴史**: 2.0.0 baseline、2.1.0 `__ASYNC__`、2.2.0 更新・オーバーレイ契約（§9）とモジュール VERSION/REQUIRES_KERNEL の一斉 seed、2.2.2 Evidence Manifest 契約（§10）、3.0.0 マーカー 4 種（`__SHUTDOWN__`/`__PAUSE__`/`__STOPLOG__`/`__STARTLOG__`）を削除（MAJOR）、3.1.0 `AutoConfirmMode` と FlexProfile、3.2.0 `Group` 列、3.3.0 `DefaultAsync`、3.4.0 Status Monitor 廃止 → Execution Toolbar、3.5.0 `__SELF__` と削除ガード、3.6.0 `__GATE__`、3.7.0（Unreleased）オーバーレイ・`New-UiFont`・AUTO 行の Segment 伝播 — [kernel/KERNEL_API.md:235-440](kernel/KERNEL_API.md)
- **§9 更新・オーバーレイ契約**: `dev/framework_overlay_rules.json`（schemaVersion 1）が唯一の真実源。kernel bundle とモジュール bundle を SemVer で比較する。`profiles/`・hostlist・`_list.csv` 等の現場固有データは絶対に上書きしない。モジュール更新の前に `REQUIRES_KERNEL` を事前チェックする — [kernel/KERNEL_API.md:454-530](kernel/KERNEL_API.md)
- **§10 Evidence Manifest 契約**（§6 節を参照）— [kernel/KERNEL_API.md:534-580](kernel/KERNEL_API.md)
- 記述のずれ: KERNEL_VERSION の実ファイルは `3.6.2` だが、§8 には 3.7.0 の追加が「[Unreleased]・実ファイル昇格はリリース指示時」と書かれている。複数のモジュールが既に `REQUIRES_KERNEL 3.7.0` を宣言している（例: acl_config, app_config, autologon_config）— [kernel/KERNEL_VERSION](kernel/KERNEL_VERSION)、[kernel/KERNEL_API.md:392](kernel/KERNEL_API.md)、[modules/standard/app_config/REQUIRES_KERNEL](modules/standard/app_config/REQUIRES_KERNEL)

### Inferences
- **(A)** 公開 API の明文化（公開と内部の区別・導入版の表・`REQUIRES_KERNEL` = 使っている API の導入版の最大値）は、OS に依存しない**ガバナンスの型**である。Android 版でも初版（例 0.1.0）から `KERNEL_API.md` 相当を置くべき。
- **(A)** `Verified` の「範囲付き三値」の規約（未検証は `$null`、ターゲットが欠けていれば `$false`）は、多数の端末を一覧表で並べる 1 対多モデルでこそ効く（端末ごと・ステップごとの OK/NG/未検証の 3 色表示）。
- **(B)** §3 の `SELECTED_*` は PC 向けの語彙（PC 名・NIC・DNS・プリンタ）。Android では、端末名・資産番号・Wi-Fi SSID・（必要なら）Google アカウント・ラベル等に**語彙を置き換える**必要がある。`__SELF__`（実機の値で自己解決する）は「hostlist に無い端末も、その場でシリアルや機種から受け入れる」機能として有用。
- **(C)** §1.7（HKCU/USERPROFILE の解決）と §1.9（Active Setup / Startup）は Windows のユーザープロファイル構造に固有。Android のマルチユーザーは業務端末ではほぼ使われないと見込む（要検証）。

### Gaps
- `KERNEL_API.md` にある 3.7.0 の各機能が、実ファイルの版表記より先に「公開 API」として運用されている状態（Unreleased）の正確な扱いは、CHANGELOG の `[Unreleased]` を全部読まないと確定できない（未実施。CHANGELOG は 6,265 行）。

---

## 3. モジュール契約（dev/template・冪等性・Post-Apply Verification・Segment/Enabled）

### Takeaway
モジュールは「フォルダ = モジュール」で、中身は `module.csv`（メニュー登録）、`<name>.ps1`、`*_list.csv`（設定データ）、`preset.csv`（Studio のドロップダウン）、`VERSION`、`REQUIRES_KERNEL`、`test.psd1`（テストリグ用の記述子）。スクリプトは**Step 1〜6 の固定フロー**（CSV 読込 → 前提確認 → dry-run 表示 → 確認 → 適用ループ → 5.5 読み返し検証 → 集計して返却）。この構造は fabriq_mac の「csv_load → 行ごと count_skip / 適用して読み返し count_ok|count_ng → finish」とほぼ同型で、Android 版にもそのまま移せる。

### Cited Findings
- `module.csv` の列: `MenuName,Category,Script,Order,Enabled`（テンプレ例: `Template Module,System,_template_script.ps1,99,1`）— [dev/template/module.csv](dev/template/module.csv)
- `_template_list.csv` の列: `Enabled,TargetName,Description,Segment`（Enabled 1=実行 / 0=スキップ）— [dev/template/_template_list.csv](dev/template/_template_list.csv)
- `VERSION` = `0.1.0`（開発中の目印）、`REQUIRES_KERNEL` = `3.5.0` — [dev/template/VERSION](dev/template/VERSION)、[dev/template/REQUIRES_KERNEL](dev/template/REQUIRES_KERNEL)。ただし CLAUDE.md は「REQUIRES_KERNEL = 現行カーネル版」と定めており、現行は 3.6.2 で食い違っている — [CLAUDE.md](CLAUDE.md)（§F）
- テンプレの流れ: ヘッダ（`[PURPOSE]` / `[NOTES]`）→ バナー（`Show-Separator` と `Write-Host`）→ 任意で P/Invoke ブロック → **Step 1** `Import-ModuleCsv -FilterEnabled -RequiredColumns`（`$null` なら Error、Count 0 なら Skipped）→ **Step 2** 前提確認で早期 return → **Step 3** dry-run 表示（`[APPLY]/[SKIP]/[NOT FOUND]` を色分け）→ **Step 4** `Confirm-ModuleExecution`（AutoPilot なら自動 Y）→ **Step 5** 適用ループ（行ごとの前提確認は try の前で `Show-Skip; $skipCount++; continue`、本処理は try/catch で `$successCount/$failCount`）→ **Step 5.5** 読み返し検証（`$verifyPass/$verifyFail` → `$verified`）→ **Step 6** `New-BatchResult -Success -Skip -Fail -Title [-Verified]` — [dev/template/_template_script.ps1:1-213](dev/template/_template_script.ps1)
- Guide.txt: テンプレのコピー手順と命名規則（フォルダ・スクリプト名は snake_case で揃える）。「Step 構成（7 ステップ）」と書かれている — [dev/template/Guide.txt](dev/template/Guide.txt)
- **Segment の記述のずれ**: Guide.txt は「Segment 空欄 = 全セグメント共通」と書く。しかし実装は**厳密一致**で、空欄の行は Segment が空のときにしか一致しない（`$rowSegment -eq $effectiveSegment`）。KERNEL_API と README も「厳密マッチ」と書いている — [dev/template/Guide.txt](dev/template/Guide.txt)、[kernel/common.ps1:1496-1515](kernel/common.ps1)、[README.md:66](README.md)
- `test.psd1`（C2 契約・schema 1）: `category`（A=headless で可 / B=fixture が必要 / C=外部インフラ / D=破壊的・revert 必要 / E=`__RESTART__` 跨ぎ）と、シナリオごとの `script`、`context`（noninteractive=WinRM Session 0 / interactive）、`winrmSafe`、`reboot`、`secrets`、`envelope`（autopilot/passphrase/selected/segment）、`fixture`、`expect{status,verified}`、`oracle`（registry-csv / file-exists / state-query / command / self-verified / none）、`idempotent{secondRun='Skipped'}`、`cleanup`（none/undo/snapshot）、`teardown` — [dev/template/test.psd1](dev/template/test.psd1)
- `preset.csv` の列は `Column,Value,Label`（UTF-8 BOM + CRLF）。Studio が列挙型の列を ComboBox で編集するために使う — [README.md:292-311](README.md)、[modules/standard/power_config/preset.csv](modules/standard/power_config/preset.csv)
- **冪等性 skip の手本**:
  - `power_config`: GUID ごとに Getter で現在値を読み、一致すれば `Show-Skip "… already set"`（L185, 334-403, 470-754）— [modules/standard/power_config/power_config.ps1](modules/standard/power_config/power_config.ps1)
  - `windows_feature_config`: `Test-FeatureStateApplied` を **Step 5 の skip と Step 5.5 の検証の両方**で使う。Pending 状態も「適用済み」と数える（L99-111, 231）— [modules/standard/windows_feature_config/windows_feature_config.ps1](modules/standard/windows_feature_config/windows_feature_config.ps1)
  - `volume_config`: 既に目標状態なら `New-ModuleResult -Status Skipped -Verified $true` で早期 return。Step 5.5 は実際の音量と Mute を読み返す（L188-195, 249-279）— [modules/standard/volume_config/volume_config.ps1](modules/standard/volume_config/volume_config.ps1)
- **検証の手本** `reg_hklm_config` の `Test-RegistryValueMatch`: 型別に比較する（DWord/QWord/Binary(hex)/MultiString/ExpandString は未展開値で比較）。ただし `$FORCE_OVERWRITE = $true` が既定なので、適用ループでの skip は無効になっている（L5, L80-130, L252）。preview と verify 専用の関数であり、冪等性 skip の手本ではない — [modules/standard/reg_hklm_config/reg_hklm_config.ps1](modules/standard/reg_hklm_config/reg_hklm_config.ps1)、[CLAUDE.md](CLAUDE.md)（§6）
- 再起動後に効く設定の検証例: `hostname_config` は、保留中のホスト名（`HKLM\SYSTEM\…\ComputerName\ComputerName`）を読み返して `-Verified` を決める（L67-91）— [modules/standard/hostname_config/hostname_config.ps1](modules/standard/hostname_config/hostname_config.ps1)
- テンプレから意図的に外れる例: `hostlist_precheck` は Step 4 の確認を省く（Cancelled だと `__GATE__` を素通りするため）。Step 6 も `New-ModuleResult` を使う（部分的な失敗も Error にするため）— [modules/standard/hostlist_precheck/hostlist_precheck.ps1:1-30](modules/standard/hostlist_precheck/hostlist_precheck.ps1)

### Inferences
- **(A)** Step 1〜6 の骨格、Enabled 列、Segment 列（厳密一致）、`RequiredColumns` での列検証、確認ゲート、件数集計 → Status 自動判定、Step 5.5 の読み返し検証。これらは Android 版のモジュール契約にそのまま使える。fabriq_mac が既に同型であることも根拠になる。Segment の空欄の意味は Guide.txt と実装で食い違っているので、**新ツールでは最初から 1 つに決めておく**べき。
- **(A)** 「同じ判定関数を、冪等性 skip（Step 5）と検証（Step 5.5）の両方に使う」パターン（windows_feature_config）は、Android でも `settings get` / `pm list packages` / `dumpsys` の読み取りで適用前の判定と適用後の検証を兼ねる形にそのまま写せる（要検証）。
- **(B)** 1 対多モデルでの Step 3（dry-run 表示）と Step 4（確認）は、**端末ごとではなくロット全体で 1 回**にすべき。端末 N 台 × モジュール M 個の Y/N は運用できない。FlexProfile が既に「実行は常に AutoPilot 動作、完了処理は常に手動」に寄せている（§4）のが先例になる。
- **(B)** `test.psd1` の考え方（独立した oracle・2 回目の実行は Skipped を期待・cleanup の種別・カテゴリ A〜E）は、Android エミュレータ（AVD）や検証用の実機で移植できる。Android では「snapshot で復元」が AVD のスナップショットや工場出荷リセットに当たる（要検証）。
- **(C)** テンプレの P/Invoke ブロックや、`reg_hklm` 固有の型比較は Windows 専用。ただし「値を型付きで比較するヘルパ」という考え方は、Android の `settings`（すべて文字列）や `getprop` では単純な文字列比較で済む見込み。

### Gaps
- 冪等性の実装状況を全モジュール分まとめた台帳は、TM t-0082 で「未着手」— [.tm/tasks.json](.tm/tasks.json)。どのモジュールが冪等 skip を持つかを網羅的に確認することはしていない。

---

## 4. プロファイル・特殊マーカー・AutoPilot/ErrorMode/Group・FlexProfile・再起動跨ぎ・async

### Takeaway
プロファイルは「順序付きのモジュール一覧 CSV」（`Order,ScriptPath,Enabled[,Description,Segment,ErrorMode,Group]`）で、実行時の制御を 6 種の特殊マーカーで差し込む。FlexProfile は、実行履歴からモジュールごとの状態（Success/Partial/Error/Skipped/Pending）を復元して一覧に出し、部分的に実行できる GUI である。これは fabriq_mac の「電子チェックシート」の原型に当たる。`__GATE__`（前段に Error/Partial/検証 NG が残る間は先へ進ませない）とともに、1 対多のマトリクス UI へ最も素直に拡張できる部分である。

### Cited Findings
- 列の定義: `Order`（必須・昇順）、`ScriptPath`（必須。`{standard,extended}/<module>/<script>.ps1` か特殊マーカー）、`Enabled`（必須）、`Description`、`Segment`、`ErrorMode`（空 / `skip` / `retry`）、`Group`（3.2.0。FlexProfile の `[Run: <Group>]`。Linear では無視）— [kernel/KERNEL_API.md:150-161](kernel/KERNEL_API.md)
- マーカーの定義: `__AUTOPILOT__`（Description に `WaitSec=N`）、`__ASYNC__`、`__RESTART__`、`__REEXPLORER__`、`__GATE__`（3.6.0）、`__AUTO_to_<User>__` — [kernel/KERNEL_API.md:163-171](kernel/KERNEL_API.md)
- 解析の実装: `Resolve-ProfileModules` は `Import-Csv` → Enabled="1"（Flex は `-IncludeDisabled`）→ `[int]Order` でソート → マーカー判定の順で処理する。`__AUTOPILOT__` はプロファイル全体の AutoPilot フラグと WaitSec を返す（位置は判定に使わない）。`__ASYNC__` は以降の `_IsAsync`（DefaultAsync=true なら何もしない）。`__AUTO_to_X__` は `autologon_config` を複製して `_AutoLogonUser` を付ける。`__RESTART__` / `__REEXPLORER__` / `__GATE__` はフラグ付きの擬似モジュールになる。通常行には `_Segment` / `_ErrorMode` / `_IsAsync` / `_Group` を付ける。見つからないパスは `InvalidPaths` に集め、警告だけ出して続行する — [kernel/common.ps1:4149-4333](kernel/common.ps1)
- Linear GUI 経路の AutoPilot は、ダッシュボードの `AutoPilot` チェックボックス（既定 ON）から渡り、WaitSec は既定 3 秒（`dashboard_form.ps1:24, 135, 417` → `main.ps1:2021`）。この経路で `__AUTOPILOT__` マーカーの戻り値が使われている箇所は見つからなかった。README は「`__AUTOPILOT__` マーカー以降は確認をスキップ」と書いている — [apps/fabriq_operator/lib/dashboard_form.ps1:135-417](apps/fabriq_operator/lib/dashboard_form.ps1)、[kernel/main.ps1:2007-2030](kernel/main.ps1)、[README.md:59](README.md)
- ErrorMode（AutoPilot 中に Error/Partial になったとき）: `skip` は記録して続行。`retry` は最大 5 回（`$script:AutoPilotMaxRetry`）を WaitSec 間隔で自動再試行。空欄は `Show-AutoPilotErrorDialog`（Retry/Cancel の **タイムアウト無し MessageBox**）— [kernel/main.ps1:660-695](kernel/main.ps1)、[kernel/common.ps1:239-265](kernel/common.ps1)
- `__RESTART__`: `Save-ResumeState` → `Register-FabriqRunOnce`（失敗したら中止）→ 履歴へ `[RESTART]` を記録 → Toolbar を隠す → `Invoke-CountdownRestart`（5 秒後に `Restart-Computer -Force`）。プロファイル外の一括実行ではマーカーを無視する — [kernel/main.ps1:504-560](kernel/main.ps1)、[kernel/common.ps1:5226-5240](kernel/common.ps1)
- Linear の再開: `Resolve-ProfileModules` をやり直し、`Order > ResumeAfterOrder` の行だけを実行する。プロファイルデータフォルダは `Test-FabriqResumeDataDir` で fail-closed 検証する — [kernel/main.ps1:1751-1850](kernel/main.ps1)
- `__GATE__`: 直前のゲートから当該ゲートまでの範囲（窓）に Error/Partial または `Verified=False` が残っていれば、それより後ろの Order を拒否して Pending のまま据え置く。判定は毎モジュールの直前に、当該 run とセッション履歴（同じ Profile かつ同じ MenuName の行に限る）から動的に行う — [kernel/main.ps1:420-500](kernel/main.ps1)、[kernel/KERNEL_API.md:170, 432-440](kernel/KERNEL_API.md)
- FlexProfile: `Get-FlexModuleStateMap`（既定は Pending。直近の実行結果と履歴で上書き）、行ごとの `[Run]`（AutoConfirmMode）と `[Log]`、`[Run Selected (N)]`、`[Run: <Group>]`、右クリックの「Mark as Pending」、`[Restart Now]`、`[Complete]`（完了処理）、PENDING FINALIZE バッジ。「実行は常に AutoPilot 動作 / 完了は常に手動」 — [apps/fabriq_operator/lib/flex_dashboard.ps1:13-40, 488, 516-542](apps/fabriq_operator/lib/flex_dashboard.ps1)、[README.md:206-227](README.md)
- Flex の再開は resume_state v2（`SelectedOrders` / `ModuleStates`）で、再起動後にダッシュボードへ戻る — [kernel/main.ps1:1851-1970](kernel/main.ps1)、[kernel/KERNEL_API.md:296-299](kernel/KERNEL_API.md)
- async: `async_config.json` は `DefaultAsync=true`（3.3.0 以降は全モジュールが監視付き Runspace で動く）。Skip と timeout は Error として記録する（§1）— [kernel/json/async_config.json](kernel/json/async_config.json)
- プロファイルの実例: `Master_Config01.csv`（ライセンス → 壁紙 → 時刻 → FW → 電源 → ユーザー → レジストリ → アプリ → ODT → DPI → 解像度 → コピー → プリンタの 17 行）、`sysprep.csv`（`ErrorMode` 列あり。`__RESTART__` を Enabled=0 で置いている）、`_test_harness_async.csv`（`__AUTOPILOT__` WaitSec=1、`__ASYNC__`、skip/retry/hang のシミュレーション）— [profiles/Master_Config01.csv](profiles/Master_Config01.csv)、[profiles/sysprep.csv](profiles/sysprep.csv)、[profiles/_test_harness_async.csv](profiles/_test_harness_async.csv)
- プロファイル別データオーバーレイ（3.7.0）: `profiles/<name>/modules/<module>/<rel>` が本体の `modules/<tier>/<module>/<rel>` より優先される（Profile-First）。本体へのフォールバックは毎回 `Show-Warning` を出し、テレメトリの `resolvedFrom` と HTML の `Data Set` に記録する。資材フォルダはフォルダ単位で全か無か。書き込みは `-ForWrite` で PDF 側へ — [kernel/KERNEL_API.md:33](kernel/KERNEL_API.md)、[dev/PROFILE_DATA_OVERLAY_PLAN.md:1-30](dev/PROFILE_DATA_OVERLAY_PLAN.md)、[profiles/_test_overlay/modules/test_harness_config/test_harness_list.csv](profiles/_test_overlay/modules/test_harness_config/test_harness_list.csv)

### Inferences
- **(A)** プロファイル CSV（Order/ScriptPath/Enabled/Description/Segment/ErrorMode/Group）と `__GATE__` の意味論（失敗は防ぐが、省略は防がない）。fabriq_mac の `order,module,enabled,description` は、これを縮めたもの。Android 版では `Segment`・`ErrorMode`・`Group` を残すと表現力が上がる。`__GATE__` は**端末ごとに評価**する（ある端末で precheck が NG なら、その端末だけが止まる）。
- **(A)** プロファイル別データオーバーレイ（案件・ロットごとに設定データをプロファイル側のフォルダへ集める）。複数顧客の端末ロットを 1 つのホストツールで扱う Android 版にはむしろ必須に近い。
- **(B)** FlexProfile の「端末 1 台 × モジュール M 行」の状態マップは、Android 版では **「端末 N 台 × ステップ M 列」のマトリクス**（電子チェックシート）に広がる。状態はシリアルごとに保存する（fabriq_mac の決定 2 と一致）。`[Run Selected]` は「選んだ端末 × 選んだステップ」、`Mark as Pending` はセル単位のリセットになる。
- **(B)** ErrorMode 空欄時のモーダルダイアログは 1 対多では致命的（N 台が並列に動く中で 1 台のために全体が止まる）。**ダイアログを使わない端末ごとの「要対応」状態**（セルを黄色にする等）に置き換えるべき。fabriq でも AutoPilot が止まる問題として TM t-0055 で指摘されている（§12）。
- **(B)** `__RESTART__` は、ホストモデルでは「`adb reboot` → 再接続待ち → 起動完了待ち → 続行」という同期的なステップに単純化できる（RunOnce・DPAPI・AutoLogon が要らない）。fabriq_mac が「再起動からの再開はしない（OS 更新だけ最後に回す）」と決めたのに対し、Android のホストモデルでは**再起動を跨いでも安く続行できる**（要検証）。
- **(C)** `__REEXPLORER__`、`__AUTO_to_<User>__`（Windows AutoLogon）は Windows 専用。`__ASYNC__` はホストモデルでは「端末ごとのワーカ = 常に非同期」になるので、マーカーとしては不要。

### Gaps
- `Invoke-FlexProfileLoop`（L831-1094）の Action ごとの詳細（RunSingle/RunGroup/RunBatch/Complete/RestartNow/ResetState/Close）は分岐名の確認までで、中身は読んでいない。
- `__AUTOPILOT__` マーカーの戻り値が Linear の GUI 経路で本当に使われていないかは、grep の範囲でしか確認していない（Flex 側は独自の WaitSec 入力欄を使う: [apps/fabriq_operator/lib/flex_dashboard.ps1:155, 289, 767](apps/fabriq_operator/lib/flex_dashboard.ps1)）。

---

## 5. kernel/csv（hostlist / workers / categories / log_destinations / manifesto）とホスト自動選択

### Takeaway
対象機ごとの値は `hostlist.csv` の 1 行で、入室時に `SELECTED_*` 環境変数へ展開する（`ENC:` は透過的に復号、`__SELF__` は実機の値に置き換え）。ホストの自動選択は「`NewPCName` == 現在の `COMPUTERNAME`」→「`__SELF__` 行」の 2 段階しかない。シリアル番号では照合しないので、Android 版では fabriq_mac 方式（シリアル → 名前 → 前回 → 手動）に置き換えるのが妥当。

### Cited Findings
- `hostlist.csv` のヘッダ: `AdminID,OldPCName,NewPCName,EthernetIP,EthernetSubnet,EthernetGateway,WifiIP,WifiSubnet,WifiGateway,DNS1..DNS4,Pin,Printer{1..10}{Name,Driver,Port}`（計 44 列）— [kernel/csv/hostlist.csv](kernel/csv/hostlist.csv)
- 環境変数への対応: `AdminID→SELECTED_KANRI_NO`、`OldPCName→SELECTED_OLD_PCNAME`、`NewPCName→SELECTED_NEW_PCNAME`、`Ethernet*→SELECTED_ETH_*`、`Wifi*→SELECTED_WIFI_*`、`DNS1-4`、`Pin`（秘密値として登録）、`Printer<i>*→SELECTED_PRINTER_<i>_*`。セルの解決規則は `__SELF__` → 実機の値（`Get-CurrentPCInfo`）、`ENC:` → 復号、それ以外はそのまま — [kernel/main.ps1:86-211](kernel/main.ps1)
- ホスト自動選択: Pass 1 は `$h.NewPCName -eq $CurrentPCName`、Pass 2 は `NewPCName` か `OldPCName` が `__SELF__` の行。一致すれば「* Auto-detected: … (matches current PC)」を表示し、検索で絞り込みもできる — [apps/fabriq_operator/lib/session_form.ps1:162-252](apps/fabriq_operator/lib/session_form.ps1)
- `workers.csv`: `ID,Name`（例 W001）→ `FABRIQ_WORKER_NAME`、履歴の Worker 列 — [kernel/csv/workers.csv](kernel/csv/workers.csv)、[kernel/main.ps1:1612-1660](kernel/main.ps1)
- `categories.csv`: `Category,Order`（Network 10, Display 20, Desktop 30 …）でメニューのカテゴリ順を決める — [kernel/csv/categories.csv](kernel/csv/categories.csv)、[kernel/common.ps1:4965-4985](kernel/common.ps1)
- `log_destinations.csv`: `Path,Type(UNC/Local),Enabled,AuthUser,AuthPass,Description`。Enabled な宛先があれば、完了処理で `log_uploader` を自動実行する — [kernel/csv/log_destinations.csv](kernel/csv/log_destinations.csv)、[kernel/common.ps1:3550-3600](kernel/common.ps1)
- `manifesto.csv`: シュルキティニスム宣言の文章（世界観の演出で、機能は無い）— [kernel/csv/manifesto.csv](kernel/csv/manifesto.csv)
- これらの現場固有 CSV（hostlist/workers/log_destinations）は、更新オーバーレイで**上書きしない**対象 — [kernel/KERNEL_API.md:465-488](kernel/KERNEL_API.md)
- 未着手の要望: hostlist とプロファイルに「ホストマーカ」を持たせ、Segment と連動させて「端末グループごとに一部のモジュールだけ実行する」— [.tm/tasks.json](.tm/tasks.json)（t-0062）

### Inferences
- **(A)** 「1 行 = 1 台」の hostlist、`ENC:` の透過復号、秘密値のマスク、作業者 CSV、カテゴリ順 CSV、ログ宛先 CSV。これらの構造はそのまま使える。
- **(B)** **照合キーをシリアルに変える**。Android の hostlist は `Serial`（必須・照合キー）、`AssetTag`（AdminID 相当）、`DeviceName`、`WifiProfile`、`Segment/Group`、端末ごとのアプリ指定、などに作り直す（fabriq_mac の `HOST_<column>` 方式と合わせる）。ホストモデルでは「接続された N 台を `adb devices` で列挙 → シリアルで hostlist を引く → 一致しない端末は手動で割り当てるか `__SELF__` 相当で受け入れる」という**ロット受付画面**になる（`adb devices` はホスト側での推論。要検証）。
- **(B)** ホストモデルでは、作業者・パスフレーズ・ホスト選択のセッション画面が「1 セッション = 1 ロット（N 台）」になる。パスフレーズは 1 回入力すれば全端末に効く。
- **(C)** 列の大半（NIC/DNS/プリンタ 30 列）は Windows の PC 前提。Android のプリンタは印刷サービスのプラグイン次第で、hostlist で列として持つ意味は薄い（推論）。

### Gaps
- `Get-CurrentPCInfo`（common.ps1:4719）の全取得項目は読んでいない。

---

## 6. エビデンス（manifest・スクリーンショット・HTML チェックリスト・実行履歴 CSV・アップロード・テレメトリ）

### Takeaway
エビデンスは 4 層で残る。①モジュールごとの自動スクリーンショット、②`evidence_config` が収集する 33 セクションのインベントリと `manifest.json`（外部ツール向けの公開契約）、③追記専用の列を持つ実行履歴 CSV、④HTML チェックリスト（総合判定は OK/NG/Incomplete/CA、hostlist の期待値と実機の値の照合を含む）。これらを `log_uploader` で共有フォルダへ送る。このうち「期待値と実測の照合表」と「manifest の前方互換ルール」は、Android 版の納品物の設計にそのまま使える。

### Cited Findings
- エビデンスの保存先: `.\evidence\{ts}_{NewPCName}_{HardwareUniqueId}_evidence\evidence\`（`FABRIQ_EVIDENCE_BASE`）— [kernel/common.ps1:1722-1777](kernel/common.ps1)
- 自動撮影: 毎モジュールの直後に `Capture-ScreenEvidence -ModuleName -Status` を呼び、`auto_capture/{ts}_{module}_{status}_{pc}.png` として保存する。`SetProcessDPIAware` を呼ぶのでプロセスの DPI モードが元に戻らなくなる。矩形は `GetSystemMetrics` から取る — [kernel/main.ps1:704](kernel/main.ps1)、[kernel/common.ps1:5311-5422](kernel/common.ps1)、[kernel/KERNEL_API.md:102](kernel/KERNEL_API.md)
- 手動撮影（「魚拓」=Gyotaq ボタン）: `Save-Screenshot` が `gyotaku/` に保存する — [kernel/common.ps1:5424](kernel/common.ps1)、[kernel/KERNEL_API.md:209](kernel/KERNEL_API.md)
- 実行履歴 CSV の列（追記専用）: `Timestamp,KanriNo,PCName,ModuleName,Category,Status,Message,WindowsUser,Worker,MediaSerial,SessionID,Verified,Order,Profile`。古いヘッダが今の列の先頭部分と一致する場合は位置で読む（後方互換）— [kernel/common.ps1:2546-2586](kernel/common.ps1)
- HTML チェックリスト: Meta 欄（Target PC New/Old、AdminID、Worker、Media Serial、Hardware ID、Data Set、Generated At、Elapsed Time）、モジュール表（#/Module/Category/Result/Verified/Time/Message）、「PC Settings Verification」（項目 / hostlist の期待値 / OS の実際の値 / Match）、プリンタ照合（Match/Not Found/Extra）。総合判定は、Error か検証 NG があれば **NG**、未実行があれば **Incomplete**、余分なプリンタがあれば **CA**、それ以外は **OK** — [kernel/common.ps1:3230-3440](kernel/common.ps1)
- 完了処理 `Complete-ProfileExecution`: `Export-ExecutionHistory` → `Export-HtmlChecklist` → `log_uploader`（Auto のときは警告だけ、Manual のときは結果を記録）→ ビューアの順 — [kernel/common.ps1:3509-3640](kernel/common.ps1)
- Evidence Manifest 契約（schemaVersion 1、kernel 2.2.2〜）: 置き場所は `{evidenceBase}/pc_information/{date}_{uid}_{pc}/manifest.json`。再実行時は `.bak` に 1 世代だけ残す。パスは manifest からの相対。必須フィールドは `schemaVersion, manifestType="fabriq-evidence-manifest", evidenceConfigVersion, fabriqKernelVersion, collectedAt, computerName/hardwareUniqueId/selectedNewPcName, sections[{id,title,files,status,reason,elapsedMs}], summary{…}`。status は Success/Skipped/Failed/Partial。前方互換ルールは、未知の major 版なら manifest を使わない従来モード、未知のセクションは raw 表示、未知の status は Failed 扱い — [kernel/EVIDENCE_MANIFEST.md](kernel/EVIDENCE_MANIFEST.md)、[kernel/KERNEL_API.md:534-580](kernel/KERNEL_API.md)
- `evidence_config` の 33 セクション: 01 System Basic Info / 02 Local Users / 03 Local Groups / 04 Group Members / 05 Domain・AzureAD + Profiles / 06 Network / 07 Printers / 08 BitLocker / 8b Disk & Partition / 09 MAC / 10 Serial / 11 Installed Software / 12 Firewall / 13 Optional Features / 14 Server Roles / 15 Power / 16 WiFi Profiles / 17 Restore Points / 18 Defender / 19 WU History / 20 TEMP Log Backup / 21 Windows License / 22 Office License / 23 Security Baseline / 24 GPO Report / 25 Certificates / 26 Battery / 27 Env Vars / 28 Startup / 29 Memory / 30 PnP / 31 HW IDs / 32 Credential Manager / 33 Outlook Accounts（モジュールは 3,250 行）— [modules/standard/evidence_config/evidence_list.csv](modules/standard/evidence_config/evidence_list.csv)
- テレメトリ（内部・AI 開発用。公開契約ではない）: `logs/telemetry/{SessionID}/_meta.json`、`_kernel.jsonl`（profile.start/end、restart.invoked、resume.consumed、finalize.*）、`modules/{seq}_{module}.jsonl`（envelope.start/end、csv.load、show.*、error、cmdlet.verbose）。log_uploader の送信対象から外す — [dev/TELEMETRY_INTERNAL.md:9-93](dev/TELEMETRY_INTERNAL.md)
- 外部ツールとの関係: Evidence Manager（WPF/.NET 8、別リポジトリ）は manifest・`auto_capture/`・`checklist/`・`export_history/` を読むだけで書かない。出力は Excel の一覧表と個別シート — [README.md:391-404](README.md)

### Inferences
- **(A)** 実行履歴 CSV の考え方（追記専用の列・古いファイルを読める後方互換・SessionID・Verified・Order・Profile）、HTML チェックリスト（期待値 vs 実測の照合表と、OK/NG/Incomplete の総合判定）、manifest の前方互換ルール（未知の版・セクション・status の安全側処理）。どれもそのまま移せる。Android 版では「1 台 1 枚」と「ロット一覧表」の 2 段にする。
- **(B)** スクリーンショットは**ホストモデルの方がむしろ簡単になる**。`adb exec-out screencap -p` で端末の画面をホストに直接取れるので（要検証）、fabriq の DPI 由来の見切れ問題（t-0085 / t-0101）とは無縁になる。各ステップ後の自動撮影は、端末ごとのフォルダに保存する。
- **(B)** `evidence_config` の 33 セクションは、Android では `getprop`（機種・OS 版・セキュリティパッチ・シリアル）、`pm list packages`（＋バージョン）、`settings list system|secure|global`、`dumpsys battery|wifi|device_policy`、ストレージ、IMEI/MAC（Android 10 以降は取得制限あり）といったインベントリに置き換わる（推論・要検証）。
- **(A)** テレメトリ（JSONL・セッション/モジュール単位・秘密値マスク）は、端末シリアル単位のサブフォルダにすれば移植できる。ArtPulse（Show-* のたびにカウンタファイルを +1）は「1 プロセス 1 台」の生存シグナルなので、ホストモデルでは**端末ワーカごとのハートビートと ADB 接続状態**に作り直す (B)。
- **(C)** BitLocker・GPO・Defender・Office・Credential Manager などのセクションは Windows 専用。

### Gaps
- `Export-HtmlChecklist`（L2907-3508・約 600 行）のテンプレート全体と、`log_uploader` の robocopy 処理の詳細は読んでいない。
- Evidence Manager のリポジトリは別にあり、ここには無いので未確認。

---

## 7. ガバナンス／開発プロセス（SemVer・CHANGELOG・テスト・テストリグ・エンコーディング・設計ゲート・CI・TM）

### Takeaway
fabriq では、コードと同じくらい**開発プロセスが資産**になっている。中身は、カーネルとモジュールを別々に SemVer 管理すること、Keep a Changelog、Pester のテスト 671 件（`It` ブロックの数）、VM と WinRM を使ったモジュールテストリグ（記述子 `test.psd1` で駆動）、CI（PowerShell 5.1 で 3 つのチェックを exit code で判定）、実装前の設計ゲート（Mermaid の状態図と自己敵対レビュー）、TM のタスク台帳。Android 版でも言語を替えて再現する価値がある。

### Cited Findings
- 版管理: `kernel/KERNEL_VERSION`（真のソース）と `KERNEL_API.md` を同じコミットで更新する。モジュールごとに `VERSION`（baseline 1.0.0）と `REQUIRES_KERNEL`（baseline 2.0.0）を持つ。MAJOR/MINOR/PATCH の判定表があり、迷ったら大きい側に倒す。コードを変えたら CHANGELOG の `[Unreleased]` にコンポーネント名付きで追記する。リリース時は 6 箇所の版表記を同期させ、`check_version.ps1` で確認する。`git tag` はユーザーに依頼する — [CLAUDE.md](CLAUDE.md)（バージョン管理ルール §A〜§J）
- CHANGELOG: 6,265 行、版見出し 31 個（`[Unreleased]`、`[3.6.2] - 2026-07-02` … `[2.2.0] - 2026-04-22`）— [CHANGELOG.md](CHANGELOG.md)
- 設計ゲート: `kernel/` や `modules/` のロジックを変える前に設計を提示して承認を得る。フル版は概要・Mermaid `stateDiagram-v2`・敵対検証（未到達状態・抜けた失敗遷移・冪等性違反・部分失敗・`__RESTART__` 跨ぎ・判断を operator に丸投げしていないか）・変更スコープ宣言。軽量版は 3 行 — [CLAUDE.md](CLAUDE.md)（§4, §C, §D）
- コード規約: `.ps1` は英語のみ。日本語を含む `.ps1` と CSV は UTF-8 BOM 必須（PS 5.1 + 日本語版 Windows で CP932 として誤読され、`-match` が壊れる実害がある）。CSV は BOM + CRLF。`check_ps1_encoding.ps1` で検出する — [CLAUDE.md](CLAUDE.md)（§7）
- 破壊的操作のガード: 再帰削除の前に対象パスを検証する（`Join-Path` が失敗して意図しないパスを消した事故が過去にある）— [CLAUDE.md](CLAUDE.md)（§8）、[kernel/KERNEL_API.md:63-69](kernel/KERNEL_API.md)
- テスト実行: `dev/run_tests.ps1`（Pester v5 以上。`tests/` と `apps/fabriq_ios/tests/` の `*.tests.ps1` を探索する）。正式なエンジンは Windows PowerShell 5.1 で、合否は exit code で判定する — [dev/run_tests.ps1:1-20](dev/run_tests.ps1)
- テストの内訳: `tests/` に 31 ファイル・`It` 586 個（kernel 21 ファイル・354、modules 6 ファイル・185、operator 2 ファイル・29、dev 1 ファイル・17、IosSmoke 1）。`apps/fabriq_ios/tests/` に 10 ファイル・85 — [tests/](tests/)、[apps/fabriq_ios/tests/](apps/fabriq_ios/tests/)。kernel テストの例: ResumeState 29、ProfileDataOverlay 45、HistoryChecklistChain 31、PathGuards 27、Invoke-BatchExecution 26、Set-SelectedHostEnvironment 26、Resolve-ProfileModules 21、GateBarrier 15、Import-ModuleCsv 16 — [tests/kernel/](tests/kernel/)
- CI: `.github/workflows/ci.yml` が push/PR で Windows PowerShell 5.1 上の `run_tests.ps1` / `check_version.ps1` / `check_ps1_encoding.ps1` を別々のステップとして実行する（`.md` と `.tm` だけの変更では起動しない）— [.github/workflows/ci.yml](.github/workflows/ci.yml)
- テストリグ: `dev/test_rig/run_module_tests.ps1` が各モジュールの `test.psd1` を読み、隔離 VM に WinRM で同期 → 振り分け → fixture → envelope → `Invoke-SafeCommand` → 独立 oracle → 判定 → 後片付け、の順に実行する。interactive なものや winrmSafe=false のものは SKIP。VM の自動復元はできない（vmrun に届かない）ので、`cleanup=snapshot` は「手動での復元が必要」と報告する。AI は実行経路に入れず、合否は決定的な層で決める（C6）— [dev/test_rig/run_module_tests.ps1:1-25](dev/test_rig/run_module_tests.ps1)
- `test.psd1` は 51 モジュールにある（category A 28 / A-interactive 5 / B 10 / C 2 / D 4 / E 1 / none 1）— `modules/*/*/test.psd1`（grep で集計）
- 関連エピック: TM t-0063（VM 検証機を使った AI 駆動のテスト自動化。PoC は reg_hklm_config で 6/6 PASS）、t-0064（79 モジュールを A〜E に分類）— [.tm/tasks.json](.tm/tasks.json)
- タスク台帳: `.tm/tasks.json` は 110 件（完了 48 / レビュー待ち 31 / 未着手 20 / 保留 7 / 対応中 4）。人と Claude が同じファイルを編集する — [.tm/tasks.json](.tm/tasks.json)、[CLAUDE.md](CLAUDE.md)（タスク管理）
- そのほかの dev ツール: `check_version.ps1`、`check_ps1_encoding.ps1`、`seed_module_versions.ps1`（冪等）、`build_framework_patch.ps1` / `build_preset_patch.ps1`（差分パッチの配布物）、`verify_comments_only.ps1`、`test_execution_toolbar.ps1`、`launcher/`（C# のランチャ `Fabriq.exe`。`requireAdministrator` マニフェスト付き。csc.exe だけでビルドできる）— [dev/](dev/)、[dev/launcher/README.md](dev/launcher/README.md)

### Inferences
- **(A)** 開発プロセス一式（カーネルとモジュールの独立 SemVer、`REQUIRES_KERNEL`、CHANGELOG の規約、設計ゲート、状態図、敵対レビュー、「無返却は Error」等の fail-closed 原則、エンコーディング検査、CI で exit code 判定、TM 台帳）は、言語にも OS にも依存しない。Android 版の初日から持ち込める。
- **(B)** テストリグの考え方は、Android では **AVD エミュレータや検証用の実機に対する `adb` 経由の独立 oracle**（`settings get` / `pm list` / `dumpsys`）になり、WinRM 版より作りやすい可能性が高い。ホストツールが最初から ADB で操作するので、テスト経路と本番経路が同じになる（推論・要検証）。
- **(C)** 「PS 5.1 が正式なエンジン」「CP932 の BOM 問題」は PowerShell 固有の規則。ただし Android 版のホスト GUI も Windows（PowerShell/WinForms）で作るなら、そのまま当てはまる。

### Gaps
- `check_version.ps1` / `check_ps1_encoding.ps1` / `build_framework_patch.ps1` の中身は見出しの確認だけで、読んでいない。

---

## 8. モジュールカタログ（全 79）と Android 対応概念

### Takeaway
モジュールのディレクトリは 79 個（`standard` 63、`extended` 16）、入口となる `.ps1` は 104 本・計 36,548 行。そのうち Android にそれらしい対応があるのは約 30 で、APK の導入・削除、設定の書き込み、Wi-Fi、画面・音量・明るさ、時刻、端末名、ファイルの push/削除、再起動、インベントリ、手作業ガイドなど。残りの約 45 は Windows 専用（AD・GPO・BitLocker・レジストリ・Office・プリンタドライバ・FW 等）。README は「Extended 18・合計 81」と書いているが、実際のディレクトリと README の一覧はどちらも extended 16。

### Cited Findings
（目的・メニュー名は各モジュールの `module.csv` とスクリプト冒頭のコメント（`[PURPOSE]` 等）から。表記は `name (VERSION / REQUIRES_KERNEL) : 目的`。出典はすべて `modules/<tier>/<name>/`）

**standard（63）**
- `acl_config` (1.1.0/3.7.0): icacls によるディレクトリ ACL のバックアップ・復元 — [modules/standard/acl_config/](modules/standard/acl_config/)
- `app_config` (1.1.0/3.7.0): ローカルの exe/msi をサイレントインストール（`app_list.csv`: AppName,FileName,Type,SilentArgs）— [modules/standard/app_config/app_list.csv](modules/standard/app_config/app_list.csv)
- `autologon_config` (1.2.0/3.7.0): 1 回限りの Windows AutoLogon をレジストリに設定（`__AUTO_to_<User>__` の実体）— [modules/standard/autologon_config/](modules/standard/autologon_config/)
- `bitlocker_config` (1.2.0/2.0.0): BitLocker の有効化・無効化・完了待ち（3 スクリプト）— [modules/standard/bitlocker_config/](modules/standard/bitlocker_config/)
- `bloatware_remove` (1.0.0/2.0.0): レジストリの UninstallString を使って、旧来型デスクトップアプリを MatchPattern で削除 — [modules/standard/bloatware_remove/](modules/standard/bloatware_remove/)
- `brightness_config` (1.0.0/2.0.0): WMI で画面の明るさを設定 — [modules/standard/brightness_config/](modules/standard/brightness_config/)
- `browser_addon_config` (1.0.1/2.0.0): Chrome/Edge 拡張を ExtensionInstallForcelist（ポリシーのレジストリ）で強制導入 — [modules/standard/browser_addon_config/](modules/standard/browser_addon_config/)
- `cert_config` (1.1.1/3.7.0): .pfx/.p12/.cer/.crt を証明書ストアへインポート（`%SELECTED_NEW_PCNAME%_*.p12` の自動振り分けあり）— [modules/standard/cert_config/cert_list.csv](modules/standard/cert_config/cert_list.csv)
- `copyfile_config` (1.1.0/3.7.0): source/ のファイル・フォルダを配布 — [modules/standard/copyfile_config/](modules/standard/copyfile_config/)
- `credential_config` (0.1.0/2.0.0): 現在ユーザーの Windows 資格情報マネージャーへ登録 — [modules/standard/credential_config/](modules/standard/credential_config/)
- `default_app_config` (1.2.0/3.7.0): 既定アプリの関連付けを XML で export / import — [modules/standard/default_app_config/](modules/standard/default_app_config/)
- `domain_join` (2.2.0/2.0.0): AD ドメイン参加（`domain.csv`。OU 指定は t-0103）— [modules/standard/domain_join/](modules/standard/domain_join/)
- `dpi_api_config` (1.1.0/2.0.0): 非公開 API で DPI スケーリングを即時変更 — [modules/standard/dpi_api_config/](modules/standard/dpi_api_config/)
- `driver_config` (1.3.0/3.7.0): サードパーティドライバを機種ごとに export / pnputil で import — [modules/standard/driver_config/](modules/standard/driver_config/)
- `evidence_config` (1.8.1/2.0.0): 33 セクションのエビデンス収集と manifest.json（3,250 行）— [modules/standard/evidence_config/](modules/standard/evidence_config/)
- `fabriq_app_launcher` (1.0.1/2.0.0): `target_apps.csv` に従って `apps/` の GUI ツールを起動 — [modules/standard/fabriq_app_launcher/](modules/standard/fabriq_app_launcher/)
- `file_delete` (1.1.1/3.5.0): CSV に沿ってファイル・フォルダを削除（パスガード付き）— [modules/standard/file_delete/](modules/standard/file_delete/)
- `firewall_config` (1.1.0/3.7.0): Windows FW のプロファイル（Domain/Private/Public）ごとの ON/OFF — [modules/standard/firewall_config/](modules/standard/firewall_config/)
- `firewall_rule_config` (1.2.0/3.7.0): FW ポリシー（.wfw）の snapshot export / import — [modules/standard/firewall_rule_config/](modules/standard/firewall_rule_config/)
- `firewall_rule_make_config` (1.1.1/2.0.0): CSV から FW ルールを作成（冪等）— [modules/standard/firewall_rule_make_config/](modules/standard/firewall_rule_make_config/)
- `generic_batch_runner` (1.1.0/3.7.0): `batch_list.csv` の .bat を実行 — [modules/standard/generic_batch_runner/](modules/standard/generic_batch_runner/)
- `generic_process_runner` (1.2.0/3.7.0): CSV に書いた任意の EXE を実行 — [modules/standard/generic_process_runner/](modules/standard/generic_process_runner/)
- `gpo_config` (0.1.0/3.5.0): Registry.pol を直接編集してローカル GPO（管理用テンプレート）を適用・バックアップ — [modules/standard/gpo_config/](modules/standard/gpo_config/)
- `hostlist_precheck` (0.1.0/3.5.0): 読み取り専用のゲート。hostlist の行と実機の保存済み設定（ホスト名・静的 IP・GW・DNS）を照合し、1 つでも不一致なら Error（後ろに `__GATE__` を置く前提）— [modules/standard/hostlist_precheck/](modules/standard/hostlist_precheck/)
- `hostname_config` (1.0.0/2.0.0): `SELECTED_NEW_PCNAME` に名前を変更し、保留中の名前を読み返して検証 — [modules/standard/hostname_config/](modules/standard/hostname_config/)
- `ipaddress_config` (1.1.1/2.0.0): hostlist に従って Ethernet / Wi-Fi の静的 IP を設定 — [modules/standard/ipaddress_config/](modules/standard/ipaddress_config/)
- `local_user_config` (1.4.0/3.7.0): ローカルユーザーの作成・削除（PC ごとの `local_user_host_list.csv` あり）— [modules/standard/local_user_config/](modules/standard/local_user_config/)
- `odt_config` (1.2.0/3.7.0): Office Deployment Tool で Office をインストール — [modules/standard/odt_config/](modules/standard/odt_config/)
- `office_license_config` (1.2.0/3.7.0): OSPP.vbs で Office のキー投入と認証 — [modules/standard/office_license_config/](modules/standard/office_license_config/)
- `office_update` (1.1.1/2.0.0): Click-to-Run の更新を起動し、完了を待つ — [modules/standard/office_update/](modules/standard/office_update/)
- `partition_config` (1.1.2/2.0.0): パーティションの縮小・作成・フォーマット — [modules/standard/partition_config/](modules/standard/partition_config/)
- `power_config` (1.1.1/2.0.0): 電源プラン、画面オフ・スリープ・休止・ボタン・蓋の動作など（冪等 skip の手本、1,164 行）— [modules/standard/power_config/power_list.csv](modules/standard/power_config/power_list.csv)
- `ppkg_config` (1.2.0/3.7.0): プロビジョニングパッケージ（.ppkg）の導入・削除 — [modules/standard/ppkg_config/](modules/standard/ppkg_config/)
- `printer_delete` (1.1.0/3.7.0): プリンタ削除（3 つのモード）— [modules/standard/printer_delete/](modules/standard/printer_delete/)
- `printer_driver_config` (1.3.0/3.7.0): プリンタドライバの導入・削除、TCP/IP ポートとプリンタの登録（hostlist の `SELECTED_PRINTER_*` 連動）— [modules/standard/printer_driver_config/](modules/standard/printer_driver_config/)
- `process_killer` (1.0.0/2.0.0): 指定プロセスを強制終了（冪等）— [modules/standard/process_killer/](modules/standard/process_killer/)
- `profile_delete` (1.2.0/3.5.0): Win32_UserProfile でユーザープロファイルを削除 — [modules/standard/profile_delete/](modules/standard/profile_delete/)
- `reg_hkcu_config` (1.2.0/3.7.0): HKCU と Default プロファイルのレジストリ設定・削除 — [modules/standard/reg_hkcu_config/](modules/standard/reg_hkcu_config/)
- `reg_hklm_config` (1.2.0/3.7.0): HKLM のレジストリ設定・削除（`Test-RegistryValueMatch` で検証）— [modules/standard/reg_hklm_config/](modules/standard/reg_hklm_config/)
- `resolution_api_config` (1.1.0/2.0.0): 画面解像度を即時変更 — [modules/standard/resolution_api_config/](modules/standard/resolution_api_config/)
- `restart_config` (1.0.0/2.0.0): 自動起動を登録して再起動 — [modules/standard/restart_config/](modules/standard/restart_config/)
- `restore_point` (1.2.1/2.0.0): システムの復元を設定し、復元ポイントを作成 — [modules/standard/restore_point/](modules/standard/restore_point/)
- `robocopy_config` (1.1.0/3.5.0): CSV のフラグ列から robocopy のジョブを組んで実行 — [modules/standard/robocopy_config/](modules/standard/robocopy_config/)
- `scheduled_task_config` (1.1.0/2.0.0): スケジュールタスクの有効化・無効化（冪等）— [modules/standard/scheduled_task_config/](modules/standard/scheduled_task_config/)
- `signout_config` (1.0.0/2.0.0): サインアウト（fabriq も終了する）— [modules/standard/signout_config/](modules/standard/signout_config/)
- `spi_config` (1.1.1/2.0.0): SystemParametersInfo による UI 設定 — [modules/standard/spi_config/](modules/standard/spi_config/)
- `ssid_config` (1.0.1/2.0.0): `netsh wlan add profile` で Wi-Fi プロファイルを登録（SSID,Authentication,Encryption,Password,AutoConnect,NonBroadcast）— [modules/standard/ssid_config/ssid_list.csv](modules/standard/ssid_config/ssid_list.csv)
- `startlayout_config` (1.1.0/3.7.0): スタートメニューのレイアウトの backup / build（WCD）/ import（ppkg）/ delete — [modules/standard/startlayout_config/](modules/standard/startlayout_config/)
- `startup_command_config` (1.1.0/2.0.0): 1 回だけ実行するスクリプトを ProgramData に置き、次回ログオン時に実行 — [modules/standard/startup_command_config/](modules/standard/startup_command_config/)
- `storeapp_config` (1.0.1/2.0.0): Store アプリ（AppX）を削除（`storeapp_list.csv`: No,AppName,…）— [modules/standard/storeapp_config/storeapp_list.csv](modules/standard/storeapp_config/storeapp_list.csv)
- `sysprep_config` (1.3.0/3.7.0): unattend.xml と SetupComplete.cmd を生成して sysprep を実行 — [modules/standard/sysprep_config/](modules/standard/sysprep_config/)
- `system_finalize` (1.0.1/2.0.0): shell32 の再登録、アイコンとサムネイルのキャッシュ削除、Explorer 再起動 — [modules/standard/system_finalize/](modules/standard/system_finalize/)
- `taskbar_config` (1.3.0/3.7.0): タスクバーのピン留め（LayoutModification.xml を生成）— [modules/standard/taskbar_config/](modules/standard/taskbar_config/)
- `temp_ipaddress_config` (1.0.2/3.7.0): GUI で IP の候補を選んで一時的に割り当て — [modules/standard/temp_ipaddress_config/](modules/standard/temp_ipaddress_config/)
- `test_error_module` (1.0.0/2.0.0): 必ず Error を返す（エラー時の流れの試験用）— [modules/standard/test_error_module/](modules/standard/test_error_module/)
- `test_harness_config` (1.0.1/2.0.0): データ駆動でカーネル機能を試すシミュレータ（Segment でシナリオを切り替え）— [modules/standard/test_harness_config/](modules/standard/test_harness_config/)
- `time_sync_config` (1.1.0/2.0.0): W32Time に NTP サーバーを設定して同期 — [modules/standard/time_sync_config/](modules/standard/time_sync_config/)
- `volume_config` (1.0.1/2.0.0): Core Audio で主音量と Mute を設定（冪等・検証の手本）— [modules/standard/volume_config/](modules/standard/volume_config/)
- `wallpaper_config` (1.3.0/3.7.0): 壁紙を設定（リムーバブルメディアでも安全な方式）— [modules/standard/wallpaper_config/](modules/standard/wallpaper_config/)
- `windows_feature_config` (0.1.1/2.0.0): DISM でオプション機能を有効化・無効化（冪等・検証の手本）— [modules/standard/windows_feature_config/](modules/standard/windows_feature_config/)
- `windows_license_config` (1.3.0/3.7.0): Windows のキー投入（CSV / OEM ファームウェアのキー）と認証 — [modules/standard/windows_license_config/](modules/standard/windows_license_config/)
- `windows_update` (1.1.2/2.0.0): WU の 1 パス実行。`module.csv` は持たず GUI ボタン専用で、専用の再起動ループ（`wu_state.json` と AutoLogon）で回す — [modules/standard/windows_update/](modules/standard/windows_update/)、[kernel/main.ps1:1156-1340](kernel/main.ps1)
- `winget_install` (1.1.1/2.0.0): winget 本体の更新、アプリの導入・upgrade（`app_list.csv`: AppID,Options）— [modules/standard/winget_install/app_list.csv](modules/standard/winget_install/app_list.csv)

**extended（16）**
- `azure_ad_join_check` (1.0.0/2.0.0): dsregcmd で Azure AD 参加を確認し、未参加なら Error（Script Looper で再試行させる）
- `builtin_admin_config` (1.1.0/2.0.0): 組み込み Administrator の設定
- `desktop_icon_config` (1.1.0/3.7.0): デスクトップアイコン配置の .reg バックアップ・復元
- `directory_cleaner` (1.1.0/2.0.0): 指定ディレクトリを再帰削除（`Test-ForbiddenPath` の元になった実装）
- `display_config` (1.1.0/2.0.0): レジストリ経由で解像度を変更
- `dpi_config` (1.1.0/2.0.0): レジストリ（HKCU と Default）でモニタごとの DPI を設定
- `group_config` (1.2.2/2.0.0): ローカルグループのメンバー追加・削除
- `history_destroyer` (1.4.0/3.7.0): Windows の履歴・キャッシュ・一時データ・Wi-Fi・AutoLogon 情報を削除
- `ipv6_config` (1.1.0/2.0.0): IPv6 の設定
- `log_uploader` (1.1.1/2.0.0): logs/ と evidence/ を宛先へまとめてコピー（robocopy の終了コードを判定）
- `manual_kitting_assistant` (1.1.0/3.7.0): `step_list.csv` と prompt/*.txt で手作業の手順を案内する GUI（Open/Copy ボタン）
- `network_profile_config` (1.0.0/2.0.0): NetworkList のカテゴリ（Public/Private/Domain）を設定
- `pianist` (2.1.0/3.7.0): 業務アプリの GUI 設定を、フェーズごとに SendKeys で半自動化（2,382 行）
- `reg_template` (1.2.0/3.7.0): レジストリキーの .reg バックアップ・インポート
- `script_looper` (2.0.0/2.0.0): 指定スクリプトを MaxRetry/IntervalSec/Condition に従って繰り返し実行
- `server_feature_config` (0.1.1/2.0.0): Windows Server の役割・機能をインストール

（extended の出典: 各 `modules/extended/<name>/` のスクリプト冒頭コメントと `module.csv`）
- 数の食い違い: README は「Extended モジュール（18）」「計 81 種」と書くが、同じ節の一覧も実際のディレクトリも 16 — [README.md:53-66, 252-265](README.md)、`ls modules/extended`

### Inferences
**Android 対応概念の対応表**（右の列はすべて推論・要 Android 側検証。分類はホスト PC + ADB モデルが前提）

| fabriq モジュール | Android での対応概念（想定する手段） | 分類 |
|---|---|---|
| app_config / winget_install | APK のサイドロード（`adb install -r [-g]`、split APK は `install-multiple`）。APK はホスト側の「アプリ資材」フォルダから。winget の「リポジトリ ID で取得」に当たるものは無い（Play ストアは手作業か EMM） | B |
| storeapp_config / bloatware_remove | プリインストールアプリの削除・無効化（`pm uninstall -k --user 0` / `pm disable-user --user 0`。戻すときは `cmd package install-existing`） | B |
| (bloatware_exporter / storeapp_editor アプリ) | 接続した端末の `pm list packages` から削除リストを作る GUI | B |
| reg_hklm_config / reg_hkcu_config / reg_template | `settings put/get system\|secure\|global <key> <val>`（レジストリの hive に相当する名前空間）。backup/restore は `settings list` の export/import。`secure` への書き込みは権限の制約あり | B（検証パターンは A） |
| spi_config | アニメーション倍率などの UI 設定（`settings put global window_animation_scale` 等） | B |
| power_config | 画面オフまでの時間（`settings put system screen_off_timeout`）、USB 給電中のスリープ無効（`svc power stayon`）、バッテリー最適化の除外（`dumpsys deviceidle whitelist`） | B |
| brightness_config | `settings put system screen_brightness` / `screen_brightness_mode` | B |
| volume_config | ストリームごとの音量（`cmd media_session volume` / 旧 `media volume`） | B |
| resolution_api_config / display_config | `wm size` | B |
| dpi_api_config / dpi_config | `wm density` | B |
| hostname_config | 端末名（`settings put global device_name`）、Bluetooth 名 | B |
| ssid_config | Wi-Fi ネットワークの追加（`cmd wifi add-network` / `connect-network` は Android の版に依存）。端末上のエージェントが要る可能性あり | B |
| ipaddress_config / temp_ipaddress_config / ipv6_config / network_profile_config | 静的 IP は ADB だけでは困難（エージェントか Device Owner が必要）。基本は DHCP の前提 | C（エージェント案なら B） |
| cert_config | CA 証明書の導入。新しい Android ではユーザー操作が必要になり、サイレント導入には Device Owner が要る | B（大きな制約あり） |
| time_sync_config | 時刻とタイムゾーンの自動設定（`settings put global auto_time/auto_time_zone`、タイムゾーンの指定） | B |
| wallpaper_config | 壁紙。汎用の shell コマンドは無い見込みで、エージェントアプリ（WallpaperManager）か手作業ガイド | B（エージェント） |
| default_app_config | 既定アプリの役割（`cmd role add-role-holder android.app.role.BROWSER <pkg>` 等） | B |
| copyfile_config / robocopy_config | `adb push` / `adb pull`（/sdcard 配下。scoped storage の制約あり） | B |
| file_delete / directory_cleaner | `adb shell rm [-r]` と、Android のパス規則版の `Test-FabriqProtectedPath`（§8 の安全ガードは必須） | B |
| process_killer | `am force-stop <pkg>` | B |
| generic_batch_runner / generic_process_runner / startup_command_config | 「任意の adb shell コマンド列」を実行するモジュール、またはホスト側スクリプトの実行（端末の次回起動時に 1 回実行する仕組みは無い） | B |
| script_looper / azure_ad_join_check | 繰り返し・再試行のラッパー（汎用）。AAD 確認自体は対応なし | A / C |
| restart_config / `__RESTART__` | `adb reboot` → `wait-for-device` → 起動完了をポーリング | B |
| signout_config | 対応なし（画面ロックはできる） | C |
| system_finalize | 仕上げ（ホーム画面に戻す、通知を消す、最後の再起動） | B |
| history_destroyer | キッティング中の痕跡の削除（`logcat -c`、`pm clear` の対象、/sdcard の一時ファイル、作業用 Wi-Fi の forget） | B |
| evidence_config | 端末インベントリ（getprop・パッケージ一覧と版・settings・battery・storage・セキュリティパッチ）と manifest | B |
| log_uploader | ホスト側の evidence を共有フォルダへ送る（ホストにはネットワークがあるので fabriq より単純） | A |
| hostlist_precheck | シリアル・機種・OS 版・バッテリー残量・ADB の認可状態などを事前照合するゲートと `__GATE__` | A（概念）/ B（実装） |
| manual_kitting_assistant | 自動化できない手順（Google アカウント、OEM の初期設定ウィザード、ストアでのログイン等）の案内と記録 | A（概念）/ B |
| pianist | `input tap/text/keyevent` と `uiautomator dump` による GUI の半自動操作（壊れやすい） | B |
| fabriq_app_launcher | ホスト側の補助ツールの起動、または `am start -n` / `monkey -p` で端末のアプリを起動 | B |
| test_harness_config / test_error_module | カーネル試験用の擬似モジュール（そのまま移植） | A |
| autologon_config / `__AUTO_to_*__` | 対応なし（画面ロックの解除・無効化（`locksettings`）が近いが別物） | C |
| local_user_config / profile_delete / group_config / builtin_admin_config | Android のマルチユーザー（`pm create-user`）は業務スマホではほぼ使わない | C |
| domain_join / gpo_config / ppkg_config | 対応なし（Android Enterprise の登録・MDM は fabriq_mac の方針「MDM なし」と同様に範囲外とする想定） | C |
| bitlocker_config | 対応なし（端末暗号化は標準で有効。evidence で状態を読むだけ） | C |
| firewall_config / firewall_rule_config / firewall_rule_make_config | 対応なし（root か VPN 型アプリが必要） | C |
| windows_license_config / office_license_config / odt_config / office_update | 対応なし（Office アプリは APK 導入として扱う） | C |
| windows_update / windows_feature_config / server_feature_config | OS 更新は ADB から起動できない見込みで、手作業ガイド（fabriq_mac の「OS 更新は最後」と同じ扱い）。機能の ON/OFF は `pm disable-user` のアプリ単位 | C（ガイド化） |
| driver_config / printer_driver_config / printer_delete / partition_config / restore_point / acl_config / scheduled_task_config / credential_config / browser_addon_config / startlayout_config / taskbar_config / desktop_icon_config | 対応なし。ホーム画面の配置はランチャー次第で、汎用 ADB は無い。プリンタは Mopria 等のプラグイン APK の導入に置き換え | C |

- 集計（上の表の推論による）: B が約 30、A が約 6、C が約 43。Android 版の初版に入れるべき中核は「APK 導入・削除」「settings 書き込みと検証」「Wi-Fi」「端末名」「画面・音量・明るさ」「時刻」「ファイルの push と削除」「再起動」「インベントリ」「precheck + ゲート」「手作業ガイド」の約 12 種類になる。

### Gaps
- Android 側のコマンドの可否（Android の版ごとの差、`cmd wifi` の有無、`settings secure` への書き込み権限、CA 証明書のサイレント導入の可否、壁紙・静的 IP の手段）は、このリポジトリ調査の範囲外。別途 Android 側の調査で検証が必要。
- 各モジュールの冪等性・Verified の実装状況は、手本の 4 つ以外は網羅的に確認していない（TM t-0081 / t-0082 が対応中・未着手）。

---

## 9. apps/ と dev ツールの Android 版での対応

### Takeaway
`apps/` には 9 つの GUI ツールがある（計 12,947 行。主役は operator ダッシュボードの 4,065 行）。Android 版に対応物が要るのは operator（ロットのマトリクスダッシュボード）、CSV 編集系（パッケージ削除リストの作成器など）、手作業ガイドの 3 系統。`fabriq_ios` は Apple の iOS ではなく **Cisco IOS 風のジョークシェル**なので、混同しないこと。

### Cited Findings
- `fabriq_operator`（4,065 行）: `dashboard_form`（メイン）、`session_form`（作業者・ホスト・パスフレーズ）、`flex_dashboard`（FlexProfile）、`execution_toolbar`（同一プロセスの STA Runspace。Skip / Gyotaq / PC Info、1,486 行）、`log_viewer`、`apps_dialog`、`quickactions_dialog`、`theme` — [apps/fabriq_operator/lib/](apps/fabriq_operator/lib/)
- `csv_editor`（1,109 行）: CSV を DataGridView で編集する（撤去を検討中: t-0099）— [apps/csv_editor/csv_editor.ps1](apps/csv_editor/csv_editor.ps1)、[.tm/tasks.json](.tm/tasks.json)
- `bloatware_exporter`（656 行）: インストール済みの旧来型アプリをレジストリから走査して `bloatware_list.csv` を編集 — [apps/bloatware_exporter/bloatware_exporter.ps1](apps/bloatware_exporter/bloatware_exporter.ps1)
- `storeapp_editor`（613 行）: インストール済みの Store アプリを一覧して `storeapp_list.csv` を編集 — [apps/storeapp_editor/storeapp_editor.ps1](apps/storeapp_editor/storeapp_editor.ps1)
- `winget_gui`（792 行）: winget のパッケージを検索して `app_list.csv` を編集 — [apps/winget_gui/winget_gui.ps1](apps/winget_gui/winget_gui.ps1)
- `local_user_setup`（458 行）: `local_user_list.csv` のプレースホルダ行をウィザード形式で埋める — [apps/local_user_setup/local_user_setup.ps1](apps/local_user_setup/local_user_setup.ps1)
- `desktop_icon_backup_app`（283 行）、`system_launcher`（301 行。Windows Search を使わず設定画面を開く。検索履歴を残さないため）— [apps/](apps/)
- `fabriq_ios`: 「Cisco IOS 風コマンドラインシェル `Fabriq IOS`…芸術的な副産物としての冗談アプリ」。モジュールの呼び出しは `Invoke-BatchExecution` 経由 — [apps/fabriq_ios/CLAUDE_CODE_PROMPT.md](apps/fabriq_ios/CLAUDE_CODE_PROMPT.md)
- `commands/`: `diag_crypto.ps1`、`explore_restart_command.ps1`、`get_evidence.ps1`、`gpupdate_command.ps1`、`system_launcher.ps1`、`temp_command.ps1` — [commands/](commands/)
- 別リポジトリの補助アプリ: Fabriq Studio（WPF/.NET 8。パスフレーズトークンの生成、hostlist・モジュール CSV・プロファイルの編集、preset のドロップダウン、レジストリカタログ）、Fabriq Evidence Manager（読むだけの consumer。Excel 納品物）、Fabriq Doc — [README.md:13-22, 377-414](README.md)
- ランチャ: `Fabriq.exe`（C#・未署名・`requireAdministrator`）と代替の `Fabriq.bat`。Defender の誤検知対策として `.bat` を並べて置いている — [README.md:44-49](README.md)、[dev/launcher/README.md](dev/launcher/README.md)

### Inferences
- **(B)** `fabriq_operator` → Android 版の**ホスト GUI の中心**。中身は、ロット受付（`adb devices` とシリアル照合）、端末×ステップのマトリクス（FlexProfile の拡張）、端末ごとのログビューア（`[Log]` の拡張）、端末ごとの Skip/撮影（Execution Toolbar の拡張）。
- **(B)** `bloatware_exporter` / `storeapp_editor` → 「基準機を 1 台つないで `pm list packages` を取り、削除・無効化リストを作る」GUI。`winget_gui` → 「APK 資材フォルダを走査して（`aapt` 等でパッケージ名と版を読み）アプリ一覧 CSV を作る」GUI（要検証）。
- **(B)** `manual_kitting_assistant` / `pianist` → Android では自動化できない手順がむしろ多い（アカウントの追加、OEM のウィザード、ストアの操作）。そのため手順ガイドとチェック記録の重要度が上がる。
- **(A)** Studio / Evidence Manager との互換は、fabriq_mac と同じく**追わない**判断が妥当（決定 6）。ただし `EVIDENCE_MANIFEST.md` の前方互換の考え方だけは自前の納品物に取り入れる。
- **(B)** ホストツールを Windows の exe で配るなら、fabriq と同じく**未署名の小さな PE ランチャーが Defender に誤検知されるリスク**を抱える（t-0059/t-0060）。最初からコード署名を予算に入れておくべき。
- **(C)** `system_launcher`、`desktop_icon_backup_app`、`local_user_setup`、`commands/gpupdate_command.ps1` 等は Windows 専用。

### Gaps
- Studio と Evidence Manager の実装は別リポジトリにあり、未確認。

---

## 10. 規模の定量化（移植の見積り用）

### Takeaway
カーネル（common 5,500 行 + main 2,296 行 ≒ 7,800 行・関数 123）、GUI アプリ約 1.3 万行、モジュール約 3.65 万行（104 スクリプト）、テスト約 8,200 行（671 `It`）。Android 版の初版でカーネル相当（結果契約・CSV ローダ・暗号・プロファイル解析・実行履歴・チェックリスト・ゲート・端末ワーカ）を作り、モジュールを約 12 種類に絞れば、fabriq の 1〜2 割程度の規模から始められる見込み（推論）。

### Cited Findings
- `kernel/common.ps1` 5,500 行・関数 114（`^function`）。`kernel/main.ps1` 2,296 行・関数 9。`kernel/ps1/manifesto.ps1` 214 行、`kernel/ps1/view_report.ps1` 78 行。`kernel/KERNEL_API.md` 586 行、`kernel/EVIDENCE_MANIFEST.md` 176 行 — `wc -l` / `grep -c '^function '`（[kernel/](kernel/)）
- 公開 API の数: 関数 32（§1.1:7、§1.2:3、§1.3:3、§1.4:3、§1.5:2、§1.6:2、§1.7:2、§1.8:2、§1.9:3、§1.10:4、§1.11:1）、公開グローバル 6、特殊マーカー 6、ModuleResult のフィールド 6 — [kernel/KERNEL_API.md:11-193](kernel/KERNEL_API.md)
- モジュール: ディレクトリ 79（standard 63 / extended 16）、`.ps1` 104 本・36,548 行、モジュールフォルダ内の CSV 218 個（module.csv と preset.csv を含む）。大きいもの: evidence_config 3,250、pianist 2,382、power_config 1,164、printer_driver_config 1,156、bitlocker_config 1,110、dpi_config 902 行 — `modules/*/*`（集計）
- apps: `.ps1` 計 12,947 行（fabriq_operator 4,065 を含む）。dev: `.ps1` 計 4,395 行 — `apps/`、`dev/`（集計）
- テスト: `*.tests.ps1` 41 ファイル（tests/ 31 + fabriq_ios 10）、`It` 671 個（tests/ 586 + fabriq_ios 85）、tests/ の `.ps1` 計 8,163 行。モジュールのテスト記述子 `test.psd1` は 51 — [tests/](tests/)、[apps/fabriq_ios/tests/](apps/fabriq_ios/tests/)
- CHANGELOG 6,265 行・版見出し 31。TM タスク 110。git コミット 50（このクローンで見える範囲）— [CHANGELOG.md](CHANGELOG.md)、[.tm/tasks.json](.tm/tasks.json)、`git log --oneline | wc -l`

### Inferences
- 移植の重みは「カーネルの契約層（約 2〜3 千行相当: 結果・CSV・暗号・秘密値マスク・プロファイル解析・ゲート・実行履歴・HTML）」と「ホストの並列実行層（端末ワーカ・ADB ラッパ・タイムアウト・再接続待ち。fabriq に相当するものが無い**新規**部分）」に分かれる。後者は fabriq から写せない新しい設計で、リスクが最も大きい（推論）。
- テストは、カーネル契約向けの 354 `It` が仕様書の役割を果たす（例: Import-ModuleCsv 16、New-BatchResult 18、Resolve-ProfileModules 21、GateBarrier 15、ResumeState 29）。Android 版では、これらのテスト名の一覧をそのまま受け入れ基準のチェックリストとして使える（推論）。

### Gaps
- `.git` の履歴が浅く（50 コミット）、開発にかかった工数（人日）は推定できない。

---

## 11. 横断: A/B/C 分類一覧と「端末上エージェント」案での差分

### Takeaway
1 対多のホスト PC モデルでは、fabriq の**契約・データ形式・プロセス**はほぼすべて A、**実行の仕組み**（プロセス全体の状態・確認ダイアログ・再起動からの再開・スクショ・生存シグナル）は B、**OS の機能を扱うモジュール**の過半は C になる。端末上にエージェントアプリを置く 1 対 1 案は fabriq の形に近いが、Android のサンドボックスと権限の制約のため、結局ホスト側の ADB を併用することになる見込み。

### Cited Findings
- 1 対 1 前提の根拠（fabriq の現状）: 対象 PC の上で、USB の作業ディレクトリから管理者として起動する（`Fabriq.exe`/`.bat` → `kernel/main.ps1`）。セッションのメディアシリアルは USB ボリュームのシリアル。状態ファイルは `kernel/json/` に置く — [README.md:44-49](README.md)、[kernel/main.ps1:1636-1641](kernel/main.ps1)、[kernel/KERNEL_API.md:215](kernel/KERNEL_API.md)
- プロセス全体の状態の根拠: §2 の公開グローバル変数と §3 の環境変数、`Invoke-BatchExecution` 内で実行ごとに `FABRIQ_SEGMENT` / `FABRIQ_AUTOLOGON_USER` を設定・解除 — [kernel/KERNEL_API.md:114-146](kernel/KERNEL_API.md)、[kernel/main.ps1:620-660](kernel/main.ps1)
- 端末を跨いだ安全性の先例: `Load-ResumeState` は HardwareUniqueId が違う状態ファイルを拒否する — [kernel/common.ps1:3764-3785](kernel/common.ps1)

### Inferences
**A（そのまま移植）**
- ModuleResult / BatchResult 契約（5 値の Status、範囲付き三値の Verified、件数から Status を決める規則、無返却は Error）
- モジュール構造（フォルダ = モジュール、`module.csv`、`*_list.csv`、`preset.csv`、Step 1〜6 + 5.5、Enabled/Segment/RequiredColumns、`$null` と空配列の区別）
- プロファイル CSV（Order/ScriptPath/Enabled/Description/Segment/ErrorMode/Group）、`__GATE__`、ErrorMode skip/retry（上限 5 回）
- `ENC:` 暗号形式（AES-256-CBC + PBKDF2）とパスフレーズ検証トークン、秘密値レジストリ（3 箇所の出力でマスク）。ただし**固定ソルトは t-0075 の弱点なので、新設計では blob ごとの乱数ソルトに改める**
- プロファイル別データオーバーレイ（Profile-First、フォールバックは必ず可視化）
- 実行履歴 CSV（追記専用の列・後方互換）、HTML チェックリスト（期待値と実測・OK/NG/Incomplete）、manifest の前方互換ルール、log_uploader、JSONL テレメトリ
- workers.csv / categories.csv、hostlist の「1 行 1 台」と `ENC:` セル
- ガバナンス一式（公開 API 文書、カーネルとモジュールの独立 SemVer、REQUIRES_KERNEL、CHANGELOG、設計ゲート、fail-closed 原則、CI で exit code 判定、TM 台帳、テスト記述子）

**B（1 対多向けに作り直し）**
- セッション: 1 セッション = 1 ロット。作業者とパスフレーズは 1 回、ホスト選択 → 接続 N 台のシリアル照合（シリアル → 名前 → 前回 → 手動）
- 端末ごとの文脈: `SELECTED_*` / `$global:` の代わりに、端末ごとにワーカプロセスを分ける（環境変数の契約を保てる）か、コンテキストを明示的に渡す
- 実行エンジン: 端末ワーカのプール、ADB 呼び出しごとのタイムアウト、切断・再接続（USB の抜き差し）への耐性、端末ごとの Skip
- 確認: モジュールごとの Y/N を廃止し、ロットで 1 回確認する。ErrorMode 空欄時のモーダル → 端末ごとの「要対応」セル
- FlexProfile → 端末×ステップのマトリクス（電子チェックシート）。状態はシリアルごとに保存
- `__RESTART__` → `adb reboot` と起動完了待ち（ホストが状態を持ち続ける）。resume_state はホストのクラッシュ対策としてだけ残す
- エビデンス: `adb exec-out screencap` によるステップごとの撮影、端末ごとのフォルダ、ロット一覧表
- 生存シグナル（ArtPulse）→ 端末ワーカのハートビートと `adb devices` の状態
- Execution Toolbar → 端末ごとの操作パネル
- `Get-HardwareUniqueId` → ADB のシリアル / `ro.serialno`（要検証）。`Test-AdminPrivilege` → 「USB デバッグが認可済みか」（`unauthorized` 状態）の検出
- `Test-FabriqProtectedPath` → Android のパス規則（/sdcard、/data/local/tmp）版
- 補助アプリ（パッケージリスト作成器・APK 資材カタログ・手作業ガイド）

**C（Windows 専用）**
- RunOnce・DPAPI による再開、AutoLogon と WU ループ、Active Setup / Startup の配備、HKCU の解決、Zone.Identifier、Explorer の再起動、SleepSuppression・QuickEdit、`New-UiFont` と `SetProcessDPIAware`（ホスト GUI を WinForms で作るなら、ホスト側には残る）、モジュールの約 43 種（§8）

**端末上エージェントアプリ案（1 対 1）にした場合の差分**（すべて推論・要検証）
- fabriq の「対象機の上で動く」形にほぼ戻るので、RunOnce 相当は `BOOT_COMPLETED` レシーバ、resume_state 相当はアプリ内部ストレージ、`__RESTART__` 跨ぎの再開ロジック（Save/Load-ResumeState の思想）が再び **A に近い B** になる。
- 権限の壁: `settings secure/global` の書き込みには `WRITE_SECURE_SETTINGS`（ADB の `pm grant` で付与）、サイレントなアプリ導入・削除、CA 証明書、静的 IP、壁紙などには Device Owner（DPC）が実質必要。Device Owner 化は初期化直後の端末で QR か `dpm set-device-owner` を使う、つまり「MDM なし」方針とぶつかる部分がある。
- スクリーンショットは MediaProjection（ユーザーの同意ダイアログ）が必要になり、ホスト側 `screencap` より面倒になる。エビデンスはエージェントが書いて、最後にホストが `adb pull` で回収する構図が現実的。
- 配布: エージェント APK と設定 CSV（`ENC:` 付き）を ADB で push する必要があり、結局ホストが要る。よって**ハイブリッド**（ホストが受付・配布・回収・一覧を担当し、エージェントは ADB で届かない操作だけを担当）が最も現実的。fabriq の「カーネル（ホスト）とモジュール（操作単位）」の分離は、このハイブリッドでも保てる。
- fabriq_mac の決定 8（SwiftUI の GUI が bash のコアを呼ぶ）と同様に、「ホスト GUI（WinForms/WPF）が CLI コア（PowerShell か別言語）を呼ぶ」二層にすると、コアを CI で単体テストできる構造になる。

### Gaps
- 端末上エージェント案の Android 権限モデル（版ごとの差、Device Owner の可否、MediaProjection の運用）は、このリポジトリからは検証できない。
- fabriq_mac のリポジトリはこのクローンに無く、課題文の記述（決定 1〜8）以上は確認できない。

---

## 12. 既知の痛点（.tm/tasks.json）と Android 設計への教訓

### Takeaway
fabriq の未解決・最近の不具合の多くは「**無人実行が止まる（ハング）**」「**Windows Defender の誤検知**」「**DPI とスクリーンショット**」「**偽の成功・偽の検証**」「**冪等性**」に集まっている。1 対多の Android ツールでは、ハングと偽成功の影響が N 倍になるので、最初から設計で封じる必要がある。

### Cited Findings
- AutoPilot が止まる系: t-0053（bloatware_remove が UninstallString を `-Wait` で上限なく待つ）、t-0054（`Wait-NetworkReady` が `while($true)` で上限なし。実装は common.ps1:1610-1627 で確認済み）、t-0055（ErrorMode 空欄時の MessageBox にタイムアウトが無い）、t-0056（モジュール内の生の `Read-Host` に AutoPilot ガードが無い）— [.tm/tasks.json](.tm/tasks.json)、[kernel/common.ps1:1610-1627](kernel/common.ps1)
- Defender: t-0059（`Fabriq_IOS.exe` が `Program:Win32/Wacapew.A!ml` と判定され削除された → 自己昇格する `.bat` に戻した）、t-0060（恒久対策の案: コード署名 / 誤検知の申告 / パス除外。`Fabriq.exe` も構造的に同じリスク）、3.4.0 で別プロセスの Status Monitor を Defender/ASR に止められて in-process に変更 — [.tm/tasks.json](.tm/tasks.json)、[kernel/KERNEL_API.md:328-351](kernel/KERNEL_API.md)
- DPI: t-0084（GUI のホストプロセスが DPI awareness を宣言しておらず、ぼやけたりサイズが変わったりする）、t-0085・t-0101（スクリーンショットの見切れ。WinForms `Screen` の静的キャッシュが原因）、t-0091（`New-UiFont`）— [.tm/tasks.json](.tm/tasks.json)
- 偽成功・偽検証: t-0012（WU のダウンロード失敗が Success になる）、t-0013（OSPP の終了コードだけで判定）、t-0015（robocopy の終了コードを無視）、t-0017（power_config に失敗カウンタが無く常に Success）、t-0020（検証が検証になっていない 8 モジュール）、t-0044（全行 Disabled のとき Skipped でなく Error になる）、t-0109（テストリグが `expect.verified=$true` を検査していない）— [.tm/tasks.json](.tm/tasks.json)
- 冪等性: t-0019（再実行で偽 Error・偽 Skip になる 4 モジュール）、t-0058・t-0082（冪等性の台帳は未着手）— [.tm/tasks.json](.tm/tasks.json)
- 暗号: t-0075（固定ソルト。全配備が同じ鍵空間になる。対策案は乱数ソルトを前置する `ENC2:` 形式）— [.tm/tasks.json](.tm/tasks.json)
- 再起動跨ぎ: t-0039（FlexProfile で `__RESTART__` を単発実行すると、再起動後に自動起動しない。保留）、t-0045（RunOnce 登録の失敗で fail-open → 修正済み）— [.tm/tasks.json](.tm/tasks.json)
- 破壊的削除: t-0009・t-0041（`\\?\`・UNC でパスガードが突破され得る）、t-0042・t-0049（`robocopy /MIR` が空のソースで宛先を全削除する）— [.tm/tasks.json](.tm/tasks.json)

### Inferences
- **(A) 設計原則としてそのまま採るもの**: ①すべての外部呼び出し（adb・pm・am・ネットワーク待ち）に上限時間を必須にする。②無人実行中の対話（モーダル・`Read-Host`）はカーネルが一律に禁止し、代わりに「要対応」状態にする。③終了コードと出力の両方で成否を判定し、検証ゼロのときは `$null` にする（範囲付き Verified）。④全行無効や Segment 不一致は Skipped にする（Error にしない）。⑤削除系はパスガードを確認ゲートの外に置く。⑥暗号は blob ごとの乱数ソルトで最初から作る。
- **(B)** Defender の問題は、ホストを Windows にするなら Android 版でも同じく起こる。PowerShell と小さな未署名 exe の組み合わせが最も危ないので、署名済みのバイナリ（.NET 単体の exe 等）か署名付きスクリプトを検討すべき（推論）。
- **(B)** DPI の問題は、ホスト GUI を WinForms で作るなら同じく起こる。最初のウィンドウを作る前に DPI awareness を宣言する、という fabriq の教訓を初日から入れる。スクリーンショットは `screencap` で端末側から取るので、DPI の問題は生じない（推論）。

### Gaps
- 各 TM タスクの `claudeNote`（調査結果）は全部は読んでいない。重要度の判定は、タスクのタイトルと content の冒頭による。
