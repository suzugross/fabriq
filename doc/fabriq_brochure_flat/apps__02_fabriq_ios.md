# fabriq_ios — Cisco IOS 風シェル

`apps/fabriq_ios/` は fabriq フレームワーク上に被せた **Cisco IOS スタイルのコマンドラインシェル**です。SPEC.md にて「シュルキティニスム宣言の芸術部門」と明確に位置付けられており、実用ツールではなく **art object / 思想の戯画化コンポーネント**として存在します。それでもタブ補完や省略補完は本物の Cisco IOS に近い水準で作り込まれています。

## VERSION (独立 SemVer)

`apps/fabriq_ios/VERSION` は現在 **0.3.5**。fabriq カーネル / モジュールの SemVer とは独立して進化します。`show version` および `show running-config` コマンドが参照するほか、KERNEL_API.md §6 の internal API (`Invoke-BatchExecution`, `Initialize-ModuleSystem`, `Resolve-ProfileModules`, `Test-MasterPassphrase`) に依存しているため、kernel PATCH 昇格のたびに再検証する旨が README に明記されています。

## エントリポイントと self-spawn 隔離

`fabriq_ios.ps1` の冒頭に **self-spawn ガード**があります。

```
if (-not $env:FABRIQ_IOS_SUBPROCESS) {
    $env:FABRIQ_IOS_SUBPROCESS = '1'
    Start-Process powershell.exe -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File',$self -Wait
    return
}
```

これは fabriq_operator の FabriqApps ダイアログから `& $appPath` で in-process 起動された場合に、**PSReadLine の KeyHandler / 環境変数の改変 / global スコープ変数**が呼び出し元の powershell に漏れないようにするためです。`-NoProfile -ExecutionPolicy Bypass -File` で隔離された子 powershell.exe がシェルを実行し、終了すると環境変数も連れて消えます。

その後、kernel/common.ps1 を dot-source してから次の関数を **no-op で global シャドウ**します:

- `Initialize-ExecutionHistory`, `Restore-ExecutionHistory`, `Write-ExecutionHistory`, `Add-ExecutionResult`
- `Export-ExecutionHistory`, `Export-HtmlChecklist`
- `Initialize-EvidenceBasePath`, `Capture-ScreenEvidence`

これにより、モジュールが間違って呼んでも fabriq_ios の操作が親 fabriq の audit trail を汚さない保証が立ちます (REPL は履歴を残さない)。

## モード構造 (5 階層)

`lib/shell_state.ps1` の `New-ShellState` で初期 Mode='UserExec'。許可遷移は `Set-ShellMode` の whitelist。

```
UserExec        fabriq>            参照系のみ。enable で昇格
PrivilegedExec  fabriq#            disable / configure terminal / reload / show
GlobalConfig    fabriq(config)#    hostname / interface / module / cleanup / copy / install / script
InterfaceConfig fabriq(config-if)# ip address ...
ModuleConfig    fabriq(config-mod)# / (config-clean)# / (config-copy)# / (config-install)# / (config-script)#
```

ModuleConfig の prompt suffix は `data/module_categories.json` の各 category の `promptSuffix` を引いて動的決定 (例: `module bitlocker_config` で入った場合は config-mod、`cleanup directory_cleaner` なら config-clean)。

各モードのディスパッチャは `lib/modes/<mode>.ps1` に分離 (`Invoke-UserExecCommand` 等)。

## コマンドパーサ (`lib/parser.ps1`)

Cisco IOS の prefix-resolution を厳密に再現:

- `ConvertTo-FabriqIosTokens` — quote 対応の単純なトークナイザ
- `Get-FabriqIosCommandVocabulary` — モードごとのコマンド一覧 (UserExec=`enable show help exit ?`, PrivilegedExec=`show configure reload disable exit help ?` など)
- `Resolve-Token` — 完全一致を優先し、prefix 一致が 1 候補のときだけ採用、複数候補は ambiguous エラー (`% Ambiguous command`)
- `Get-FabriqIosSubVocabulary` — 2 層目語彙 (`PrivilegedExec.show` → `version running-config profiles modules evidence manifesto` など)

これにより `conf t`、`sh ru`、`int eth0` のような Cisco 流の省略入力が動きます。

## タブ補完 (`lib/completer.ps1`)

PSReadLine 統合。`Register-FabriqIosCompleter` がシェル開始時に Tab と `?` をフックします。コア関数 `Get-FabriqIosCompletion` は **pure function** で、(Line, Position, Mode, State) を取り候補配列を返すだけ。PSReadLine 抜きで単体テスト可能 (実際 `tests/completer.tests.ps1` がそのテストハーネス)。

動的補完ソース:

- `hostname <TAB>` → `kernel/csv/hostlist.csv` の NewName 列
- `interface <TAB>` → `Get-NetAdapter` の InterfaceAlias (日本語「イーサネット」も候補に出る — むしろ味)
- `show <TAB>` → モード別 sub-vocabulary
- `module <TAB>` / `cleanup <TAB>` / `install <TAB>` 等 → `module_categories.json` の対応 category 内のモジュール名
- ModuleConfig の `set <col> <val>` / `add <col> <val>` 系は **位置パリティ**で次候補を col 名 / 値 enum に切り替え

`?` キーは Cisco 慣習を再現:

- 空 prompt で `?` → そのモードの全 vocabulary を help_text.csv から引いて表示 (deferred AcceptLine 経由)
- 入力途中で `?` → 候補一覧を表示してバッファを復元、編集を続行可能

## 主要コマンド (`lib/commands/*.ps1`)

| ファイル | 提供コマンド |
|---|---|
| `show.ps1` | show version / show host(s) / show running-config / show profiles / show modules / show evidence / show manifesto |
| `enable_disable.ps1` | enable (passphrase 認証で PrivilegedExec へ) / disable (UserExec へ降格) |
| `hostname.ps1` | hostname `<NewName>` で hostlist エントリ選択 + hostname_config 起動 |
| `interface.ps1` | interface `<alias>` で InterfaceConfig 入り |
| `ip_address.ps1` | ip address from-hostlist / ip address `<ip> <mask>` |
| `categories.ps1` | module / cleanup / copy / install / script の category 切替 |
| `module.ps1` | module `<name>` で ModuleConfig 入り |

`enable` は `Test-MasterPassphrase` を流用するため、fabriq の DPAPI 暗号化済みパスフレーズと同一のものを要求します。

## syslog システム (`lib/syslog.ps1` + `data/syslog_messages.csv`)

Cisco IOS 完全準拠の syslog 行を生成します:

```
*Apr 29 14:23:01.234: %FABRIQ-5-HOSTNAME: The name NEW-PC-01 has been carved into the silicon. The machine awaits the next dawn.
```

書式:

- Timestamp は `*MMM dd HH:mm:ss.fff` (英語 month 強制、locale 非依存)
- Severity 0–7 (Cisco 準拠: Emergency / Alert / Critical / Error / Warning / Notification / Informational / Debug)
- Mnemonic = `FABRIQ-X-<FACILITY>` (HOSTNAME, INTERFACE, IPADDR, MANIFESTO, AUTOMATE, RESTART, ENABLE, EXIT)

メッセージテンプレートは `data/syslog_messages.csv` で外部編集可能。例:

```
HOSTNAME,5,success,The name {NewName} has been carved into the silicon. The machine awaits the next dawn.
HOSTNAME,4,reboot_required,The machine must close its eyes and open them again to remember its new name.
IPADDR,6,whispered,Address {Ip}/{Prefix} has been whispered to interface {Interface}. The wire understands.
INTERFACE,5,configured,"Interface {Alias}, the metallic vein, is now configured."
MANIFESTO,7,quote,"Surkittinism is the convulsive beauty of mass deployment, or it is nothing."
```

設計思想 3 本柱 (SPEC.md より):

1. **タブ補完は真面目に** — 操作体験は本物の Cisco IOS 水準
2. **syslog メッセージは Cisco 風英語、ただしシュルレアリスム** — 形式厳格・内容詩的
3. **コマンド階層は Cisco IOS 丸パクリ**

## バナー (`data/version_banner.txt`)

起動時に表示される:

```
Fabriq IOS Software, Version 3.0(1)Surkittinism
Copyright (c) 1924-2026 by Andre Breton & Anonymous Kitting Operators.
Compiled in the Dream Hours by automatic-writing.

The machine yawns. Press RETURN to begin the seance.
```

著作権年 1924-2026 は André Breton の『シュルレアリスム宣言』(1924) を起点としており、fabriq_ios の "art object" としての自意識を体現しています。

## テスト (`tests/`)

Pester 5 ベース。モジュール完成度ごとに `_phase3_smoke.ps1` ... `_phase9b_smoke.ps1` および `parser.tests.ps1`, `completer.tests.ps1`, `prompt.tests.ps1`, `shell_state.tests.ps1`, `syslog.tests.ps1` が存在。`Get-FabriqIosCompletion` 等の pure function はそのまま単体テスト可能。

## 存在理由

「真面目なフレームワークの中に冗談を一つ混ぜる」「Cisco ルータと日本語 Windows キッティング現場という本来交わらない二つの世界をロートレアモン的に手術台の上で出会わせる」(SPEC.md より要約)。kernel 内部 API への結合は「the price of the joke」として README 上で意識的に受容されており、kernel PATCH のたびに再検証することが運用上のルール。

```
fabriq>                              User EXEC mode (read-only)
fabriq#                              Privileged EXEC mode (after enable)
fabriq(config)#                      Global Configuration mode (after configure terminal)
fabriq(config-if)#                   Interface Configuration mode (after interface XXX)
fabriq(config-mod)#                  Module Configuration mode (after module XXX)
```
