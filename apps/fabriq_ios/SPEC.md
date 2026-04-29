# Fabriq IOS - 仕様書（v0.1 初期構想）

## 概要

**Fabriq IOS** は、Fabriq フレームワークの上に被せる Cisco IOS 風コマンドラインシェルである。
実用ツールではなく、**シュルキティニスム宣言の芸術部門**として位置付けられる「冗談アプリ」だが、タブ補完など基本的な操作性は真面目に作り込む。

- 親プロジェクト: Fabriq v3.0（Windows PC キッティングフレームワーク）
- 配置: `apps/fabriq_ios/` 配下（または別リポジトリ `fabriq-ios` として独立）
- 言語: PowerShell 5.1
- 依存: Fabriq kernel（`Initialize-ModuleSystem`, `Resolve-ProfileModules`, `Invoke-BatchExecution` 等）、PSReadLine

## 設計思想

### 三本柱

1. **タブ補完は真面目に** - 操作体験は本物の Cisco IOS に近い水準を目指す
2. **syslog メッセージは Cisco 風英語、ただしシュルレアリスム** - 構文は `*Mar 1 00:00:00.000: %FACILITY-SEVERITY-MNEMONIC: message` 形式を踏襲、内容は詩的・不条理
3. **コマンド階層は Cisco IOS 丸パクリ** - User EXEC → Privileged EXEC → Global Config → Interface Config の四階層

### 文化的位置付け

- Fabriq の実用GUIダッシュボードと並列に存在する「思想の戯画化」コンポーネント
- Cisco ルータと日本語Windowsキッティング現場という、本来交わらないはずの二つの世界が手術台の上で出会う
- ロートレアモン的な異質なものの邂逅を技術的に再現する

## コマンド階層

```
fabriq>                              User EXEC mode（参照系のみ）
fabriq#                              Privileged EXEC mode（enable後）
fabriq(config)#                      Global Configuration mode（configure terminal後）
fabriq(config-if)#                   Interface Configuration mode（interface XXX後）
```

### モード遷移

| 現モード | コマンド | 遷移先 |
|---------|---------|-------|
| User EXEC | `enable` | Privileged EXEC |
| Privileged EXEC | `disable` | User EXEC |
| Privileged EXEC | `configure terminal` (`conf t`) | Global Config |
| Global Config | `exit` / `end` | Privileged EXEC |
| Global Config | `interface <name>` | Interface Config |
| Interface Config | `exit` | Global Config |
| Interface Config | `end` | Privileged EXEC |
| 任意 | `exit` (User EXEC時) | シェル終了 |

## 実装するコマンド（v0.1スコープ）

### User EXEC mode

| コマンド | 説明 | 対応Fabriq動作 |
|---------|------|---------------|
| `enable` | 特権モードへ昇格 | パスフレーズ入力（`Unprotect-FabriqValue` で検証） |
| `show version` | バージョン情報表示 | Fabriq IOS版数、シュルキティニスム宣言年 |
| `show host` | 選択中のホスト情報 | `$env:SELECTED_NEW_PCNAME` 関連の環境変数を表示 |
| `show hosts` | hostlist.csv 一覧 | `kernel/csv/hostlist.csv` を整形表示 |
| `?` / `help` | 利用可能コマンド一覧 | 動的生成 |
| `exit` | シェル終了 | プロセス終了 |

### Privileged EXEC mode

| コマンド | 説明 |
|---------|------|
| `show running-config` | 現在の選択ホスト＋プロファイル状態を Cisco 風に出力 |
| `show profiles` | profiles/*.csv 一覧 |
| `show modules` | modules/standard 配下の MenuName 一覧 |
| `show evidence` | 直近の evidence ディレクトリ概要 |
| `show manifesto` | シュルキティニスム宣言を表示 |
| `configure terminal` | Global Config モードへ |
| `reload` | `__RESTART__` 発火（実装は後回し、v0.1では宣言だけ） |
| `disable` | User EXEC へ降格 |

### Global Config mode

| コマンド | 説明 | 対応Fabriq動作 |
|---------|------|---------------|
| `hostname ?` | hostlist.csv の NewName 一覧をタブ補完候補として提示 | hostlist.csv から動的生成 |
| `hostname <NewName>` | ホスト選択＋hostname_config 実行 | 環境変数セット → 当該モジュール実行 → Cisco風syslogで再起動要求を詩的に通知 |
| `interface ?` | 利用可能NIC一覧 | `Get-NetAdapter` の InterfaceAlias 一覧（日本語名「イーサネット」も含む） |
| `interface <name>` | Interface Config モードへ | `$script:CurrentInterface` にセット |
| `exit` / `end` | 上位モードへ |

### Interface Config mode

| コマンド | 説明 | 対応Fabriq動作 |
|---------|------|---------------|
| `ip address ?` | hostlistから引いた候補を提示 | 選択中ホストのEthernetIP/SubnetMask/Gateway/DNSを候補表示 |
| `ip address <ip> <mask>` | IPアドレス設定 | hostlist記載の値で `ipaddress_config` 一式を実行 |
| `ip address from-hostlist` | hostlistから一括設定 | 上記のシュガー構文 |
| `exit` | Global Config へ |

## タブ補完仕様

### 真面目に作る範囲

- **コマンド名補完**: 各モードで利用可能なコマンドをアルファベット順に提示
- **動的補完**:
  - `hostname <TAB>` → hostlist.csv の NewName 列
  - `interface <TAB>` → `Get-NetAdapter | Select InterfaceAlias`
  - `show <TAB>` → 利用可能な show サブコマンド
- **省略補完**: Cisco IOS と同じく `conf t` = `configure terminal`、`sh ru` = `show running-config`
- **`?` ヘルプ**: コマンド途中で `?` を押すとヘルプが出る（PSReadLine の Key Handler で実装）

### 妥協する範囲

- 引数のリアルタイム検証は最小限
- IPアドレス形式の妥当性チェックは `ipaddress_config` モジュール側に丸投げ

## syslog メッセージ仕様

### 書式（Cisco IOS 完全準拠）

```
*Apr 29 14:23:01.234: %FABRIQ-5-HOSTNAME: Host name of the machine has been inscribed as NEW-PC-01. Reboot is required for the dream to take effect.
*Apr 29 14:23:02.567: %FABRIQ-6-INTERFACE: Interface イーサネット, the metallic vein, is now configured.
*Apr 29 14:23:03.890: %FABRIQ-4-RESTART: The machine shall close its eyes and open them again. Save your unfinished poems.
```

### Severity 値（Cisco IOS 準拠）

| 値 | 名称 | Fabriq IOSでの用法 |
|----|------|------------------|
| 0 | Emergency | システム停止級（ほぼ使わない） |
| 1 | Alert | 致命的エラー |
| 2 | Critical | 重大エラー |
| 3 | Error | 通常エラー |
| 4 | Warning | 再起動要求等の注意事項 |
| 5 | Notification | 設定変更の通知 |
| 6 | Informational | 一般情報 |
| 7 | Debug | デバッグ |

### Mnemonic（FACILITY部分）

- `FABRIQ-X-HOSTNAME`: ホスト名関連
- `FABRIQ-X-INTERFACE`: インターフェース関連
- `FABRIQ-X-IPADDR`: IP設定関連
- `FABRIQ-X-MANIFESTO`: 宣言関連（特殊コマンド時）
- `FABRIQ-X-AUTOMATE`: 自動筆記モード（エラー時の詩的応答）

### メッセージ生成ルール

1. **形式は厳格に Cisco 風**
2. **内容はシュルレアリスム的**:
   - 機械を擬人化する（`The machine sleeps`, `the cable dreams of electrons`）
   - 自動筆記的な不条理（突然動植物や天体が出てくる）
   - キッティング作業を儀式・芸術行為として描写
3. **エラーメッセージほど詩的に** - 失敗ほど自動筆記が深まる
4. **メッセージライブラリは外部CSV**: `apps/fabriq_ios/syslog_messages.csv` で管理し、編集可能にする

### サンプルメッセージ集

成功系:
- `%FABRIQ-5-HOSTNAME: The name has been carved into the silicon. NEW-PC-01 awaits the next dawn.`
- `%FABRIQ-6-IPADDR: Address 192.168.1.100/24 has been whispered to the network adapter. The wire understands.`
- `%FABRIQ-5-INTERFACE: Interface イーサネット is now an open mouth speaking in packets.`

エラー系:
- `%FABRIQ-3-HOSTNAME: The machine refuses the name you offered. Perhaps it dreams of another.`
- `%FABRIQ-4-IPADDR: The address you proposed is already inhabited by an unknown butterfly.`

特殊系:
- `%FABRIQ-7-MANIFESTO: Surkittinism is the convulsive beauty of mass deployment, or it is nothing.`

## ファイル構成

```
apps/fabriq_ios/
├── README.md                       # 「これは芸術作品です」
├── fabriq_ios.ps1                  # メインエントリポイント
├── lib/
│   ├── parser.ps1                  # コマンドパーサ
│   ├── completer.ps1               # PSReadLine タブ補完ハンドラ
│   ├── prompt.ps1                  # プロンプト生成（モード反映）
│   ├── syslog.ps1                  # syslog風メッセージ出力
│   ├── modes/
│   │   ├── user_exec.ps1
│   │   ├── privileged_exec.ps1
│   │   ├── global_config.ps1
│   │   └── interface_config.ps1
│   └── commands/
│       ├── show.ps1                # show系コマンド全般
│       ├── hostname.ps1
│       ├── interface.ps1
│       ├── ip_address.ps1
│       └── enable_disable.ps1
├── data/
│   ├── syslog_messages.csv         # syslog風メッセージライブラリ
│   ├── help_text.csv               # `?` で出るヘルプ文
│   └── version_banner.txt          # `show version` の出力
└── tests/
    ├── parser.tests.ps1
    └── completer.tests.ps1
```

## 起動シーケンス

```powershell
# 起動例
PS> .\apps\fabriq_ios\fabriq_ios.ps1

Fabriq IOS Software, Version 3.0(1)Surkittinism
Copyright (c) 1924-2026 by André Breton & Anonymous Kitting Operators.
Compiled in the Dream Hours by automatic-writing.

The machine yawns. Press RETURN to begin the séance.

fabriq> enable
Passphrase: ********

fabriq# show host
Selected Host: NEW-PC-01 (AdminID: 1)
  OldName:     OLD-PC-01
  NewName:     NEW-PC-01
  EthernetIP:  192.168.1.100/24
  Gateway:     192.168.1.1
  DNS:         192.168.1.10, 192.168.1.11

fabriq# configure terminal
Enter configuration commands, one per line. End with CNTL/Z.

fabriq(config)# hostname ?
  NEW-PC-01     [hostlist.csv #1]  current selection
  NEW-PC-02     [hostlist.csv #2]
  NEW-PC-03     [hostlist.csv #3]
  ...

fabriq(config)# hostname NEW-PC-01
*Apr 29 14:23:01.234: %FABRIQ-5-HOSTNAME: The name NEW-PC-01 has been carved into the silicon.
*Apr 29 14:23:01.456: %FABRIQ-4-RESTART: The machine must close its eyes and open them again to remember its new name.

fabriq(config)# interface ?
  イーサネット            Ethernet adapter
  Wi-Fi                Wireless adapter
  Bluetooth Network    The forgotten interface

fabriq(config)# interface イーサネット
fabriq(config-if)# ip address ?
  from-hostlist        Use values from hostlist.csv (recommended)
  192.168.1.100        [hostlist.csv #1] IP for NEW-PC-01
  <ip-address>         Manual IP entry

fabriq(config-if)# ip address from-hostlist
*Apr 29 14:23:05.789: %FABRIQ-6-IPADDR: Address 192.168.1.100/24 has been whispered to interface イーサネット.
*Apr 29 14:23:05.890: %FABRIQ-6-IPADDR: Gateway 192.168.1.1 stands at the threshold.
*Apr 29 14:23:06.012: %FABRIQ-6-IPADDR: DNS servers 192.168.1.10, 192.168.1.11 are now the oracles.

fabriq(config-if)# end
fabriq# exit
The séance ends. The machine returns to its silent dream.
```

## 実装上の重要判断ポイント

### 1. Fabriq モジュールとの統合方法

二択:
- **A案**: 既存の `Invoke-BatchExecution` を呼ぶラッパー方式（CSVを動的生成して食わせる）
- **B案**: 各モジュールの core 部分（例: `hostname_config_core.ps1`）を直接呼ぶ

**推奨**: A案。既存の正規ルートを使う方が evidence 出力やログが揃って便利。

### 2. パスフレーズ管理

`enable` で要求するパスフレーズは、Fabriq の DPAPI 暗号化済みパスフレーズと同じものを使う。
`Unprotect-FabriqValue` を流用。

### 3. 日本語インターフェース名の扱い

`Get-NetAdapter` が返す `InterfaceAlias` をそのまま使う。
タブ補完で `イーサネット` が候補に出るのは仕様（むしろ味）。

### 4. 補完エンジン選定

PSReadLine の `Set-PSReadLineKeyHandler` で `Tab` と `?` をフックする方式。
独自REPLは作らず、`Read-Host` ベースのループにPSReadLineの補完を被せる。

### 5. v0.1 で実装しないもの

- `running-config` の双方向編集（read-onlyのみ）
- ネットワーク機器との実際のSSH/Telnet接続風機能
- 複数ホスト同時設定
- 他のモジュール（bitlocker, bloatware等）の取り込み（v0.2以降の楽しみとして残す）

## 受け入れ基準（Definition of Done for v0.1）

- [ ] PowerShell 5.1 環境で `fabriq_ios.ps1` 起動時にバナー表示
- [ ] User EXEC → Privileged EXEC → Global Config → Interface Config の4階層モード遷移が動作
- [ ] `enable` で Fabriq の暗号化パスフレーズ検証が通る
- [ ] `show host` / `show hosts` / `show version` / `show manifesto` が動作
- [ ] `hostname ?` でhostlist.csvから候補が出る
- [ ] `hostname <NewName>` で実際にホスト名変更が走る
- [ ] `interface ?` でNIC一覧が出る
- [ ] `interface <name>` でモード遷移
- [ ] `ip address from-hostlist` で hostlist記載のIP一式が反映
- [ ] 全成功時に Cisco 風syslog（シュルレアリスム文）が表示される
- [ ] タブ補完が各モードで効く
- [ ] `exit` / `end` の階層的振る舞いが正しい
- [ ] `syslog_messages.csv` を編集すると出力メッセージが変わる

## 将来構想（v0.2以降のメモ、実装しない）

- `copy running-config startup-config` で profile CSV 書き出し
- `show evidence detail` で evidence_config の詳細展開
- `reload` で実際の `__RESTART__` 発火
- `terminal monitor` で他端末のFabriq実行ログを監視
- LLM連携（自然言語 → FQL生成）
- bitlocker, bloatware_remove, odt_config 等の取り込み
- `show running-config` をペーストで他端末に流し込む再現機能

## 配布形態についての検討事項

- **同梱案**: `apps/fabriq_ios/` として Fabriq 本体に同梱
  - メリット: 一緒に試せる、kernel API変更時の追従が容易
  - デメリット: 「真面目なフレームワーク」の中に冗談が混ざる
- **分離案**: `fabriq-ios` として独立リポジトリ化、Fabriq 本体をサブモジュール参照
  - メリット: 「精神的続編」として独立した存在感
  - デメリット: バージョン整合の管理コスト

**初期推奨**: 同梱で開発開始 → 安定したら分離検討。

## 参考実装パターン

PowerShellでCisco風CLIを書いた前例として:
- `Posh-SSH` のインタラクティブモード
- `PSReadLine` のサンプル補完ハンドラ

これらを参考にしつつ、Fabriqの `Initialize-ModuleSystem` 戻り値を補完候補ソースとして再利用する方式が最も実装コストが低い。
