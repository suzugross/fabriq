# SELECTED_* / FABRIQ_* 環境変数契約

`KERNEL_API.md §3` で公式宣言。fabriq の最重要 IPC（プロセス内コミュニケーション）。`hostlist.csv` の選択行から流れた値が、すべてのモジュールに共通して見える形で配信される。

---

## §3.1 ホスト情報変数（hostlist.csv 由来）

`Set-SelectedHostEnvironment` がホスト選択時点で `$env:` に流し込み。`ENC:` 暗号化フィールドは **この時点で透過復号** される。

### 識別系

| 環境変数 | hostlist.csv 列 | 用途 |
|---|---|---|
| `SELECTED_KANRI_NO` | `AdminID` | 管理 ID（実行履歴の一級識別子、Restore-ExecutionHistory のフィルタキー、HTML チェックリスト Meta） |
| `SELECTED_OLD_PCNAME` | `OldPCName` | 旧 PC 名（リネーム前の参照用） |
| `SELECTED_NEW_PCNAME` | `NewPCName` | 新 PC 名（hostname_config が適用先、エビデンスベースパス命名 / HTML / status.json で常時参照） |

### ネットワーク系

| 環境変数 | hostlist.csv 列 | 用途 |
|---|---|---|
| `SELECTED_ETH_IP` | `EthernetIP` | イーサネット IPv4 |
| `SELECTED_ETH_SUBNET` | `EthernetSubnet` | イーサネットサブネットマスク |
| `SELECTED_ETH_GATEWAY` | `EthernetGateway` | イーサネットゲートウェイ |
| `SELECTED_WIFI_IP` | `WifiIP` | Wi-Fi IPv4 |
| `SELECTED_WIFI_SUBNET` | `WifiSubnet` | Wi-Fi サブネットマスク |
| `SELECTED_WIFI_GATEWAY` | `WifiGateway` | Wi-Fi ゲートウェイ |
| `SELECTED_DNS1` 〜 `SELECTED_DNS4` | `DNS1`..`DNS4` | DNS サーバ最大 4 件 |

ipaddress_config / network_profile_config / DNS 系モジュールが消費。

### セキュリティ系

| 環境変数 | hostlist.csv 列 | 用途 |
|---|---|---|
| `SELECTED_PIN` | `Pin` | セットアップ時の PIN（cert_config / sysprep 等で参照、`ENC:` 暗号化推奨） |

### プリンタ系（10 スロット）

| 環境変数 | hostlist.csv 列 | 用途 |
|---|---|---|
| `SELECTED_PRINTER_<N>_NAME` | `Printer<N>Name` | プリンタ名（N=1..10） |
| `SELECTED_PRINTER_<N>_DRIVER` | `Printer<N>Driver` | ドライバ名 |
| `SELECTED_PRINTER_<N>_PORT` | `Printer<N>Port` | ポート（`IP_192.168.1.50` / `192.168.1.50` / TCPIP_*）|

`printer_driver_config` の register subroutine が消費。HTML チェックリストの Printer Cross-Check では Expected vs Actual で照合。

---

## §3.2 プロファイル実行パラメータ

| 環境変数 | 由来 | 寿命 |
|---|---|---|
| `FABRIQ_SEGMENT` | profile CSV の `Segment` 列 | モジュール 1 件の実行スコープ（Invoke-BatchExecution が前後で setup/teardown） |
| `FABRIQ_AUTOLOGON_USER` | `__AUTO_to_<User>__` マーカー | 同上（モジュール実行直前に立て、終了後に `$null`） |
| `FABRIQ_WORKER_NAME` | session.json `WorkerName` | セッション全体（Initialize-Session で設定） |
| `FABRIQ_EVIDENCE_BASE` | `Initialize-EvidenceBasePath` | セッション全体（`$global:FabriqEvidenceBasePath` と同値、モジュール内 `Join-Path` 用） |

### `FABRIQ_SEGMENT` の動作詳細

profile CSV の `Segment=office` 行から呼ばれた場合：

```
Invoke-BatchExecution の loop:
   1. $env:FABRIQ_SEGMENT = "office"
   2. & $module.Script   ← この間中 $env:FABRIQ_SEGMENT が見える
   3. $env:FABRIQ_SEGMENT = $null
```

モジュール側 `Import-ModuleCsv` のデフォルト引数 `-Segment $env:FABRIQ_SEGMENT` で自動的にフィルタが効く。

### `FABRIQ_AUTOLOGON_USER` の動作詳細

`__AUTO_to_admin01__` マーカーから呼ばれた場合：

```
Invoke-BatchExecution の loop:
   1. $env:FABRIQ_AUTOLOGON_USER = "admin01"
   2. & $autologon_config.ps1   ← 内部で $env:FABRIQ_AUTOLOGON_USER を読み、autologon_list.csv から User=admin01 行を選択
   3. $env:FABRIQ_AUTOLOGON_USER = $null
```

---

## §2 公開グローバル変数（補足）

環境変数とは別に、`$global:` スコープの公開変数も存在：

| グローバル | 型 | 環境変数版 | 違い |
|---|---|---|---|
| `$global:FabriqMasterPassphrase` | string | (なし) | 環境変数化しないことで child process 漏洩を防ぐ。Runspace 注入時のみ明示転送 |
| `$global:AutoPilotMode` | bool | (なし) | プロセス内 flag（環境変数化すると spawn された script や子プロセスにも影響して混乱） |
| `$global:AutoPilotWaitSec` | int | (なし) | 同上 |
| `$global:AutoConfirmMode` | bool | (なし) | FlexProfile 単発実行用 (kernel 3.1.0+) |
| `$global:FabriqEvidenceBasePath` | string | `FABRIQ_EVIDENCE_BASE` | 環境変数版も提供（モジュール内 `Join-Path` のため） |

---

## 環境変数のライフサイクル

### セッション全体

```
ホスト選択（Show-SessionSetupForm or resume の Restore-HostEnvironment）
   ↓
Set-SelectedHostEnvironment → 全 SELECTED_* 設定
   ↓
Initialize-Session → FABRIQ_WORKER_NAME 設定
   ↓
Initialize-EvidenceBasePath → FABRIQ_EVIDENCE_BASE 設定
   ↓
... モジュール実行中、SELECTED_* / FABRIQ_WORKER_NAME / FABRIQ_EVIDENCE_BASE は不変 ...
   ↓
Reset-FabriqState（New Session / Refabriq）
   ↓
すべての SELECTED_* / FABRIQ_AUTOLOGON_USER / FABRIQ_EVIDENCE_BASE を null 化
```

### モジュール 1 件のスコープ

```
Invoke-BatchExecution の foreach:
   1. _AutoLogonUser があれば $env:FABRIQ_AUTOLOGON_USER を立てる
   2. _Segment があれば $env:FABRIQ_SEGMENT を立てる
   3. & $module.Script
   4. $env:FABRIQ_AUTOLOGON_USER = $null（クリア）
   5. $env:FABRIQ_SEGMENT = $null
```

### 再起動跨ぎ

```
__RESTART__ 検出
   ↓
Save-ResumeState
   ├── 全 SELECTED_* を hash table 化（HostEnvironment fields）
   ├── EvidenceBasePath を json 保存
   └── ProtectedPassphrase（DPAPI 暗号化）も保存
   ↓ 再起動 ↓
   ↓
main.ps1 が Restore-HostEnvironment で全 SELECTED_* を json から戻す
   ↓
EvidenceBasePath / SessionID / ProtectedPassphrase（DPAPI 復号 → $global:FabriqMasterPassphrase）も復元
```

---

## 環境変数の Runspace 継承

`__ASYNC__` で Runspace 実行に切り替わった場合：

- `$env:*` は Process スコープなので **Runspace に自動継承される**（child runspace が parent process と同じ environment block を共有）
- `$global:*` は Runspace のスコープが異なるため **明示注入が必要**

`Invoke-SafeCommandAsync` の `$inject` ハッシュテーブルが `$global:*` を Runspace の global スコープへ Set-Variable で注入する。

---

## 環境変数の運用ルール

### 1. モジュールは SELECTED_* を読み取り専用扱いとする

```powershell
$pcName = $env:SELECTED_NEW_PCNAME   # OK: 読み取り
$env:SELECTED_NEW_PCNAME = "..."     # NG: モジュールから書き換えてはならない
```

理由: 同セッション内の他モジュールへ漏れて副作用となる。書き換えはカーネルの `Set-SelectedHostEnvironment` / `Restore-HostEnvironment` のみ許容。

### 2. ENC: 復号は kernel 経由で透過完了している

モジュールは `Unprotect-FabriqValue` を直接呼ぶ必要がない（hostlist.csv 由来は env 配信時、`_list.csv` 由来は `Import-ModuleCsv` で復号される）。

### 3. 機密値の env 配信を最小限に保つ

`SELECTED_PIN` / `SELECTED_PRINTER_<N>_*` 等の機密項目は env に出すが、**子プロセスを spawn する際は意識的に環境変数を絞る**（明示的な `Start-Process` 呼び出し時の Environment 引数で）。

### 4. fabriq_ios サブシェルからの参照

`apps/fabriq_ios` の `show ip interface` 等のコマンドは `$env:SELECTED_ETH_IP` 等を参照する。fabriq_ios も SELECTED_* 環境変数を「IOS のシステム情報」として表示する設計のため、契約破壊は fabriq_ios の挙動も壊す。
