# Fabriq ver2.2

**Manifeste du Surkitinisme**

Windows PC キッティング自動化フレームワーク（PowerShell + WinForms GUI）

## 概要

Fabriq は、Windows 11 PC の初期セットアップ（キッティング）を自動化する PowerShell ベースのフレームワークです。

- **CSV 駆動 + モジュール型アーキテクチャ** — コード改修なしに対象 PC・設定値・実行順序を切り替え
- **GUI ダッシュボード**（WinForms）から個別モジュール実行・プロファイル一括実行・セッション切替を操作
- **AutoPilot** による完全無人キッティング。モジュール単位のエラー処理ポリシー（`skip` / `retry`）、再起動跨ぎ自動再開、HTML チェックリスト出力、スクリーンショット自動取得
- **AES-256-CBC + PBKDF2** による CSV 内機密値の暗号化保持
- **70 種類超の標準・拡張モジュール**を同梱

## 前提条件

- **Windows 11**
- **PowerShell 5.1** 以降
- **管理者権限**（`Fabriq.exe` が自動昇格）
- **Fabriq Studio** によるマスターパスフレーズ設定（後述）

> **重要**: Fabriq の起動時にはマスターパスフレーズの入力が必須です。パスフレーズ検証トークン（`kernel/txt/passphrase_verify.txt`）が存在しない場合、Fabriq は起動できません。このトークンは **Fabriq Studio** でパスフレーズを設定することで生成されます。必ず初回起動前に Fabriq Studio でパスフレーズを設定してください。

## 主な機能

| 機能 | 説明 |
|------|------|
| **モジュールシステム** | Standard 57 種、Extended 14 種、計 71 種以上のモジュール（ホスト名、IP、レジストリ、アプリ、BitLocker、Sysprep 等） |
| **GUI ダッシュボード** | `Fabriq.exe` 起動後、WinForms ダッシュボードから全操作を実施。CLI モードは廃止 |
| **プロファイル実行** | 複数モジュールを順序付きで一括実行。`__AUTOPILOT__` マーカー以降は完全自動化 |
| **AutoPilot ErrorMode** | プロファイル CSV の `ErrorMode` 列でモジュール単位に `skip` / `retry`（最大 5 回）を宣言し、AutoPilot 中のエラー対応を自動化 |
| **再起動跨ぎ** | `__RESTART__` マーカーにより、再起動後に RunOnce 経由で自動再開 |
| **`__AUTO_to_<User>__`** | `autologon_config` の `autologon_list.csv` から `User` 列一致のエントリを指定して呼び出す専用マーカー（例: `__AUTO_to_admin01__`） |
| **CSV 駆動** | `hostlist.csv` で PC 毎の設定を定義。各モジュールの設定も CSV で管理 |
| **暗号化** | AES-256-CBC + PBKDF2-SHA256（10 万回、固定ソルト）で CSV 中の機密値を `ENC:<Base64>` 形式で保持 |
| **エビデンス自動取得** | モジュール実行ごとにスクリーンショットを自動保存 |
| **HTML チェックリスト** | プロファイル実行後に実行結果・ネットワーク照合・プリンタ照合を HTML レポート出力 |
| **ステータスモニタ** | 別ウィンドウで実行状況・PC 情報比較をリアルタイム表示 |
| **セグメント** | 同じモジュールをプロファイル内で設定値別に呼び分け可能（厳密マッチ） |
| **Post-Apply Verification** | 設定適用後に読み返し検証を行い、実行履歴に `Verified` 列として記録 |
| **プリセット UI** | 各モジュールの `preset.csv` により、設定 CSV の列挙型列を Fabriq Studio 側でドロップダウン編集可能 |
| **ログ管理** | PowerShell Transcript + 実行履歴 CSV + `log_destinations.csv` による外部共有フォルダへの自動アップロード |

## ディレクトリ構成

```
fabriq/
├── Fabriq.exe              # エントリーポイント（管理者自動昇格、GUI 起動）
├── Deploy.bat              # USB から対象 PC へのデプロイツール
├── kernel/
│   ├── main.ps1            # メインスクリプト（Fabriq.exe から呼び出し）
│   ├── common.ps1          # 共通関数ライブラリ（90+ 関数）
│   ├── csv/                # マスタ CSV（categories, hostlist, workers, log_destinations, manifesto）
│   ├── json/               # ランタイム状態（session, status, resume_state, art_pulse）
│   ├── ps1/                # カーネルサブスクリプト（manifesto, status_monitor, view_report, art_display）
│   └── txt/                # パスフレーズ検証トークン、アート文言、silence フラグ
├── modules/
│   ├── standard/           # 標準モジュール群（57）
│   └── extended/           # 拡張モジュール群（14）
├── profiles/               # 実行プロファイル CSV
├── apps/                   # FabriqApps：GUI アプリツール群
│   ├── fabriq_operator/    # メインダッシュボード GUI
│   ├── csv_editor/         # CSV 編集 GUI
│   ├── system_launcher/    # OS 機能ランチャ
│   ├── bloatware_exporter/ # インストール済みアプリ一覧エクスポート
│   ├── desktop_icon_backup_app/
│   ├── local_user_setup/
│   ├── storeapp_editor/
│   └── winget_gui/
├── commands/               # ユーティリティコマンド（diag_crypto, get_evidence, gpupdate 等）
├── evidence/               # エビデンス（スクリーンショット + HTML チェックリスト）出力先
├── logs/                   # ログ出力先（Transcript + 実行履歴 CSV）
└── dev/                    # 開発用（template, ico, launcher, cert_config_test, odt_config 等）
```

## クイックスタート

### 1. Fabriq Studio でパスフレーズを設定

Fabriq Studio を起動し、ワークスペースとして Fabriq フォルダを開き、マスターパスフレーズを設定します。検証トークン `kernel/txt/passphrase_verify.txt` が生成され、Fabriq が起動可能になります。

### 2. デプロイ

`Deploy.bat` を実行して USB メモリから対象 PC へ Fabriq フォルダをコピーします（フォルダを直接配置しても可）。

### 3. 起動

`Fabriq.exe` をダブルクリックします（UAC により管理者権限へ自動昇格 → PowerShell コンソールを経由してダッシュボード GUI が立ち上がります）。

### 4. セッション開始

1. マスターパスフレーズを入力
2. 作業者を選択（`workers.csv` から選択、または手入力）
3. 対象 PC を選択（`hostlist.csv` から選択。PC 名が一致すれば自動選択）

### 5. モジュール実行

ダッシュボード GUI の各ボタンから実行します。

## 使い方

### GUI ダッシュボード

ダッシュボードは複数タブで構成され、代表的なボタンは以下の通りです。

| ボタン | 動作 |
|---|---|
| **Execute Profile** | プロファイル CSV を選択して一括実行 |
| **View Details** | 選択中プロファイルの構成モジュールを事前確認 |
| **Execute** | 個別モジュールをチェック選択して連続実行 |
| **Open Folder**（Evidence） | エビデンス保存先をエクスプローラで開く |
| **Open CSV Editor** | 付属 CSV エディタで hostlist / モジュール設定を編集 |
| **Export History** | 実行履歴をエビデンスとして明示エクスポート |
| **Regenerate Checklist** | 実行履歴から HTML チェックリストを再生成 |
| **Windows Update** | Windows Update を再起動跨ぎで実行 |
| **Restart PC** | 再起動（fabriq を継続するかは確認） |
| **Refabriq** | セッション・履歴・トランスクリプトをリセットして再開 |
| **System Launcher** | Windows の設定系ショートカット集 |
| **FabriqApps** | `apps/` 配下の GUI ツール群を起動 |
| **New Session** | 作業者・対象 PC を切り替えて新セッション開始 |
| **Manifeste du Surkitinisme** | マニフェスト表示 |

### プロファイル実行

プロファイル CSV で実行するモジュールと順序を定義し、AutoPilot モードで完全自動実行が可能です。

**プロファイル CSV の書式:**

```csv
Order,ScriptPath,Enabled,Description,Segment,ErrorMode
10,__AUTOPILOT__,1,WaitSec=3,,
20,standard/hostname_config/hostname_config.ps1,1,ホスト名設定,,
30,standard/ipaddress_config/ipaddress_config.ps1,1,IP アドレス設定,,retry
40,__RESTART__,1,再起動,,
50,standard/reg_hklm_config/reg_hklm_config.ps1,1,レジストリ設定,,skip
```

| 列 | 説明 |
|---|---|
| `Order` | 実行順（整数、昇順） |
| `ScriptPath` | `standard/<module>/<script>.ps1` 形式。区切りは `/` / `\` どちらも可 |
| `Enabled` | `1`=実行 / `0`=スキップ |
| `Description` | プロファイル UI 表示用コメント |
| `Segment` | 同モジュールを設定値別に呼び分ける際のセグメント名（省略可。`_list.csv` 側の `Segment` 列と厳密マッチ） |
| `ErrorMode` | AutoPilot 実行時のエラー処理ポリシー（省略=ダイアログ確認 / `skip`=自動スキップ / `retry`=最大 5 回自動リトライ） |

**特殊マーカー:**

| マーカー | 動作 |
|---|---|
| `__AUTOPILOT__` | 以降のプロファイル実行を AutoPilot モードで自動化。`Description` に `WaitSec=N` でモジュール間ウェイト秒を指定可能 |
| `__ASYNC__` | 以降のモジュールを監視付き Runspace で実行。ハング時に Status Monitor の **Skip** ボタン、または `async_config.json` の `DefaultTimeoutSec` で強制スキップ可能（`Enabled: false` で全体無効化） |
| `__RESTART__` | Windows を再起動し、RunOnce 経由で次モジュールから自動再開 |
| `__REEXPLORER__` | Explorer を再起動（レジストリ変更の即時反映等） |
| `__AUTO_to_<User>__` | `autologon_config` の `autologon_list.csv` から `User` 列一致のエントリを呼び出す（例: `__AUTO_to_admin01__`） |

## モジュール一覧

### Standard モジュール（57）

| カテゴリ | モジュール |
|---|---|
| **Network** | `hostname_config`, `ipaddress_config`, `domain_join`, `ssid_config` |
| **Display** | `brightness_config`, `dpi_api_config`, `resolution_api_config` |
| **Desktop** | `wallpaper_config`, `taskbar_config`, `startlayout_config` |
| **Security** | `bitlocker_config`, `firewall_config`, `cert_config`, `office_license_config`, `windows_license_config` |
| **User Management** | `local_user_config`, `profile_delete` |
| **Printer** | `printer_driver_config`, `printer_delete` |
| **Applications** | `app_config`, `winget_install`, `bloatware_remove`, `bloatware_export`, `storeapp_config`, `odt_config`, `browser_addon_config`, `fabriq_app_launcher` |
| **Power** | `power_config` |
| **Maintenance** | `acl_config`, `copyfile_config`, `file_delete`, `office_update`, `partition_config`, `robocopy_config`, `system_finalize` |
| **System** | `autologon_config`, `default_app_config`, `driver_config`, `generic_process_runner`, `ppkg_config`, `process_killer`, `restart_config`, `restore_point`, `scheduled_task_config`, `signout_config`, `spi_config`, `sysprep_config`, `time_sync_config`, `volume_config` |
| **Registry** | `reg_hklm_config`, `reg_hkcu_config` |
| **Scripts** | `generic_batch_runner`, `startup_command_config` |
| **Evidence** | `evidence_config` |
| **Test** | `test_error_module`, `test_harness_config` |

`windows_update` は GUI ダッシュボードの **Windows Update** ボタン専用で、`module.csv` を持たず Script Menu には表示されません。

### Extended モジュール（14）

| カテゴリ | モジュール |
|---|---|
| **Network** | `ipv6_config`, `network_profile_config` |
| **Display** | `display_config`, `dpi_config` |
| **Desktop** | `desktop_icon_config` |
| **User Management** | `builtin_admin_config`, `group_config` |
| **Maintenance** | `directory_cleaner`, `history_destroyer` |
| **System** | `azure_ad_join_check`, `reg_template` |
| **Scripts** | `script_looper` |
| **ManualWorks** | `manual_kitting_assistant` |
| **Evidence** | `log_uploader` |

## モジュール構成

各モジュールは以下のファイルで構成されます。

| ファイル | 役割 |
|---|---|
| `module.csv` | メニュー名・カテゴリ・表示順・有効無効（1 モジュール内に複数エントリ可） |
| `<name>.ps1` | 実行スクリプト本体。`dev/template/_template_script.ps1` をベースに実装 |
| `<name>_list.csv` | 設定データ（対象リスト等） |
| `Guide.txt` | 使い方ガイド（日本語） |
| `preset.csv` | 設定 CSV のドロップダウン UI 定義（Fabriq Studio が検出。列挙型列に候補を提示） |

### 結果ステータス

全モジュールは `New-ModuleResult` または `New-BatchResult` を通じて統一ステータスを返却します。

| ステータス | 意味 |
|---|---|
| `Success` | 正常完了 |
| `Partial` | 一部成功・一部失敗 |
| `Skipped` | 全件スキップ（対象が無い等） |
| `Cancelled` | ユーザーがキャンセル |
| `Error` | エラー発生 |

Post-Apply Verification を実装したモジュールは、`-Verified $true/$false` で実行履歴に検証結果を記録します。

### preset.csv の書式

Fabriq Studio のモジュール編集画面は、`preset.csv` を検出すると該当列を編集可能なドロップダウン（ComboBox）UI に切り替えます。

```csv
Column,Value,Label
Enabled,1,有効
Enabled,0,無効
Type,REG_DWORD,32 ビット整数 (DWORD)
Type,REG_SZ,文字列 (SZ)
```

| 列 | 意味 |
|---|---|
| `Column` | 対象 CSV の列名（`module.csv` 以外の CSV ヘッダと一致） |
| `Value` | セルに実際に書き込まれる値（fabriq PowerShell が受け付ける文字列そのまま） |
| `Label` | 表示用ラベル |

エンコーディングは **UTF-8 BOM**、改行は **CRLF**。

## カスタマイズ

### hostlist.csv

対象 PC ごとの設定を定義します。

```csv
AdminID,OldPCName,NewPCName,EthernetIP,EthernetSubnet,EthernetGateway,...,Printer1Name,Printer1Driver,Printer1Port,...
1,OLD-PC-01,NEW-PC-01,192.168.1.100,255.255.255.0,192.168.1.1,...
```

機密性のあるフィールドは Fabriq Studio で `ENC:<Base64>` 形式に暗号化できます。実行時は `Import-ModuleCsv` がマスターパスフレーズで透過的に復号します。

### 新規モジュール作成

1. `dev/template/` フォルダを `modules/standard/` または `modules/extended/` にコピー
2. フォルダ名をモジュール名にリネーム（snake_case、例: `my_new_config`）
3. `module.csv`、実行スクリプト、設定 CSV を編集
4. 必要に応じて `preset.csv` を追加
5. `kernel/common.ps1` の共通関数（`Show-Info` / `New-BatchResult` / `Import-ModuleCsv` 等）を必ず使用
6. 可能な限り Post-Apply Verification（Step 5.5）を実装

開発時のガイドラインは [CLAUDE.md](CLAUDE.md) を参照してください。

### 暗号化仕様

CSV 中の機密値は `ENC:<Base64>` 形式で保持されます。

- **鍵導出**: PBKDF2-HMAC-SHA256、100,000 回、固定ソルト `fabriq-fixed-salt-2024`
- **暗号化**: AES-256-CBC、PKCS7 パディング
- **エンコード**: UTF-8（平文）、Base64（暗号文）

Fabriq Studio でフィールド単位・行単位・列単位で暗号化／復号できます。実行時にはマスターパスフレーズにより自動復号されます。

### ログ配送

`kernel/csv/log_destinations.csv` に宛先（UNC 共有フォルダ / ローカルパス）と認証情報を登録すると、`extended/log_uploader` モジュールが Transcript・実行履歴・エビデンスを外部へ自動アップロードします。

## Fabriq Studio との関係

Fabriq Studio（別プロジェクト。WPF / .NET 8.0）は Fabriq の **GUI 管理ツール** です。本体 Fabriq との接点は以下に限定されます。

| 機能 | 本体との関係 |
|---|---|
| パスフレーズ設定 | `kernel/txt/passphrase_verify.txt` を生成（**Fabriq 起動の必須条件**） |
| ホスト管理 | `kernel/csv/hostlist.csv` を編集、`ENC:` 暗号化／復号 |
| モジュール設定管理 | 各モジュールの `_list.csv` を編集。`preset.csv` を検出してドロップダウン UI を提供 |
| プロファイル管理 | `profiles/*.csv` の作成・編集 |
| レジストリカタログ | レジストリ設定のライブラリからワークスペースへエクスポート |

本体 Fabriq は Studio のバージョン・機能に依存しません。Studio 側の具体的な機能セットは Studio リポジトリを参照してください。

## ライセンス

[MIT License](LICENSE)

## Author

yuki.suzuki@suzugross.com
