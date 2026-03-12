# Fabriq ver2.1

**Manifeste du Surkitinisme**

Windows PC キッティング自動化フレームワーク（PowerShell）

## 概要

Fabriq は、Windows PC の初期セットアップ（キッティング）を自動化する PowerShell ベースのフレームワークです。

- CSV 駆動のデータ定義 + モジュール型アーキテクチャにより、コード変更なしに設定を切り替え可能
- 単体実行・バッチ実行・プロファイル一括実行（AutoPilot）に対応
- 再起動をまたぐ自動継続実行、HTML チェックリスト出力、スクリーンエビデンス自動取得を内蔵

## 前提条件

- **Windows 11**
- **PowerShell 5.1** 以降
- **管理者権限**
- **Fabriq Studio** によるパスフレーズ設定（後述）

> **重要**: Fabriq の起動時にはマスターパスフレーズの入力が必須です。パスフレーズの検証トークン（`kernel/txt/passphrase_verify.txt`）が存在しない場合、Fabriq は起動できません。このトークンは **Fabriq Studio** でパスフレーズを設定することで生成されます。必ず初回起動前に Fabriq Studio でパスフレーズを設定してください。

## 主な機能

| 機能 | 説明 |
|------|------|
| **モジュールシステム** | 40 種類以上のモジュール（ホスト名、IP、レジストリ、アプリ、BitLocker、Sysprep 等） |
| **プロファイル実行** | 複数モジュールを順序付きで一括実行。AutoPilot モードで完全自動化 |
| **再起動跨ぎ** | `__RESTART__` マーカーにより、再起動後に RunOnce 経由で自動再開 |
| **CSV 駆動** | `hostlist.csv` で PC 毎の設定を定義。各モジュールの設定も CSV で管理 |
| **暗号化** | AES-256-CBC + PBKDF2 で CSV 中の機密値（パスワード、IP 等）を暗号化保持 |
| **エビデンス自動取得** | モジュール実行ごとにスクリーンショットを自動保存 |
| **HTML チェックリスト** | プロファイル実行後に実行結果・ネットワーク照合・プリンタ照合を HTML レポート出力 |
| **ステータスモニタ** | 別ウィンドウで実行状況・PC 情報比較をリアルタイム表示（WinForms GUI） |
| **セグメント** | 同じモジュールをプロファイル内で設定値別に呼び分け可能（厳密マッチ） |
| **ログ管理** | PowerShell Transcript + 実行履歴 CSV + 外部共有フォルダへの自動アップロード |

## ディレクトリ構成

```
fabriq/
├── Fabriq.bat              # エントリーポイント（管理者自動昇格）
├── Deploy.bat              # USB からPCへのデプロイツール
├── kernel/
│   ├── main.ps1            # メインスクリプト
│   ├── common.ps1          # 共通関数ライブラリ
│   ├── csv/                # マスタCSV（categories, hostlist, workers 等）
│   ├── json/               # ランタイム状態（session, status, resume_state）
│   ├── ps1/                # カーネルサブスクリプト（manifesto, status_monitor 等）
│   └── txt/                # パスフレーズ検証トークン等
├── modules/
│   ├── standard/           # 標準モジュール群
│   └── extended/           # 拡張モジュール群
├── profiles/               # 実行プロファイル（CSV 定義）
├── apps/                   # GUI アプリツール群
├── commands/               # ユーティリティコマンド
├── evidence/               # エビデンス出力先
├── logs/                   # ログ出力先
└── dev/                    # 開発用テンプレート・ツール
```

## クイックスタート

### 1. Fabriq Studio でパスフレーズを設定

Fabriq Studio を起動し、ワークスペースとして Fabriq フォルダを開き、マスターパスフレーズを設定します。これにより検証トークンが生成され、Fabriq が起動可能になります。

### 2. デプロイ

`Deploy.bat` を実行して USB メモリから対象 PC へ Fabriq フォルダをコピーします（フォルダを直接配置しても可）。

### 3. 起動

`Fabriq.bat` を実行します（管理者権限に自動昇格）。

### 4. セッション開始

1. マスターパスフレーズを入力
2. 作業者を選択（`workers.csv` から選択、または手入力）
3. 対象 PC を選択（`hostlist.csv` から選択。PC 名が一致すれば自動選択）

### 5. モジュール実行

- **メインメニュー** → `[S] Script Menu` でモジュールを個別実行
- **メインメニュー** → `[P] Run Profile` でプロファイルによる一括実行

## 使い方

### メインメニュー

```
[S] Script Menu       モジュール選択・実行
[A] FabriqApps        GUI アプリツール群
[C] Command           ユーティリティコマンド
[P] Run Profile       プロファイル一括実行
[N] New Session       セッション切替
[R] History           実行履歴表示
```

### Script Menu

カテゴリ別にモジュールが番号付きで表示されます。番号を入力して個別実行するほか、バッチ入力（`1,3,5` や `1-5`）で複数モジュールを連続実行できます。

### プロファイル実行

プロファイル CSV で実行するモジュールと順序を定義し、AutoPilot モードで完全自動実行が可能です。

**プロファイル CSV の例:**

```csv
Order,ScriptPath,Enabled,Description
10,standard\hostname_config\hostname_config.ps1,1,ホスト名設定
20,standard\ipaddress_config\ipaddress_config.ps1,1,IP アドレス設定
30,__RESTART__,1,再起動
40,standard\reg_hklm_config\reg_hklm_config.ps1,1,レジストリ設定
50,__SHUTDOWN__,1,シャットダウン
```

**特殊マーカー:**

| マーカー | 動作 |
|---------|------|
| `__RESTART__` | Windows を再起動し、RunOnce 経由で自動再開 |
| `__SHUTDOWN__` | Windows をシャットダウン |
| `__PAUSE__` | ユーザー入力待ちで一時停止 |
| `__REEXPLORER__` | Explorer を再起動（レジストリ変更の即時反映等） |
| `__STOPLOG__` / `__STARTLOG__` | トランスクリプトの停止・再開 |

## モジュール一覧

### Standard モジュール

| カテゴリ | モジュール |
|---------|-----------|
| Network | `hostname_config`, `ipaddress_config`, `domain_join` |
| Desktop | `wallpaper_config`, `taskbar_config` |
| Security | `bitlocker_config`, `firewall_config` |
| User Management | `local_user_config`, `autologon_config`, `profile_delete` |
| Printer | `printer_driver_config`, `printer_delete` |
| Applications | `app_config`, `winget_install`, `bloatware_remove`, `bloatware_export`, `storeapp_config`, `odt_config`, `fabriq_app_launcher` |
| Power | `power_config`, `brightness_config` |
| Maintenance | `copyfile_config`, `file_delete`, `process_killer`, `generic_batch_runner`, `generic_process_runner` |
| System | `sysprep_config`, `restart_config`, `signout_config`, `windows_license_config` |
| Display | `dpi_api_config`, `resolution_api_config` |
| Registry | `reg_hklm_config`, `reg_hkcu_config` |
| Evidence | `evidence_config` |

### Extended モジュール

| カテゴリ | モジュール |
|---------|-----------|
| Security | `azure_ad_join_check`, `builtin_admin_config` |
| Desktop | `desktop_icon_config` |
| Display | `display_config`, `dpi_config` |
| Applications | `edge_config`, `heif_config` |
| Network | `ipv6_config` |
| Maintenance | `directory_cleaner`, `history_destroyer` |
| Scripts | `script_looper` |
| Registry | `reg_template` |
| ManualWorks | `manual_kitting_assistant` |
| System | `log_uploader` |

## モジュール構成

各モジュールは以下のファイルで構成されます。

| ファイル | 役割 |
|---------|------|
| `module.csv` | メニュー名、カテゴリ、表示順、有効/無効 |
| `<name>.ps1` | 実行スクリプト本体 |
| `<name>_list.csv` | 設定データ（対象リスト等） |
| `Guide.txt` | 使い方ガイド |

全モジュールは `New-ModuleResult` で統一された結果ステータスを返却します。

| ステータス | 意味 |
|-----------|------|
| `Success` | 正常完了 |
| `Error` | エラー発生 |
| `Cancelled` | ユーザーがキャンセル |
| `Skipped` | スキップ |
| `Partial` | 一部成功・一部失敗 |

## カスタマイズ

### hostlist.csv

対象 PC ごとの設定を定義します。

```csv
AdminID,OldPCName,NewPCName,EthernetIP,EthernetSubnet,EthernetGateway,...,Printer1Name,Printer1Driver,Printer1Port,...
1,OLD-PC-01,NEW-PC-01,192.168.1.100,255.255.255.0,192.168.1.1,...
```

機密性のあるフィールドは Fabriq Studio で `ENC:<Base64>` 形式に暗号化できます。

### 新規モジュール作成

1. `dev/template/` フォルダを `modules/standard/` または `modules/extended/` にコピー
2. フォルダ名をモジュール名にリネーム（例: `my_new_config`）
3. `module.csv`、実行スクリプト、設定 CSV を編集

### 暗号化

CSV 中の機密値は `ENC:<Base64>` 形式で保持されます。Fabriq Studio でフィールド単位・行単位・列単位で暗号化/復号が可能です。実行時にはマスターパスフレーズにより AES-256-CBC（PBKDF2 鍵導出）で自動復号されます。

## Fabriq Studio との関係

Fabriq Studio（WPF / .NET 8.0）は Fabriq の **GUI 管理ツール** です。

| 機能 | 説明 |
|------|------|
| パスフレーズ設定 | 検証トークンの生成（**Fabriq 起動の必須条件**） |
| ホスト管理 | `hostlist.csv` の GUI 編集・暗号化 |
| モジュール管理 | モジュール設定 CSV の GUI 編集 |
| プロファイル管理 | プロファイル CSV の作成・編集 |
| モジュール生成 | Autokey Recipe / Script Looper / Digital Gyotaq モジュールの自動生成 |
| レジストリカタログ | レジストリ設定のライブラリ管理・ワークスペースへのエクスポート |

## ライセンス

[MIT License](LICENSE)

## Author

yuki.suzuki@suzugross.com
