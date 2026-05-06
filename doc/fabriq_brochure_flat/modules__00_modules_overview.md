# モジュール全体図 — 標準 60 / 拡張 15

fabriq の機能はすべて `modules/{standard,extended}/<name>/` に packaging されている。各モジュールは独立 SemVer で配備され、要求カーネル版を `REQUIRES_KERNEL` で宣言する。

カテゴリは `kernel/csv/categories.csv` で定義され、ダッシュボードのグルーピング順序を決定する。

詳細な per-module 解説は本ディレクトリの `<name>.md` 各ファイルを参照。

---

## Standard（60 件）

### Network

| モジュール | 主な役割 | Verification |
|---|---|---|
| `hostname_config` | PC 名変更（再起動で反映） | あり（pending value 検証） |
| `ipaddress_config` | 静的 IP / Subnet / Gateway / DNS 設定 | あり（読み返し） |
| `temp_ipaddress_config` | プールから一時 IP 取得（GUI 選択 + DAD） | あり |
| `domain_join` | ドメイン参加（`ENC:` パスワード） | 不可（再起動後反映） |
| `ssid_config` | Wi-Fi プロファイル追加（netsh wlan add profile） | あり |

### Display

| モジュール | 主な役割 | Verification |
|---|---|---|
| `brightness_config` | 輝度設定（WMI） | あり |
| `dpi_api_config` | DPI スケール（Win32 API） | 不可 |
| `resolution_api_config` | 解像度変更（ChangeDisplaySettings） | 部分 |

### Desktop

| モジュール | 主な役割 | Verification |
|---|---|---|
| `wallpaper_config` | 壁紙（SystemParametersInfo） | なし |
| `taskbar_config` | LayoutModification.xml を Default User へ | なし |
| `startlayout_config` | スタートメニュー Backup/Build/Import/Delete pair | 部分 |

### Security

| モジュール | 主な役割 | Verification |
|---|---|---|
| `bitlocker_config` | BitLocker 有効化（async, await pair） | 不可（async 完了は別） |
| `firewall_config` | プロファイル別 ON/OFF + ルール設定 | あり |
| `firewall_rule_config` | ルール export / import pair | あり（両方） |
| `firewall_rule_make_config` | ルール手動定義 | あり |
| `cert_config` | 証明書ストア配置（`ENC:` パスワード） | あり |
| `office_license_config` | Office キー install + Activate pair | de-facto（auth が verify） |
| `windows_license_config` | Windows キー install + Activate pair | 部分（Activate のみ） |

### User Management

| モジュール | 主な役割 | Verification |
|---|---|---|
| `local_user_config` | ローカルユーザ作成 + 削除 pair | あり（作成のみ） |
| `profile_delete` | ユーザプロファイル削除 | なし（推奨実装あり、未実装） |

### Printer

| モジュール | 主な役割 | Verification |
|---|---|---|
| `printer_driver_config` | ドライバ install + register + uninstall trio | あり（install + register） |
| `printer_delete` | プリンタ削除（printer_list.csv 参照） | あり |

### Applications

| モジュール | 主な役割 | Verification |
|---|---|---|
| `app_config` | EXE/MSI/MSU インストール（共有 CSV） | 不可（複雑） |
| `winget_install` | winget update / install / upgrade trio | なし（winget 自身で完結） |
| `bloatware_remove` | UWP / Provisioned / Capability 削除 | あり |
| `bloatware_export` | 現状の bloatware 一覧 export | N/A（read-only） |
| `storeapp_config` | StoreApp 一括削除 | あり |
| `odt_config` | ODT で Office インストール | なし |
| `browser_addon_config` | Edge / Chrome 拡張機能 ADMX 経由 | あり |
| `fabriq_app_launcher` | apps/ ランチャ（FabriqApps ボタン本体） | N/A |

### Power

| モジュール | 主な役割 | Verification |
|---|---|---|
| `power_config` | 電源プラン（P/Invoke powrprof.dll、HP OEM 対策で Win32 API） | あり |

### Maintenance

| モジュール | 主な役割 | Verification |
|---|---|---|
| `acl_config` | ACL backup + restore pair | 除外（誤 PASS リスク） |
| `copyfile_config` | ファイル / ディレクトリコピー | 除外（誤 PASS リスク） |
| `file_delete` | ファイル / ディレクトリ削除 | あり |
| `office_update` | Office 更新（OffScrubC2R or click-to-run update） | de-facto |
| `partition_config` | パーティション作成 / 拡張 | あり（±5%） |
| `robocopy_config` | UNC / Local の robocopy（`ENC:` パスワード） | なし |
| `system_finalize` | shell32 reload + cache flush + Explorer 再起動 | なし |

### System

| モジュール | 主な役割 | Verification |
|---|---|---|
| `autologon_config` | AutoLogon registry 設定（`ENC:` パスワード）| 不可（再起動で反映） |
| `default_app_config` | DISM / SetUserFTA で既定アプリ設定 | 不可（新プロファイルで反映） |
| `driver_config` | DISM driver export + import pair | 不可 |
| `generic_process_runner` | EXE / MSI を generic に発火 | なし |
| `ppkg_config` | プロビジョニングパッケージ install + uninstall | なし |
| `process_killer` | プロセス強制終了（idempotent） | 意図的になし |
| `restart_config` | 再起動（AutoPilot 非対応） | 不可 |
| `restore_point` | システム復元ポイント作成 | 部分 |
| `scheduled_task_config` | タスクスケジューラ enable/disable pair（共有 CSV）| 部分 |
| `signout_config` | サインアウト | 不可 |
| `spi_config` | SystemParametersInfo + Active Setup（HKCU 系） | 除外 |
| `sysprep_config` | unattend.xml + SetupComplete.cmd 生成 | 除外 |
| `time_sync_config` | w32tm NTP 設定 + sync source verify | あり（retry 込） |
| `volume_config` | Core Audio API でマスターボリューム | あり |

### Registry

| モジュール | 主な役割 | Verification |
|---|---|---|
| `reg_hklm_config` | HKLM レジストリ + delete pair（Test-RegistryValueMatch 共有関数） | あり |
| `reg_hkcu_config` | HKCU + Default Profile + Active Setup/Startup Batch + delete pair | あり |

### Scripts

| モジュール | 主な役割 | Verification |
|---|---|---|
| `generic_batch_runner` | 任意 .bat / .ps1 / .cmd を generic 実行 | なし |
| `startup_command_config` | Default User Startup folder にコマンド配置 | なし |

### Evidence

| モジュール | 主な役割 | Verification |
|---|---|---|
| `evidence_config` | 22 セクションのシステム情報収集 + manifest.json 生成 | N/A |

### Test

| モジュール | 主な役割 | Verification |
|---|---|---|
| `test_error_module` | ErrorMode 検証用エラー発生器 | N/A |
| `test_harness_config` | マルチシナリオテストハーネス | あり（by design） |

### Standalone（Profile 直接登録不可）

| モジュール | 主な役割 |
|---|---|
| `windows_update` | スタンドアロン COM-API ループ。`module.csv` 無し、`Invoke-WindowsUpdateLoop` から `[wu]` 経由起動 |

---

## Extended（15 件）

### Network

| モジュール | 主な役割 | Verification |
|---|---|---|
| `ipv6_config` | IPv6 binding toggle（adapter ごと） | なし |
| `network_profile_config` | NetworkList Category 書き換え（baseline + override パターン） | あり |

### Display

| モジュール | 主な役割 | Verification |
|---|---|---|
| `display_config` | PrimSurfSize.cx/cy registry（マルチモニタ）| なし（再起動要） |
| `dpi_config` | per-monitor DPI（embedded C# DpiScaleResolver、HKCU + Default-hive dual-write）| なし |

### Desktop

| モジュール | 主な役割 | Verification |
|---|---|---|
| `desktop_icon_config` | Desktop アイコン .reg backup + restore pair（HKCU↔HKU\<SID> 正規化）| なし |

### User Management

| モジュール | 主な役割 | Verification |
|---|---|---|
| `builtin_admin_config` | Built-in Administrator アカウント設定（Enabled 列で Enable/Disable）| なし |
| `group_config` | ローカルグループメンバ管理（WMI で CurrentUser 解決）| なし |

### Maintenance

| モジュール | 主な役割 | Verification |
|---|---|---|
| `directory_cleaner` | ディレクトリクリーンアップ（hardcoded forbidden-path whitelist 3 重 guard）| なし |
| `history_destroyer` | 13 カテゴリ履歴削除（Edge / Chrome / Search / Wi-Fi / 7 special handlers）| なし |

### System

| モジュール | 主な役割 | Verification |
|---|---|---|
| `azure_ad_join_check` | dsregcmd 出力解析（read-only、script_looper retry 設計）| N/A |
| `reg_template` | レジストリ template backup / import pair（汎用テンプレ）| なし |

### Scripts

| モジュール | 主な役割 | Verification |
|---|---|---|
| `script_looper` | OnError / Always retry framework（pipeline + global 二重 ModuleResult 検出）| なし |

### ManualWorks

| モジュール | 主な役割 | Verification |
|---|---|---|
| `manual_kitting_assistant` | 手動作業アシスタント WinForms（Gundam Light theme、NoActivate window）| N/A（Enabled=0 デフォ） |

### Evidence

| モジュール | 主な役割 | Verification |
|---|---|---|
| `log_uploader` | robocopy で UNC / Local へエビデンス + ログ転送（log_destinations.csv 参照）| 不可（best-effort） |

### Special: Pianist (extended)

| モジュール | 主な役割 | Verification |
|---|---|---|
| `pianist` | **マルチフェーズ GUI maestro**。3-tab Phase view（Procedure / Samples / Values）+ section markers + modeless image viewer + Pause/Stop/Speed (v1.6.0+)。apps/ から extended/ へ昇格（2026-05-02）、2189 行の独立色濃いモジュール | あり（Manual phase status aggregation） |

---

## モジュール構成ファイル（共通）

各モジュールは以下のファイル群で構成される：

| ファイル | 役割 | overlay 区分 |
|---|---|---|
| `module.csv` | メニュー名・カテゴリ・表示順・有効無効（複数行可） | framework |
| `<name>.ps1` | 実行スクリプト本体（dev/template ベース） | framework |
| `<other>.ps1` | 補助スクリプト（_install / _uninstall / _backup / _restore 等） | framework |
| `<name>_list.csv` | 設定データ（対象リスト等） | site-specific |
| `Guide.txt` | 使い方ガイド（日本語） | framework |
| `preset.csv` | Studio 用ドロップダウン UI 定義（任意） | framework |
| `VERSION` | モジュール SemVer（1 行 X.Y.Z） | framework |
| `REQUIRES_KERNEL` | 要求最小カーネル版（1 行 X.Y.Z） | framework |

---

## カテゴリ統計

| カテゴリ | Standard | Extended |
|---|---|---|
| Network | 5 | 2 |
| Display | 3 | 2 |
| Desktop | 3 | 1 |
| Security | 7 | 0 |
| User Management | 2 | 2 |
| Printer | 2 | 0 |
| Applications | 8 | 0 |
| Power | 1 | 0 |
| Maintenance | 7 | 2 |
| System | 14 | 2 |
| Registry | 2 | 0 |
| Scripts | 2 | 1 |
| Evidence | 1 | 1 |
| Test | 2 | 0 |
| ManualWorks | 0 | 1 |
| Standalone | 1 | 0 |

---

## Verification 実装率

- **実装済み**: 25 / 75 ≈ 33%
- **意図的除外**（誤 PASS リスク or 検証不可能）: 約 15 件
- **未実装（推奨）**: 残り

検証除外リスト（feedback memory `project_verification_exclusions`）:
- `acl_config`: ACL ツリー完全読み返しは膨大、サブセット検証で false PASS の risk
- `spi_config`: Default Profile への hive load 経由でログイン後にしか反映されない
- `copyfile_config`: ファイル存在 != 内容正しい、ハッシュ検証は重い
- `sysprep_config` / `restart_config` / `signout_config` / `domain_join`: 再起動後 / OS 再起動後 / 別ユーザログオン後にしか確認できない（技術的不可）
