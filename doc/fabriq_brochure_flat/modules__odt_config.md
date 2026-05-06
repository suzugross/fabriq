# odt_config (Standard)

**カテゴリ**: Applications
**メニュー名**: ODT Install
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（setup.exe ExitCode のみが判定根拠）
**サブスクリプト**: `odt_install.ps1`（メイン処理 1 本。`assets/setup.exe` + 製品別 `assets/<dir>/configuration.xml` を同梱）

## 目的
Office Deployment Tool 経由で Microsoft 365 / Visio / Project 等を導入するモジュール。
エントリ単位で Offline（事前ダウンロード済み Office\）と Online（CDN ダウンロード）を
切り替えられるよう、構成 XML の `<Add SourcePath>` 属性を実行時に書き換える設計です。
既存 Click-to-Run Office を検出すると共存不可のため Error で中止。
ODT ログは `evidence/odt_log/` に hostname プレフィクスで自動収集されます。

## 入力 (CSV)
`odt_list.csv`:
- `Enabled`: 有効フラグ
- `XmlFileName`: ODT 構成 XML のファイル名（AssetsFolder 配下で解決）
- `Description`: 表示用
- `AssetsFolder`: エントリ固有 assets フォルダ（相対は `$PSScriptRoot` 基準、空欄時は `assets\`）
- `Mode`: `Offline`（既定）/ `Online`
- `Segment`

## 主要ステップ
1. `odt_list.csv` 読み込み（Enabled=1 のみ）
2. ドライラン: 各エントリの XML / AssetsFolder 存在確認、`assets\setup.exe` 必須
3. Step 3.5: 環境事前チェック + クリーンアップ
   - Office 系プロセス（WINWORD, EXCEL, OUTLOOK, ONENOTE, MSPUB, MSACCESS,
     VISIO, LYNC, Teams, OfficeClickToRun, OfficeC2RClient）強制終了
   - `ClickToRunSvc` 停止
   - ストア版 Office AppX（OneNote, Office.Desktop 等）を AppX + Provisioned 双方から削除
   - `msiserver` が Disabled なら Manual に昇格
   - System ドライブ空き容量確認（10GB 未満は Warning）
   - 既存 C2R Office 検出 → あれば Error 中止
4. 実行確認（AutoPilot は自動 Y）
5. エントリループ:
   - 構成 XML を `[xml]` で読み込み、Mode に応じて `SourcePath` 書換え/除去
   - 一時 XML を `%TEMP%` に保存し `setup.exe /configure "<temp_xml>"` を `-Wait` 同期実行
   - finally で一時 XML 削除 + ODT ログ収集（`C:\Windows\Temp\{COMPUTERNAME}-*.log` を
     セッション ts プレフィクス付きで `evidence/odt_log/` にコピー）
6. 集計返却（全成功=Success, 一部失敗=Partial, 全 XML 不在=Skipped, それ以外=Error）

## 注意点・運用メモ
- 既存 C2R Office があると Error 中止。SaRA で事前アンインストール必須
- Offline モードは AssetsFolder 配下に `Office\` オフラインソースが必須
  （本モジュールは `setup.exe /download` 機能を持たない）
- Online モードは Microsoft CDN へ HTTPS 疎通必須、数 GB 単位の通信が発生し
  setup.exe は Wait 同期のため数十分要することあり
- ストア版 Office を Cleanup フェーズで自動削除する点に注意（ユーザーデータ影響を事前確認）
- Acrobat や独自 EXE は `generic_process_runner` で別途実行する設計分離

## 検証
Post-Apply Verification は未実装。`setup.exe /configure` の ExitCode のみが成否判定で、
`New-ModuleResult` に `-Verified` は渡さない。アクティベーション状態は
別モジュール `office_license_config` の `/dstatus` 解析で別途検証する設計。
詳細トラブルシュートは `evidence/odt_log/` に自動収集された ODT ログで実施。
