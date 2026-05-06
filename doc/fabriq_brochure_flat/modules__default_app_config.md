# default_app_config (Standard)

**カテゴリ**: System
**メニュー名**: Export App Associations / Default App Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（新規ユーザープロファイル作成時に反映される仕様のため検証困難）
**サブスクリプト**:
- `default_app_config.ps1` … 本番キッティング用（`Dism /Import-DefaultAppAssociations`）
- `export_app_associations.ps1` … マスター準備用（`Dism /Export-DefaultAppAssociations`）

## 目的
Windows 10/11 で「PDF を Acrobat Reader で開く」「ブラウザは Chrome を既定にする」など、**既定アプリの関連付け** を ProgId ハッシュ保護に対応した安全な方法で配布するモジュールです。マスター PC で手動設定 → エクスポートで XML 化（`Export-DefaultAppAssociations`）、本番 PC で XML をインポート（`Import-DefaultAppAssociations`）する 2 段構成。インポートは「**新規ユーザープロファイル作成時** に適用」される Windows の仕様に基づいているため、キッティングで sysprep 前に書いておくのが定石です。Segment 列で「営業部用 / 技術部用」など部署別の関連付けセットを切り替え可能。

## 入力 (CSV)
`default_app_list.csv` の主な列:
- `Enabled` … 1=実行 / 0=スキップ
- `XmlFile` … `xml/` 配下の XML ファイル名
- `Description` … 表示ラベル
- `Segment` … セグメント名（部署／拠点別の使い分け）

## 主要ステップ（Import 側 = `default_app_config.ps1`）
1. `default_app_list.csv` 読み込み
2. Pre-flight（DISM 利用可否、`xml/` ディレクトリ）
3. Dry-run 表示
4. 実行確認（AutoPilot 時は自動 Y）
5. `Dism /Online /Import-DefaultAppAssociations:"<XML>"` を順次実行
6. 結果集計（`New-BatchResult`）

Export 側（`export_app_associations.ps1`）は同じスケルトンで `/Export-DefaultAppAssociations` を実行し、`xml/` に上書き出力。

## 注意点・運用メモ
- **管理者権限必須**（`Dism /Online` は管理者専用）
- インポートは **新規ユーザープロファイル作成時に反映**。既存ユーザーには即時反映されない（sysprep 前に適用するのが標準）
- エクスポートは現在ログオンユーザーの関連付けを取得。既存 XML は毎回上書き
- 冪等性は **非対称**: Export は毎回 DISM 実行（スナップショット用途）、Import は XML 不在で Skip / 存在で毎回 DISM 実行（差分チェックなし）
- 環境変数は使用しない。XML パスは `$PSScriptRoot\xml` 配下で解決

## 検証
Post-Apply Verification は **未実装**。`Dism /Import-DefaultAppAssociations` の `ExitCode` のみで成否判定し、`-Verified` は未渡し（Verified 列は空欄）。理由は、関連付けの実反映が「次回新規ユーザープロファイル作成時」という遅延型仕様のため、適用直後にレジストリを読んでも何も変わっておらず検証が事実上不可能なためです。エビデンスとしては XML ファイルの存在と DISM ExitCode で代替し、実反映は新規ユーザーログオン後に手動確認する運用です。
