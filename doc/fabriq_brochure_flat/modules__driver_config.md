# driver_config (Standard)

**カテゴリ**: System
**メニュー名**: Driver Export / Driver Import
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（限定的な検査として `*.inf` カウントと `pnputil` ExitCode は判定）
**サブスクリプト**:
- `driver_export_config.ps1` … `Export-WindowsDriver -Online` で現 OS のサードパーティドライバを `driver/{model}/` に書き出し
- `driver_import_config.ps1` … `pnputil /add-driver "*.inf" /subdirs /install` で対象 PC に展開

## 目的
キッティング工程で「マスター PC からドライバ一式をエクスポート → 同モデルの量産機にインポート」という典型ワークフローを CSV 1 行で表現するモジュールです。Export では現 OS にインストールされたサードパーティドライバ（OS 標準ドライバを除く）を `driver/{モデル名}/` 配下に展開し、Import では `pnputil` でそれを再投入します。`model` 列を空欄にすると **`Win32_ComputerSystem.Model` から自動取得** するため、ホスト名ベースではなく機種ベースで自動振り分けが可能です（例: ThinkPad X1 → `ThinkPad_X1` フォルダ）。

## 入力 (CSV)
`driver.csv` の主な列:
- `Enabled` … 1=実行 / 0=スキップ
- `Id` … 識別連番
- `model` … ドライバフォルダ名の明示指定（空欄ならホストの `Win32_ComputerSystem.Model` を自動採用、サニタイズ: 空白→`_`、`\/:*?"<>|` 除去、80 文字上限）
- `Segment` … Segment フィルタ

## 主要ステップ
（Export / Import で同形）
1. `driver.csv` 読み込み
2. Pre-flight: 管理者権限確認 / `driver/` ディレクトリ存在確認
3. Dry-run 表示
4. 実行確認（AutoPilot 時は自動 Y）
5. 適用ループ（Export: `Export-WindowsDriver -Online -Destination ...`、Import: `pnputil /add-driver`）
6. 結果集計

## 注意点・運用メモ
- **管理者権限必須**（両スクリプトとも）
- Export はサードパーティドライバのみ対象（OS 標準ドライバは含まれない）
- Export 時、既存の `driver/{model}/` フォルダは中身がクリアされてから再エクスポート（差分追記ではなく完全置換）
- Import 後に再起動が必要な場合あり（`pnputil` 終了コード 3010）
- Export 時はドライバ容量分のディスク空き容量が必要

## 検証
Post-Apply Verification は **限定的**。`-Verified` フラグは渡しません。
- Export: 出力先フォルダの `*.inf` 数をカウント表示し、0 件ならエラー扱い
- Import: 実行前後の `*.inf` カウント参照 + `pnputil` ExitCode（0/3010=Success）

OS に組み込まれたドライバのバージョン比較等の厳密検証は行いません。必要に応じ `pnputil /enum-drivers` で手動確認する運用です。
