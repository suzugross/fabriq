# file_delete (Standard)

**カテゴリ**: Maintenance
**メニュー名**: File Delete
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（`Test-Path` で削除後の非存在を検証）
**サブスクリプト**: なし

## 目的
キッティング工程で「マスターイメージから引き継いだ不要ファイル」「テスト用一時ファイル」「アプリケーションのキャッシュフォルダ」などを CSV ベースで一括削除するモジュールです。`%TEMP%` / `%LOCALAPPDATA%` などの環境変数展開に対応し、`IfNotFound=Skip` / `IfNotFound=Error` で「存在しなくても OK」「絶対あるべきもの」を行単位に区別できる設計のため、冪等な再実行と「必須ファイルの取りこぼし検出」の両立が可能です。

## 入力 (CSV)
`delete_list.csv` の主な列:
- `Enabled` … 1=実行 / 0=スキップ
- `Description` … 表示ラベル
- `TargetPath` … 削除対象パス（環境変数展開対応）
- `IfNotFound` … `Skip`（既定: 不在でも正常終了）/ `Error`（不在ならエラーとして記録）
- `Segment` … Segment フィルタ

## 主要ステップ
1. `delete_list.csv` 読み込み + 各パスの存在状況表示
2. 実行確認（AutoPilot 時は自動 Y）
3. 各対象を順次削除（フォルダは再帰削除）
4. **Post-Apply Verification**: 削除後に `Test-Path` で対象パスの非存在を全件検証
5. `New-BatchResult ... -Verified $verified` で返却

## 注意点・運用メモ
- 管理者権限は **状況依存**（`Program Files` 等の保護領域なら必須、ユーザー領域なら不要）
- フォルダ指定時は再帰削除のため、誤って親フォルダを書くと連鎖的に大量削除される。CSV メンテ時は注意
- `IfNotFound=Skip` を使えば「テストで一時ファイルを置いた → 本番マスターでは存在しない」状況でも冪等
- `IfNotFound=Error` は「絶対消すべきファイル（個人情報の残骸など）」の検出に使う

## 検証
Post-Apply Verification は **実装あり**。Step 4 で `Test-Path` を使って対象パスの非存在を検証し、全件 PASS のとき `-Verified $true`、1 件でも残存していれば `$false` を `New-BatchResult` に渡して返却。実行履歴の Verified 列に結果が記録され、Evidence Manager 側で削除完了が突合可能です。`copyfile_config` の Verification 非実装と対照的に、削除側は「不在の検証」が `Test-Path` で確実に判定できるため実装している、という設計判断です。
