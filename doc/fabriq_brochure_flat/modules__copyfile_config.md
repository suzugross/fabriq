# copyfile_config (Standard)

**カテゴリ**: Maintenance
**メニュー名**: File Copy Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 不可（誤 PASS リスクのため意図的に非実装）
**サブスクリプト**: なし（コピー元実体は `source/` 配下に配置）

## 目的
`source/` ディレクトリに配置したファイル／フォルダを、CSV に記述された宛先（`%USERPROFILE%`、`%LOCALAPPDATA%`、`C:\Program Files\...` 等）にコピーするモジュールです。フォルダ指定時は `Copy-Item -Recurse -Force` で中身ごと再帰コピーします。コピー後に **Mark-of-the-Web（Zone.Identifier ADS）を自動解除** するため、ネットワーク経由で受け取ったテンプレート／設定ファイルでも「ブロック解除済み」状態で配置されるのが運用上の利点です。`Expand-UserEnvironmentVariables` により昇格セッションでもログオンユーザーの `%USERPROFILE%` を解決します。

## 入力 (CSV)
`copy_list.csv` の主な列:
- `Enabled` … 1=実行 / 0=スキップ
- `FileName` … `source/` 直下のファイル名 or フォルダ名（フォルダなら再帰）
- `DestPath` … コピー先ディレクトリ（環境変数展開対応: `%USERPROFILE%` / `%LOCALAPPDATA%` / `%APPDATA%` / `%ProgramData%` / `%SystemRoot%` 等）
- `Overwrite` … 1=上書き / 0=既存スキップ
- `Description` / `Segment`

## 主要ステップ
1. `copy_list.csv` 読み込み
2. `source/` ディレクトリ存在確認（無ければ Error）
3. （前処理）DestPath 環境変数展開
4. Dry-run 表示（`[Missing]` / `[Current]` / `[Overwrite]` / `[Copy]`）
5. 実行確認（AutoPilot 時は自動 Y）
6. コピー実行（DestPath 自動作成 → `Copy-Item -Recurse -Force` → `Remove-ZoneIdentifier` で MoTW 解除）
7. `New-BatchResult` で結果集計

## 注意点・運用メモ
- 管理者権限は **状況依存**: `%USERPROFILE%` 等のユーザー領域なら不要、`C:\Program Files` 等の保護領域なら必須
- フォルダ指定時は `-Recurse -Force` で丸ごと上書き。個別ファイル選別は不可
- UNC パス（`\\server\share`）は認証次第で失敗。事前 `net use` 推奨
- `Overwrite=0` 指定は既存ファイル保護で擬似冪等性あり（タイムスタンプ更新も発生しない）

## 検証
Post-Apply Verification は **意図的に非実装**（CLAUDE.md 記載の除外モジュール）。理由は、`Copy-Item` 成功後に `Test-Path` で存在を見るだけだと「コピーしたつもりが内容が古い」「権限が前のまま」「実体が古いコピー」といったケースで誤 VERIFIED となるリスクが大きいためです。真の検証には内容ハッシュ比較・ACL 比較・MoTW 二重確認が必要ですが、配置先・用途が多様で一律実装が困難なため、安易な検証で誤 PASS を出さない方針を選択しています。代替として、コピー直後のログで `[Copy]` / `[Overwrite]` / `[Skip]` / Error を明確に記録し、Mark-of-the-Web は `Remove-ZoneIdentifier` で必ず解除しています。`-Verified` は未渡しで Verified 列は空欄。厳密検証は運用側で `Get-FileHash` / `icacls` を別途実行します。
