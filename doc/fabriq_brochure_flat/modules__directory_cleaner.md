# directory_cleaner (Extended)

**カテゴリ**: Maintenance
**メニュー名**: Directory Cleaner
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（contents モードで削除直後に新規一時ファイルが生成され得るため false PASS/FAIL を回避）
**サブスクリプト**: なし

## 目的
CSV で指定したディレクトリの「中身のみ削除（contents）」または「ディレクトリごと削除（directory）」を実行するクリーンアップモジュール。
最大の特徴はハードコードされた**禁止パスホワイトリスト**による多重ガード。`C:\Windows`, `C:\Program Files`, `%USERPROFILE%`, fabriq リポジトリルートなどシステム重要パスを 3 種類のチェック（完全一致 / 親子関係 / セグメント数 < 3）で自動ブロックする。表示ステップと実行ステップの両方で二重チェックされ、ヒットすると `[BLOCKED]` でスキップ。

## 入力 (CSV)
`clean_list.csv`:
- **Enabled**: 有効フラグ
- **TargetPath**: 対象ディレクトリ（`%LOCALAPPDATA%` 等の環境変数展開対応）
- **Mode**: `contents`（中身のみ削除、フォルダは残す） / `directory`（フォルダごと削除）
- **Description**: 表示用説明
- **Segment**: Segment フィルタ

## 主要ステップ
1. CSV 読み込み + `Expand-UserEnvironmentVariables` で環境変数展開
2. Mode 値検証（`contents`/`directory` 以外は Error）
3. ドライラン表示（`[BLOCKED]` / `[NOT FOUND]` / `[DELETE]` を色分け）
4. `Confirm-ModuleExecution`
5. 削除ループ（実行時にも `Test-ForbiddenPath` を再評価する double-gate / 不在 / 既に空 → Skip）
6. `New-BatchResult` で集計

## 注意点・運用メモ
- 禁止パスチェックは 3 種類（A: 完全一致、B: 対象が禁止パスの親、C: セグメント数 3 未満）
- 削除中にロックされたファイルは `SilentlyContinue` で個別スキップし、件数を集計表示（contents モードでは部分削除を Warning 扱い）
- Past incident: 2026-04-25 にユーザーのデスクトップを Join-Path 失敗で破壊した事故への対策として禁止パスガードが入っている経緯あり

## 検証
contents モードでは削除直後に OS が新規一時ファイルを生成する性質があり、読み返し検証は false FAIL/PASS リスクが高いため意図的に未実装。`-Verified` 未渡しで履歴の Verified 列は空欄。
