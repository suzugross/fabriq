# log_uploader (Extended)

**カテゴリ**: Evidence
**メニュー名**: Log Upload
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（robocopy ExitCode による成否判定はあるが厳密な読み返し検証なし）
**サブスクリプト**: なし

## 目的
プロファイル実行完了後に呼ばれるアップロードモジュール。`logs/` と `evidence/` を robocopy で UNC 共有またはローカルパスへ複製する。
モジュール内に専用 CSV を持たず、**fabriq 共通の `kernel\csv\log_destinations.csv` を参照**する点がアーキテクチャ上の特徴（同じ宛先設定を他モジュールとも共有可能）。finalize（プロファイル末尾）で自動実行される想定だが、Script Menu からの手動実行にも対応。

## 入力 (CSV)
`kernel\csv\log_destinations.csv`（モジュール外、共通配置）:
- **Path**: ローカルパス or UNC パス
- **Type**: `UNC` / `Local`
- **Enabled**: 有効フラグ
- **AuthUser**: UNC 認証ユーザー（省略可）
- **AuthPass**: UNC 認証パスワード（省略可、`ENC:` プレフィックス対応）
- **Description**: 説明（Segment 列なし、全セグメント共通）

## 主要ステップ
1. `kernel\csv\log_destinations.csv` を読み込み（Enabled=1 のみ）
2. `kernel\json\session.json` から MediaSerial を取得（フォルダ名に使用）
3. 宛先フォルダ名を生成（Unified モード: evidence ルートのフォルダ名 / Fallback: `タイムスタンプ_ホスト名_シリアル番号`）
4. `logs/` と `evidence/` の中身有無確認（両方空なら Skipped 終了）
5. トランスクリプト一時停止（ログファイルロック解除）
6. 各宛先に対して: `net use` で UNC 認証 → `New-Item` で宛先フォルダ作成 → `robocopy /E /NJH /NJS /NDL /NP /R:2 /W:1` で logs/, evidence/ をコピー → `session.json` をメタデータとしてコピー
7. `finally` で `net use /delete` 実行（必ず切断）+ `AuthPass` を null クリア
8. トランスクリプト再開
9. 全件成功なら Success / 一部失敗なら Partial / 全失敗なら Error

## 注意点・運用メモ
- finalize 時に自動呼び出しされる（Profile 末尾の標準モジュール）
- robocopy オプション `/R:2 /W:1` で再試行 2 回・待機 1 秒（一時的なネットワーク不安定に強め）
- AuthUser/AuthPass の片方だけ空欄なら認証スキップして現在のユーザーコンテキストでアクセス
- `ENC:` 暗号化パスワードは `kernel/common.ps1` の復号処理に委譲

## 検証
robocopy ExitCode 単位で成功/失敗集計し、`New-ModuleResult` で `Success` / `Partial` / `Error` を返却。`-Verified` 未渡しで Verified 列は空欄。手動検証は宛先側のフォルダ（タイムスタンプ_ホスト名_シリアル番号）を直接確認。
