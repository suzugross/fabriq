# builtin_admin_config (Extended)

**カテゴリ**: User Management
**メニュー名**: Built-in Admin Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（パスワード値が `Get-LocalUser` で読み返せず Verification 全体を意図的に非実装）
**サブスクリプト**: なし

## 目的
ビルトイン Administrator アカウントのパスワード設定・有効化/無効化・パスワード無期限フラグ・説明文を一括設定するモジュール。
`Enabled` 列が「実行有効化」と「Enable/Disable」の両方を兼ねる特殊仕様で、`0` の行でも処理対象として読み込まれ Disable 操作が走る（このため `Import-ModuleCsv -FilterEnabled` ではなく手動取得しているわけではないが、CSV を 1 行のみ前提としている）。

## 入力 (CSV)
`builtin_admin.csv`:
- **Enabled**: 1=Enable+パスワード設定, 0=Disable
- **Password**: 設定パスワード（`ENC:` プレフィックス暗号化対応）
- **PasswordNeverExpires**: 1=無期限 / 0=期限あり
- **Description**: アカウント説明文
- **Segment**: Segment フィルタ（`Import-ModuleCsv` が暗黙参照）

## 主要ステップ
1. CSV 読み込み（先頭 1 行のみ使用）
2. Password 空欄チェック
3. `Get-LocalUser` で Administrator 存在確認
4. 設定内容（Enable/Disable, パスワード期限, Description）を画面表示
5. `Confirm-ModuleExecution` で実行確認
6. Enable/Disable → Set-LocalUser でパスワード → PasswordNeverExpires → Description の順に適用

## 注意点・運用メモ
- パスワード値は冪等性チェックなしで毎回 `Set-LocalUser` を呼ぶ（`Get-LocalUser` ではパスワードを取得できないため構造的に比較不可）
- 何度実行しても結果は同じだが、書き込みは毎回発生する
- Administrator アカウントが存在しない環境（一部のクリーンインストール直後）では Error 終了
- 管理者権限必須（`Set-LocalUser` の前提）

## 検証
パスワード読み返し不可のため Verification 全体を未実装。Enabled / PasswordExpires は技術的には検証可能だが、肝心のパスワード検証ができないため整合性を保つ目的で全体非対応。`-Verified` 未指定で Verified 列は空欄。
