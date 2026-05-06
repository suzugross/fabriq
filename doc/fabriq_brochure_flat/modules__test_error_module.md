# test_error_module (Standard)

**カテゴリ**: Test
**メニュー名**: Test Error Module
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 該当なし（常に Error を返す性質）
**サブスクリプト**: `test_error.ps1`

## 目的
意図的に必ず `Status=Error` で終了するテスト用モジュール。AutoPilot の ErrorMode（`skip` /
`retry` / 未指定時の Retry/Skip ダイアログ）の動作検証に使う。副作用（ファイル / レジストリ /
サービスの変更）は一切なく、どの環境でも安全に走らせられる。本番 Profile に組み込む用途は
想定せず、開発・検証専用の足場部品。

## 入力 (CSV)
設定 CSV なし（module.csv のみで動作）。

## 主要ステップ
1. モジュール表示（"Test Error Module"）
2. 実行確認（Y/N、AutoPilot 自動 Y）
3. `Show-Error` で模擬エラー表示
4. `New-ModuleResult -Status "Error"` を返却

## 注意点・運用メモ
- 管理者権限不要、副作用なし
- 本番 Profile（Master_Config*.csv 等）には絶対組み込まない
- 検証パターン例:
  - **ErrorMode=skip**: モジュール 20 で発生したエラーが自動 skip され 30 が実行される
  - **ErrorMode=retry**: 最大 5 回リトライされ、毎回失敗で最終的に Error として記録
    （AutoPilotMaxRetry=5 が main.ps1 側の上限）
  - **ErrorMode 空欄**: AutoPilot 中に Retry/Skip ダイアログがポップアップ
- メモ `project_autopilot_skip_rejected.md` の通り skip/timeout 機能自体は否決済み
  だが、Profile 列の ErrorMode は別概念で本モジュールはその検証に使う

## 検証
Post-Apply Verification は概念的に該当しない（常に Error を返すモジュールなので
「適用された設定の読み返し」というステップが存在しない）。`-Verified` は渡さず
履歴 Verified 列は空欄。本モジュール自体が他モジュールの Verified 機構の検証ハーネス。
