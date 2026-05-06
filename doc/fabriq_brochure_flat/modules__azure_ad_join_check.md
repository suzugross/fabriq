# azure_ad_join_check (Extended)

**カテゴリ**: System
**メニュー名**: Azure AD Join Check
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 不可（状態確認のみで適用ステップが存在しない。Status 自体が検証結果を兼ねる）
**サブスクリプト**: なし（`azure_ad_join_check.ps1` 単体）

## 目的
`dsregcmd /status` を実行して Azure AD (Entra ID) Join が完了しているかを確認するモジュール。
未完了であれば Error を返すため、`script_looper` と OnError 条件で組み合わせることで「Join 完了まで自動で待ち続ける」ループを構築できる。
状態確認のみでシステムへの変更は一切行わない非破壊モジュールであり、`Confirm-ModuleExecution` も呼び出さず AutoPilot/手動どちらでも同一動作。

## 入力 (CSV)
なし。`dsregcmd /status` の出力から `AzureAdJoined` 値を自動判定する。

## 主要ステップ
1. `dsregcmd.exe` の存在確認（`%SystemRoot%\System32\dsregcmd.exe`）
2. `dsregcmd /status` を実行し標準出力を取得
3. 出力内 `AzureAdJoined : YES` を正規表現で検出して画面表示
4. YES なら Success、NO や検出不可なら Error を `New-ModuleResult` で返却

## 注意点・運用メモ
- 単体実行も可能だが、本来の用途は `script_looper` の `Condition=OnError` でリトライ駆動させること
- 管理者権限不要（`dsregcmd` は一般ユーザーで動作）
- Windows 10 以降前提（`dsregcmd.exe` が必要）

## 検証
適用ステップが無いため Post-Apply Verification は実装されない。`-Verified` を渡さないため履歴の Verified 列は空欄。Status 値（Success/Error）が「現在 Join 済みかどうか」をそのまま表す。
