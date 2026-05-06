# signout_config (Standard)

**カテゴリ**: System
**メニュー名**: Sign-Out
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（プロセス自身が消えるため検証不可、`-Verified` 未渡し）
**サブスクリプト**: なし

## 目的
現在のユーザーセッションを `logoff.exe` 経由でサインアウトするモジュール。
Profile 末尾に配置することで、キッティング作業の終了をユーザーセッションごと閉じる
クリーンアップステップとして機能する。Order=100（最終位置）の指定通り、本モジュールが
走ると後続モジュールはまったく実行されない。

## 入力 (CSV)
設定 CSV なし。

## 主要ステップ
1. 警告バナー表示（"fabriq will be TERMINATED after sign-out"）
2. 実行確認（AutoPilot 時は自動 Y）
3. `New-ModuleResult -Status Success -Message "Sign-out initiated"` を先に確定
   （`$global:_LastModuleResult` フォールバックにより、フレームワークが履歴を捕捉できるように）
4. `Show-Warning` で再度警告
5. `Invoke-CountdownSignout -Seconds 7` で 7 秒カウントダウン後にサインアウト実行

## 注意点・運用メモ
- AutoPilot 対応（無人実行を前提とした運用で重要）。
  メモ `feedback_autopilot_wording.md` の通り「完全無人」ではなくオペレーター立ち会い前提だが、
  Sign-Out は Wizard 終了時のセッション切替として意図的に AutoPilot 自動 Y を許可
- Profile 上の配置位置は必ず最後（Order=100）。途中に置くと後続モジュールが死ぬ
- `logoff.exe` は対話セッションの終了が主用途で、コンソールセッションでは
  fabriq 自体が即座に終了する
- ローカル状態 JSON や履歴ファイルへの書き込みは Step 3 で完了させる設計（プロセス死後の
  書き込みは保証されないため）
- カウントダウン中の Ctrl+C などのキャンセル余地は `Invoke-CountdownSignout` 側の実装による

## 検証
本モジュールに Post-Apply Verification は実装されていない。`logoff.exe` の実行直後に
fabriq プロセス自身が終了するため、モジュール内で読み返しを行う余地が技術的に存在しない。
`New-ModuleResult` 呼び出しに `-Verified` を渡していないため履歴の Verified 列は空欄。
事実上の検証は「サインアウトが起き、再ログオン後に Fabriq.exe が起動していないこと」
で確認する運用。
