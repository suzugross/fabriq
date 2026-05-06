# restart_config (Standard)

**カテゴリ**: System
**メニュー名**: Restart with AutoRun
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（再起動で自プロセスが消えるため検証不可、`-Verified` 未渡し）
**サブスクリプト**: なし

## 目的
Fabriq 自身を `HKLM\...\RunOnce` に登録した上で、PC を再起動するモジュール。
Profile の最終ステップに置くことで「再起動 → 再起動後に Fabriq が自動再開」のループを成立
させる、フレームワークの中核フロー部品。本モジュールの主機能は RunOnce 登録であり、
再起動はその後の付随アクションとして連結されている。

## 入力 (CSV)
設定 CSV なし。Order/Enabled は `module.csv` のみで管理（Order=99 で Profile 末尾向け）。

## 主要ステップ
1. `restart_config` から 3 階層上 (`..\..\..`) を fabriq ルートとして解決し、`Fabriq.exe` 存在確認
2. `HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce` の現在値を確認・表示
3. RunOnce 登録予定内容（Path / Name / Value）をドライラン表示
4. 実行確認（AutoPilot 時でも確認スキップしない設計 = 不可逆操作の安全策）
5. `RunOnce\FabriqAutoStart` に `"...\Fabriq.exe"` を書き込み
6. `New-ModuleResult -Status Success` を先に確定（再起動で履歴保存が間に合うように）
7. `Invoke-CountdownRestart -Seconds 10` で 10 秒カウントダウン後 OS 再起動

## 注意点・運用メモ
- 管理者権限必須（HKLM RunOnce 書き込み）
- AutoPilot 非対応（メモ `feedback_autopilot_wording.md` の通り、AutoPilot は「無人」ではなく
  オペレーター立ち会い前提のスキップだが、再起動はそれでも明示確認を要求するクリティカル例外）
- RunOnce は OS 再起動後の自動起動が成功すれば値が消えるため、リブート後の Fabriq 自動起動の
  有無で登録成否を判定する運用
- `Invoke-CountdownRestart` は kernel/common.ps1 提供。Ctrl+C キャンセル余地あり
- Order=99 だが、Profile 設計次第ではモジュール途中段階のリブートにも使える
  （Profile 上で複数回呼ぶケース）

## 検証
本モジュールに Post-Apply Verification は実装されていない。RunOnce 登録は理屈上は
レジストリ読み返しで検証可能だが、登録直後にプロセス自身が消える（OS 再起動）ため、
モジュールから戻り値を返したあとの Verified 評価フェーズが存在しないという技術的制約により
意図的に未実装。`New-ModuleResult` 呼び出しに `-Verified` を渡していないため、
実行履歴の Verified 列は空欄になる。再起動後の Fabriq 自動起動が事実上の検証手段。
