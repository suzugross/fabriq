# hostname_config (Standard)

**カテゴリ**: Network
**メニュー名**: Change Hostname
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（保留中ホスト名のレジストリ読み返し）
**サブスクリプト**: `hostname_config.ps1`（単一スクリプト、ローカル CSV なし）

## 目的
ホスト名（コンピュータ名）を変更します。設定値は専用 CSV ではなく、
fabriq 起動時に `hostlist.csv` で選択したホストの `NewPCName` 列由来の
環境変数 `$env:SELECTED_NEW_PCNAME` のみから取得します。
hostlist 中心アーキテクチャを象徴するモジュールで、
Profile から呼ばれた際は無人で自動適用されます（hostlist 選択そのものが同意の代わり）。

## 入力 (CSV)
ローカル CSV なし。以下の環境変数のみ使用：
- `$env:SELECTED_NEW_PCNAME`: 新しいホスト名（hostlist.csv の NewPCName 列由来、空欄時は Skip）
- `$env:COMPUTERNAME`: 現行ホスト名（冪等性チェックに使用）

## 主要ステップ
1. `$env:SELECTED_NEW_PCNAME` 取得。未設定なら Skip
2. 現在ホスト名と目標値を表示
3. 冪等性チェック: 既に同名なら Skip
4. 実行確認（AutoPilot は自動 Y）
5. `Rename-Computer -NewName -Force` でホスト名変更
6. Step 5.5: Post-Apply Verification（後述）
7. `New-ModuleResult -Verified` で結果返却（再起動必要メッセージ付き）

## 注意点・運用メモ
- 管理者権限必須（`Rename-Computer` 自体が AdminGuard 対象）
- 反映は OS 再起動が必要。検証時点では `$env:COMPUTERNAME` は旧名のまま
- ホスト名規約違反（15 文字超、不正文字）はエラーで中断
- Domain 参加状態の場合、`Rename-Computer` には認証情報が必要となるが本モジュールは
  ワークグループ前提（domain_join モジュールとの順序依存）

## 検証
Post-Apply Verification は実装済み。`Rename-Computer` は再起動で反映されるため
現行ホスト名は変わらない代わりに、レジストリの保留値
`HKLM:\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName\ComputerName` を読み返し、
目標値と一致するかで「適用が受理されたこと」を検証します。
結果は `New-ModuleResult -Verified $true/$false` で返却され、
実行履歴 / evidence_config の Verified 列に True/False が記録されます。
