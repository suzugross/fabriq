# domain_join (Standard)

**カテゴリ**: Network
**メニュー名**: Domain Join
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（再起動後に反映されるため、再起動前の検証では情報量が足りない）
**サブスクリプト**: なし

## 目的
PC を Active Directory ドメインに参加させるモジュールです。`Add-Computer` コマンドレットを使用し、参加に失敗した場合は **GUI ダイアログでエラー内容を表示してリトライループ** に入る設計のため、現場担当者がパスワード入力ミスなどに対処しやすくなっています。AutoPilot 運用時は確認スキップで自動参加させますが、リトライダイアログが立ち上がる仕様上、Profile 側で `ErrorMode=retry` と組み合わせる運用が想定されます。

## 入力 (CSV)
`domain.csv` の主な列:
- `Enabled` … 1=実行 / 0=スキップ
- `domain` … ドメイン名（例: `example.local`）
- `user` … ドメイン参加用アカウント
- `pass` … パスワード（**`ENC:` プレフィックス暗号化対応**）
- `dns` … DNS サーバ IP（接続確認 Ping 用）
- `Description` / `Segment`

## 主要ステップ
1. `domain.csv` から有効エントリ読み込み
2. DNS サーバへの Ping 接続確認
3. 実行確認（AutoPilot 時は自動 Y）
4. `Add-Computer` でドメイン参加（失敗時は GUI ダイアログでリトライループ。`adminstop` 入力で中断）
5. `New-ModuleResult` で結果返却

## 注意点・運用メモ
- **管理者権限必須**（`Add-Computer` 実行のため）
- DNS サーバへのネットワーク接続が必須（Ping 失敗で Error 検出）
- ENC: 暗号化パスワードを使う場合は Fabriq 起動時のマスターパスフレーズ入力が必要
- リトライダイアログは無人運用に向かない。AutoPilot 構成時は Profile 側 `ErrorMode=retry` と組み合わせ
- 反映は **再起動後**（Windows 仕様）。本モジュールは再起動はトリガーしない（後続 `restart_config` で集約再起動）

## 検証
Post-Apply Verification は **未実装**。理由は、ドメイン参加状態は `Add-Computer` 実行直後ではなく **次回再起動後** に `(Get-WmiObject Win32_ComputerSystem).PartOfDomain = $true` として反映される Windows 仕様のため、再起動前のメモリ上で検証しても判定が曖昧になるためです。`-Verified` は未渡しで Verified 列は空欄。`Add-Computer` の成否とリトライダイアログでの到達可否で結果判定し、再起動後の確認は手動または `evidence_config` の収集レポートで行う運用です。
