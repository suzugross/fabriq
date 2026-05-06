# firewall_rule_make_config (Standard)

**カテゴリ**: Security
**メニュー名**: Firewall Rule Maker
**VERSION**: 0.1.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（DisplayName 取得 + 主要属性比較）
**サブスクリプト**: なし

## 目的
CSV に記述された定義から **個別の Windows ファイアウォール rule を `New-NetFirewallRule` で作成** するモジュールです。`firewall_rule_config` がポリシー全体を `.wfw` で丸ごとバックアップ／復元する **マクロ視点** を担うのに対し、本モジュールは「site-specific な rule を 1 本ずつ追加」する **マイクロ視点** を担います。冪等性は同一 `DisplayName` の rule 既存検出で SKIP、`Name` 列を明示すると内部 GUID 名で厳密マッチが効くため、後の削除や更新が確実に追跡可能になります。CSV 24 列で `Profile` / `Protocol` / `LocalPort` / `RemoteAddress` / `Program` / `Service` / `InterfaceType` / `LocalUser` (SDDL) / `EdgeTraversalPolicy` まで網羅し、`New-NetFirewallRule` の主要パラメータを CSV 1 行で表現できる広さがあります。

## 入力 (CSV)
`firewall_rule_make_list.csv` の 24 列（要約）:
- **必須**: `Enabled`, `DisplayName`, `Direction`, `Action`
- **ID**: `Name`（GUID 様。未指定は自動生成。指定時は冪等性チェックを Name で実施）
- **推奨**: `RuleEnabled`, `Profile`（複数 `;` 区切り）, `Protocol`, `LocalPort`, `RemoteAddress`
- **詳細**: `Description`, `Group`, `RemotePort`, `LocalAddress`, `IcmpType`, `EdgeTraversalPolicy`
- **対象**: `Program`, `Service`, `InterfaceType`, `InterfaceAlias`
- **セキュリティ（SDDL）**: `LocalUser`, `RemoteUser`, `RemoteMachine`
- **fabriq 標準**: `Segment`

セル内複数値は CSV 区切り `,` と衝突しないよう **`;` 区切り**（例: `Profile=Domain;Private`、`LocalPort=80;443`）。

## 主要ステップ
1. CSV 読み込み（Enabled=1）
2. 各行を validation（Direction / Action / Profile / Protocol / Port 解釈可否） + 既存 rule 重複チェック
3. plan を `[CREATE]` / `[SKIP - exists]` / `[INVALID]` で分類してプレビュー
4. 実行確認（AutoPilot 時は自動 Y）
5. `[CREATE]` のみ `New-NetFirewallRule` を splat で呼出
5.5 **Post-Apply Verification**: 作成 rule ごとに存在 / Name / Direction / Action / Enabled / Profile / EdgeTraversalPolicy を比較
6. `New-BatchResult ... -Verified $verified` で返却

## 注意点・運用メモ
- **管理者権限必須**（`New-NetFirewallRule` のため）
- 冪等性: 同 `DisplayName` で複数 rule が存在する場合も SKIP（重複は安全側で touched しない）
- `Program` パスが存在しない場合は警告のみで rule 作成は継続（Windows は不在パスでも rule 作成を許可）
- 不正な行は実行前 reject（Fail カウント加算、他行は継続）
- `Service` 列を使う場合 Direction との組み合わせ制約あり（cmdlet が例外を返すケース。試行錯誤推奨）
- 3 ファイアウォール系モジュールの併用順序（推奨）:
  1. `firewall_rule_config (Import)` でベースライン restore
  2. `firewall_rule_make_config` で site-specific rule 追加（本モジュール）
  3. `firewall_config` で profile on/off の最終調整

## 検証
Post-Apply Verification は **実装あり**。Step 5.5 で作成された rule ごとに `Get-NetFirewallRule -Name <created.Name>` を呼び、以下を比較:
- 存在
- Name（CSV 指定時）
- Direction / Action（CSV 値と一致）
- Enabled（CSV `RuleEnabled` と一致、未指定時は True 想定）
- Profile（CSV 値を集合として比較、順序非依存）
- EdgeTraversalPolicy（CSV 指定時）

全 PASS なら `-Verified $true` で `New-BatchResult` に返却。**`InterfaceType` / `InterfaceAlias` / `LocalUser` / `RemoteUser` / `RemoteMachine` は v1 verification の対象外**（別 cmdlet `Get-NetFirewall*Filter` 経由で読まないと取れないため、cmdlet が正しく適用したことを信頼する設計）。必要なら `Get-NetFirewallInterfaceTypeFilter -AssociatedNetFirewallRule <rule>` で手動確認。
