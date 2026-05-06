# firewall_rule_config (Standard)

**カテゴリ**: Security
**メニュー名**: Firewall Rule Export / Firewall Rule Import
**VERSION**: 0.1.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（Export: ファイル整合性、Import: Name-set 包含検証）
**サブスクリプト**:
- `firewall_rule_export.ps1` … `netsh advfirewall export` でファイアウォールポリシー全体のスナップショット保存
- `firewall_rule_import.ps1` … `netsh advfirewall import` でポリシー全体を復元（**破壊的・全置換**）

## 目的
Windows ファイアウォール全体（rule + profile 状態 + logging 設定 + IPsec）を **`.wfw` 形式の真実源** として丸ごとバックアップ／復元するモジュールです。`netsh advfirewall export/import` を使い、人間可読なサイドカー（`rules.json` / `rules_show.txt` / `profiles.json` / `manifest.txt` / `rule_names.txt`）を併産することで、バイナリ `.wfw` の中身を監査可能にしています。Import は **`IAcknowledgeReplace=1` が無いと暴発しない** ゲート設計で、AutoPilot 中でも誤って全ポリシー置換が走るリスクを排除しています。Export 成功時は CSV 末尾に対応 Import 行（`Enabled=0`, `IAcknowledgeReplace=0`）を自動追記して復元動線を作る運用補助も搭載。

## 入力 (CSV)
`firewall_rule_list.csv` の主な列:
- `Enabled` … 1=実行 / 0=スキップ
- `Mode` … `Export` / `Import`
- `SourcePath` … Import 時の `.wfw` パス（UNC 可、相対パスは `module\backup\` 基準で解決、ディレクトリ指定なら `policy.wfw` を自動付与）
- `DestinationPath` … Export 時の保存先（空欄なら `module\backup\<yyyyMMdd_HHmmss>\`）
- `IAcknowledgeReplace` … Import 時の **暴発防止承認フラグ**（0=拒否 / 1=実行許可）
- `Description` / `Segment`

## 主要ステップ
[Export]
1. CSV 読み込み（Enabled=1 + Mode=Export）
2. Pre-flight（管理者権限、`netsh` 利用可否）
3. Dry-run 表示
4. 実行確認（AutoPilot 時は自動 Y）
5. `netsh advfirewall export` 実行 + サイドカー（rules.json / profiles.json / rules_show.txt / rule_names.txt / manifest.txt with SHA256）出力
5.5 **Verification**: `policy.wfw` 存在 + サイズ > 1KB / `manifest.txt` 存在 / `rule_names.txt` 行数 == manifest 記録数
6. CSV 末尾に Import 行を自動追記（Enabled=0, IAcknowledgeReplace=0）+ `New-BatchResult -Verified`

[Import]
1. CSV 読み込み（Enabled=1 + Mode=Import + IAcknowledgeReplace=1）
2. Pre-flight + 相対パス解決
3. manifest 読み込み + OS 版差異の警告
4. 実行確認
5. `netsh advfirewall import` 実行
5.5 **Verification**: `rule_names.txt` の各 Name が Import 後 `Get-NetFirewallRule` に **包含されているか**（after が expected の superset で OK）。旧 backup なら count-only fallback
6. `New-BatchResult -Verified`

## 注意点・運用メモ
- **管理者権限必須**（両スクリプト）
- `netsh advfirewall import` は **全ポリシーを置換**（追加ではない）
- 取り込み元 `.wfw` の OS 版が現在 OS 版と異なると部分失敗の可能性あり（manifest との照合で警告）
- 相対パス先頭に `backup/` を付けると重複展開で not found（操作者向け注意点）
- Excel で CSV を開いた状態で Export を実行するとファイルロックで Import 行追記失敗 → 警告のみで Export は成功扱い

## 検証
Post-Apply Verification は **実装あり**、両モードで `-Verified` を `New-BatchResult` に返却。
- **Export**: `policy.wfw` 存在 + サイズ > 1KB、`manifest.txt` 存在、`rule_names.txt` 行数 == manifest 記録数。**現在の `Get-NetFirewallRule` count とは比較しない**（Windows 動的変動を意識した設計判断）
- **Import**: `rule_names.txt` の Name 集合が Import 後にすべて存在するか **包含検証**（後から mpssvc / AppX / GPO が動的 rule を追加しても影響なし）。1 つでも欠損すれば FAIL（欠損 Name を最大 5 個報告）
- 旧 backup（`rule_names.txt` 不在）は count-only fallback。Windows 動的変動で誤 fail 可能性あり（弱い検証）
