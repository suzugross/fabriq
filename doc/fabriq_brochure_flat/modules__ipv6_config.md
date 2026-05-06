# ipv6_config (Extended)

**カテゴリ**: Network
**メニュー名**: IPv6 Settings
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: なし（`Set-NetAdapterBinding` の戻り値を信頼する設計）
**サブスクリプト**: なし

## 目的
ネットワークアダプターの IPv6 バインディング（`ms_tcpip6`）を有効化／無効化するシンプルなモジュール。
ワイルドカード対応の `AdapterPattern` で、OS 言語による「イーサネット」「Ethernet」名差を 1 CSV 内で吸収できる。

## 入力 (CSV)
`ipv6_list.csv`:
- **Enabled**: 有効フラグ
- **AdapterPattern**: アダプター名のワイルドカードパターン（例: `イーサネット*`, `Ethernet*`, `Wi-Fi*`）
- **IPv6State**: `0`=Disable / `1`=Enable（数値文字列で判定。`Enabled`/`Disabled` 文字列は不可）
- **Description**: 説明
- **Segment**: Segment フィルタ

## 主要ステップ
1. `Test-AdminPrivilege` で権限チェック
2. CSV 読み込み（手動で Enabled=1 をフィルタリング、無効行も表示用に保持）
3. 対象アダプター一覧表示（無効化されたエントリは `[DISABLED]` でグレー表示）
4. `Confirm-ModuleExecution`
5. `Get-NetAdapter | Where-Object Name -like $pattern` でマッチ → `Set-NetAdapterBinding -ComponentID ms_tcpip6 -Enabled $targetState`
6. `New-BatchResult` 集計

## 注意点・運用メモ
- AdapterPattern にワイルドカード使用可
- 言語環境で名前が変わるため日本語名と英語名の両方を登録しておくのが定石
- マッチするアダプターが 0 件のパターンは Skip 集計

## 検証
未実装。手動確認は `Get-NetAdapterBinding -Name <name> -ComponentID ms_tcpip6`。`-Verified` 未渡しで Verified 列は空欄。
