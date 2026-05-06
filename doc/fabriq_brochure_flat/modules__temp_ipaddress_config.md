# temp_ipaddress_config (Standard)

**カテゴリ**: Network
**メニュー名**: Temp IP Address
**VERSION**: 0.1.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（IP/Prefix/Gateway/DNS 読み戻し + Gateway ping）
**サブスクリプト**: なし（GUI 選択ダイアログを WinForms で同居）

## 目的
顧客要件で「本番 IP がまだ旧 PC で使用中」の状況下、新 PC を一時的な作業 IP で立ち上げる
ためのモジュール。CSV に列挙された候補 IP プールを GUI で作業者に提示し、選択された IP を
NIC に付与する。Windows DAD（Duplicate Address Detection）が assignment 時に同 LAN 上の
衝突を検出し、衝突した行は `[DUPLICATE]` でグレーアウト表示して再選択を促す。本番 IP への
切替は別モジュール `ipaddress_config` で行う排他関係。

## 入力 (CSV)
`temp_ipaddress_list.csv`
- `Enabled`: 0=スキップ / 1=プールに含める
- `IPAddress`: 候補 IP（IPv4、strict 4-octet）
- `SubnetPrefix`: CIDR prefix (1-32)
- `Gateway`: デフォルト GW（空欄ならスキップ）
- `DNS1`, `DNS2`, `DNS3`: DNS（空欄なら DNS 設定変更しない）
- `AdapterPattern`: NIC 名ワイルドカード（例 `Ethernet*`）— 全行同一推奨（v1 では混在不許可）
- `Description`: GUI 表示用
- `Segment`: Segment フィルタ（任意）

デフォルト同梱: 192.168.100.201〜210（10 IP）すべて Enabled=0 のテンプレート。

## 主要ステップ
1. `Test-AdminPrivilege` チェック
2. CSV 読込（Enabled=1）
3. 各行 validate（IP/Prefix/Gateway/DNS 形式チェック）
4. 全行の AdapterPattern 一致確認 → NIC 解決
5. NIC subnet 整合性チェック（情報表示のみ、機能は継続）
6. GUI ダイアログ表示（AutoPilot でも必ず modal、kitting 中 1 回限りの戦略的判断のため）
7. 選択 IP が現 IP と一致なら sticky SKIP（reassignment なし）
8. assignment（既存 IP/Route 削除 → `New-NetIPAddress` + DNS 設定）
9. DAD 検証（500ms 待機 → AddressState 確認、Duplicate なら roll back + GUI 再表示）
10. Step 5.5: Post-Apply Verification（IP / Prefix / Gateway / DNS 読み戻し + Gateway へ
    `Test-Connection` 1 発）
11. `New-ModuleResult -Verified $verified` 返却

## 注意点・運用メモ
- 管理者権限必須
- AutoPilot 実行中でも GUI ダイアログは必ず表示（自動選択しない）= 設計判断
- 事前 probe（ICMP/ARP）は honest design として行わない:
  「真新しい kitting PC は IP 未取得 → ICMP/ARP 送れない」「切断中 PC は応答せず FREE 誤判定」
  といった限界があるため、信頼できない probe より人間の協調 + DAD + 出荷前検査の三段構えで
  運用カバーする方針
- 既存 IP/Route は assignment 時に削除（DHCP 解除含む）
- 切断中 PC が後で再接続したケースの衝突は DAD 検出不能 = 根本制約
- VERSION は 0.1.0（dev/template ベースで未リリース扱い）

## 検証
Step 5.5 で `Get-NetIPAddress` / `Get-NetRoute` / `Get-DnsClientServerAddress` で読み戻して
CSV 値と比較。AddressState=Preferred を確認。Gateway が指定されていれば `Test-Connection -Count 1`
で疎通を 1 発確認し、reachable なら `-Verified $true` を `New-ModuleResult` に渡す。
ICMP がブロックされている環境では Verified=false になりうるが、IP 設定自体は適用済み。

GUI ダイアログには Status 列に `[CURRENT]`（現 IP と一致）/ `[DUPLICATE]`（DAD 検出済）/
空欄 を表示し、`[CURRENT]` が pool 内にあれば sticky pre-select する UX。
