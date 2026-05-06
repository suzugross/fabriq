# ipaddress_config (Standard)

**カテゴリ**: Network
**メニュー名**: IP Address Settings
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（IP / Gateway / DNS の 3 軸読み返し）
**サブスクリプト**: `ipaddress_config.ps1`（単一スクリプト、ローカル CSV なし）

## 目的
Ethernet および Wi-Fi の IP アドレス・サブネット・デフォルトゲートウェイ・DNS を
hostlist.csv 由来の環境変数から一括設定します。
物理アダプタを `Get-NetAdapter -Physical` + `InterfaceDescription` の正規表現で
Ethernet / Wi-Fi に振り分け、netsh 経由で適用する Windows レガシースタックと
モダンスタック双方に整合性のある実装。確認プロンプトを持たず hostlist の値が
そのまま実行同意となる「データ駆動」モジュールの典型例です。

## 入力 (CSV)
ローカル CSV なし。hostlist.csv 由来の以下の環境変数を参照：
- `$env:SELECTED_KANRI_NO`, `SELECTED_OLD_PCNAME`, `SELECTED_NEW_PCNAME`: 表示用
- `$env:SELECTED_ETH_IP / SUBNET / GATEWAY`: Ethernet 設定
- `$env:SELECTED_WIFI_IP / SUBNET / GATEWAY`: Wi-Fi 設定
- `$env:SELECTED_DNS1 / DNS2 / DNS3 / DNS4`: 共通 DNS（Ethernet・Wi-Fi 両方に適用）

## 主要ステップ
1. 管理者権限チェック（`Test-AdminPrivilege`）
2. 環境変数読み込み + 設定内容表示
3. 物理アダプタ自動検出（Ethernet 用は Wi-Fi/Wireless/WLAN/802.11/Bluetooth 否定マッチ、
   Wi-Fi 用は同じ語の肯定マッチ。`InterfaceDescription` は OS locale 非依存で常に英語）
4. 実行確認は省略（hostlist 選択が同意代わり）
5. `netsh interface ip set address` / `set dns` / `add dns` で IP・GW・DNS 適用
6. Step 5.5: Post-Apply Verification（後述）
7. `New-ModuleResult -Verified` で集計返却

## 注意点・運用メモ
- IP 列が空欄のアダプタは設定スキップ（部分構成可能）
- Disabled アダプタは除外、Disconnected（ケーブル抜け）は対象に含む
- netsh はレガシー TCP/IP プロパティ GUI と同等のストアに書き込み、
  DHCP→Static 移行を自動処理
- サブネット表記は CIDR ではなく `255.255.255.0` 形式（内部で prefix 長変換）

## 検証
Post-Apply Verification は実装済み。各設定アダプタについて 3 軸を読み返し：
- IP + プレフィックス長: `Get-NetIPAddress` で完全一致
- デフォルトゲートウェイ: `Get-NetIPConfiguration` の `IPv4DefaultGateway.NextHop` 一致
- DNS サーバ: `Get-DnsClientServerAddress` の `ServerAddresses` を順序込み完全一致

全項目一致で `-Verified $true`、1 件でも不一致で `$false`、
hostlist で Ethernet/Wi-Fi 両方空欄の場合は `$null`（検証対象なし）を返却。
