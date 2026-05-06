# cert_config (Standard)

**カテゴリ**: Security
**メニュー名**: Certificate Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（サムプリントによるストア再検索）
**サブスクリプト**: なし（証明書実体は `certs/` 配下に配置）

## 目的
証明書ファイル（`.pfx` / `.p12` / `.cer` / `.crt`）を Windows の証明書ストアにインポートするモジュールです。**明示指定モード** （CSV でストアスコープ／ストア名を直接指定）と **AutoRoute モード** （PFX 内の各証明書を BasicConstraints CA フラグで判定し、CA は `LocalMachine\Root`、クライアントは `LocalMachine\My` に自動振り分け）の 2 動作を持ちます。FileName は環境変数 + ワイルドカード対応のため、`%SELECTED_NEW_PCNAME%_*.p12` のように書けば `hostlist.csv` ベースで「PC ごとに異なるクライアント証明書」を自動配布する運用が可能です。PFX パスワードは ENC: プレフィックス暗号化対応。

## 入力 (CSV)
`cert_list.csv` の主な列:
- `Enabled` … 1=実行 / 0=スキップ
- `FileName` … `certs/` 内のファイル名（環境変数 / ワイルドカード可）
- `StoreScope` … `LocalMachine` / `CurrentUser`（AutoRoute=1 時は空欄可）
- `StoreName` … `My` / `Root` / `CA` / `TrustedPublisher` 等（AutoRoute=1 時は空欄可）
- `Password` … PFX パスワード（**`ENC:` プレフィックス対応**、`.cer`/`.crt` は空欄）
- `AutoRoute` … 1=自動振り分け / 0=明示指定
- `Overwrite` … 1=置換 / 0=既存スキップ
- `FriendlyName` … クライアント証明書のフレンドリ名
- `Description` / `Segment`

## 主要ステップ
1. `cert_list.csv` 読み込み
2. `certs/` ディレクトリ存在確認
3. 各証明書の状態確認 + 対象一覧表示（`[IMPORT]`/`[SKIP]`/`[REPLACE]`）
4. 実行確認（AutoPilot 時は自動 Y）
5. インポート（明示指定 or AutoRoute で振り分け）
5.5 **Post-Apply Verification**: `Test-CertificateInStore` で `X509Store.Certificates` をサムプリント検索
6. `New-BatchResult ... -Verified $verified` で返却

## 注意点・運用メモ
- **管理者権限必須**（`LocalMachine` ストア書き込み。明示チェックなしで失敗時に検出）
- AutoRoute モードでは複数証明書（ルート CA + 中間 CA + クライアント）を含む PFX も適切に振り分け
- `%SELECTED_NEW_PCNAME%` + ワイルドカードで PC 別証明書を自動マッチ可能
- ENC: 復号は `Import-ModuleCsv` 側で自動実施（マスターパスフレーズ要）
- CSV の文字エンコーディングは Shift-JIS / UTF-8 BOM 付き両対応（fabriq 共通）
- 環境変数: `$env:SELECTED_NEW_PCNAME` / `$env:SELECTED_SEGMENT`

## 検証
Post-Apply Verification は **実装あり**。Step 5.5 で実際にインポート／既存維持した全証明書を `Test-CertificateInStore` でサムプリント検索し、ストア内に存在することを確認します。AutoRoute モードでは PFX 内の各証明書を個別に検証。全件成功で Verified=true、1 件でも失敗すれば false。`-Verified` フラグ付きで `New-BatchResult` に返却するため、Evidence Manager 側で「証明書配布完了」を突合できます。
