# windows_license_config (Standard)

**カテゴリ**: Security
**メニュー名**: Install Product Key / Activate Windows License
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 部分実装（Activate 側に再取得確認あり、`-Verified` 引数は未使用）
**サブスクリプト**:
- `windows_license_install.ps1` … プロダクトキー投入
- `windows_license_auth.ps1` … オンライン認証

## 目的
Windows プロダクトキーのインストールとライセンス認証を 2 つのスクリプトで分離して扱うモジュール。
キー投入は `slmgr /ipk` 相当を内部 API で行い、認証は `SoftwareLicensingService.RefreshLicenseStatus`
を起動して結果を `SoftwareLicensingProduct` から再取得することで状態確認する。インターネット未接続
環境では Activate がスキップされる運用設計。

## 入力 (CSV)
`license_key.csv`（Install 側で使用）
- `Enabled`: 1=実行 / 0=スキップ
- `ProductKey`: プロダクトキー (`XXXXX-XXXXX-XXXXX-XXXXX-XXXXX`)
- `Description`: 説明（例: "Windows 11 Pro Volume License"）
- `Segment`: Segment フィルタ（任意）

CSV にキーが無い・無効なら手動入力にフォールバック。Activate 側は CSV を持たない。

## 主要ステップ
**[Install Product Key]**
1. `license_key.csv` からキー読み込み（無ければ手動入力）
2. 現在のライセンス状態を `SoftwareLicensingProduct` から取得・表示
3. 実行確認（AutoPilot 自動 Y）
4. プロダクトキー投入（`SoftwareLicensingService.InstallProductKey` 相当）
5. OS 側のエラーコードで成否判定 → `New-ModuleResult` 返却

**[Activate Windows License]**
1. `SoftwareLicensingProduct` で現状表示
2. 冪等性チェック（`LicenseStatus=1`（Licensed）なら Skip）
3. 実行確認（AutoPilot 自動 Y）
4. `SoftwareLicensingService.RefreshLicenseStatus` 起動
5. 3 秒待機 → `SoftwareLicensingProduct` 再取得 → `LicenseStatus` 確認 →
   1 なら Success、それ以外なら Error

## 注意点・運用メモ
- 管理者権限必須（両スクリプトとも）
- Activate はインターネット接続必須。プロキシ環境では `slmgr.vbs /skms` 等の事前設定が必要
- Volume License (KMS) の場合は KMS サーバ到達性が前提
- License key は CSV 上は平文（Sysprep 用 `unattend.xml` 内の Key も同様）。
  公開リポジトリで誤コミットしないよう運用注意（ENC: 対応は本モジュールでは未実装）
- Order=91/92 で並んでおり、Install → Activate の順で Profile に組み込む想定
- Activate Skip 時のメッセージで「Already Licensed」を明示するため運用ログ上は分かりやすい

## 検証
- **Install Product Key**: Post-Apply Verification は実装していない。プロダクトキーの
  インストール成否は OS 側エラーコードで判定するのみ
- **Activate Windows License**: Step 5 が事実上の Post-Apply Verification。
  `RefreshLicenseStatus` 呼び出し後 3 秒待機 → `SoftwareLicensingProduct` を再取得して
  `LicenseStatus=1` を確認する。ただし `New-ModuleResult` の `-Verified` パラメータは
  現状使用しておらず、履歴の Verified 列は空欄

将来的には `slmgr /xpr` の出力比較 + `LicenseStatus` の数値比較を `-Verified` 統合する
余地あり（`project_crypto_security_review.md` 系の改善対象に含まれる）。
