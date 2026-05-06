# office_license_config (Standard)

**カテゴリ**: Security
**メニュー名**: Install Office Product Key / Activate Office License
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: install=なし／auth=スクリプト内で `/dstatus` 再解析（事実上の検証、`-Verified` は未渡し）
**サブスクリプト**: `office_license_install.ps1`（プロダクトキー登録）, `office_license_auth.ps1`（ライセンス認証）

## 目的
Microsoft Office のプロダクトキー登録（`OSPP.vbs /inpkey`）と
ライセンス認証（`OSPP.vbs /act`）を行う 2 スクリプト構成のモジュール。
OSPP.vbs パスは C2R 64/32bit、MSI 64/32bit の 4 候補を優先順で自動検出し、
特殊環境向けに CSV から OsppPath を強制指定可能なハイブリッド方式です。
ENC: プレフィクスによる暗号化キーの保管に対応し、複数製品（Office 本体 + Visio +
Project）を 1 度の認証で一括処理します。

## 入力 (CSV)
`office_key.csv`:
- `Enabled`: 有効フラグ
- `ProductKey`: `XXXXX-XXXXX-XXXXX-XXXXX-XXXXX` 形式（ENC: プレフィクス対応）
- `ActivationType`: `MAK` / `KMS`（疎通チェック分岐に使用）
- `OsppPath`: OSPP.vbs フルパス（空欄=自動検出）
- `Description`, `Segment`

## 主要ステップ（install）
1. 管理者権限チェック → 失敗で Error 終了
2. `office_key.csv` 読み込み（Enabled=1 のみ）
3. OSPP.vbs 自動検出。失敗時に CSV の OsppPath が皆無なら Error 終了
4. ドライラン表示（[APPLY] / [INVALID KEY] / [ENC ERROR] / [OSPP NOT FOUND] でマーキング）
5. 実行確認（AutoPilot は自動 Y）
6. 各キーごとに ENC 残存ガード → 形式バリデーション → `cscript //Nologo OSPP.vbs /inpkey:<KEY>` 実行
7. `New-BatchResult` で集計

## 主要ステップ（auth）
1. 管理者権限チェック
2. OSPP.vbs 検出
3. CSV から ActivationType を判定し MAK の場合のみ疎通チェック
   - 第一段: `Test-NetConnection activation.sls.microsoft.com:443`
   - 失敗時: `Test-Connection 8.8.8.8` フォールバック（ping 成功なら Warning で続行＝プロキシ環境想定）
   - 両方失敗で Error 中止
4. `/dstatus` で現状ステータス取得 + 全製品の LICENSE STATUS 解析
5. 全製品が LICENSED なら冪等 Skip（複数製品ロバスト判定）
6. 実行確認 → `cscript /act` 実行
7. 3 秒待機後 `/dstatus` 再実行 → 検証 → Status / Partial / Error 判定

## 注意点・運用メモ
- 管理者権限必須
- MAK: インターネット必須（プロキシ環境は ping フォールバックで継続判断）
- KMS: 社内 KMS サーバー到達性のみで OK、180 日サイクル自動再認証
- ENC: プレフィクスのまま到達した場合は passphrase 未入力検出として Error 計上
- `office_license_install.ps1` の `Find-OsppVbs` 関数は両スクリプトでローカル定義
  （`_office_common.ps1` に切り出す案は guide にコメントあるが現状未統合）

## 検証
- install: 検証は実装せず、cscript の ExitCode のみで判定。`-Verified` は未渡し
- auth: スクリプト内で `/dstatus` を再実行し全製品の LICENSE STATUS を再解析、
  Status / Message で結果を表現するが `New-ModuleResult -Verified` は未使用。
  事実上の検証はスクリプト内で完結しており、
  Verified カラム連携が必要な場合は今後 `-Verified` 引数追加で対応可能
