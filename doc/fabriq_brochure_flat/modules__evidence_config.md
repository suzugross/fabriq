# evidence_config (Standard)

**カテゴリ**: Evidence
**メニュー名**: Collect Evidence
**VERSION**: 1.6.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 概念的に不適用（読み出し系のためシステム状態を変更しない）
**サブスクリプト**: なし

## 目的
PC のシステム情報を **22+ セクション** にわたって収集し、テキスト／CSV／HTML のエビデンスファイル一式を `evidence/pc_information/<日付>_<UID>_<PC名>/` に書き出す、fabriq の納品エビデンス採取の中核モジュールです。受入検査・官公庁監査・トラブル時の証跡として機能し、収集完了時に **`manifest.json`** を生成して各セクションの `id / title / files / status / reason / elapsedMs` を機械可読サマリ化します。これにより外部ツール（fabriq_evidence_manager 等）が manifest を起点にエビデンスを統合できる契約構造になっています。シリアル番号は 4 ソース横断 + canonical 選定 + OEM 無効値（"Default string" 等）の拒否までやる多重ソース戦略。

## 入力 (CSV)
**設定 CSV なし**。すべてのセクションが固定で実行されます。

## 収集セクション（抜粋）
1. システム基本情報 / 2. ローカル管理者 / 3. ネットワーク / 4. プリンター / 5. BitLocker / 6. MAC アドレス
7. PC シリアル番号（多重ソース + canonical 選定）／ 10. シリアル番号別ファイル
8. デスクトップアプリ / ストアアプリ CSV
9. ファイアウォール（プロファイル + ルール）CSV
10/11. Windows 機能 / Server Roles & Features CSV
20. System TEMP テキストログバックアップ（セーフティネット）
21. Windows ライセンス（SoftwareLicensingProduct + slmgr /dlv）
22. Office ライセンス（OSPP + vNext per-user スキャン + 自動解釈 verdict）
23. Security Baseline（TPM / Secure Boot / VBS / HVCI / Credential Guard / LSA Protection / BIOS）
24. Group Policy Report（`gpresult /h` HTML + サマリ TXT）
25. Certificates CSV（4 ストア統合、HasPrivateKey フラグのみ）
26. Battery Report（`powercfg /batteryreport` HTML、受入検査での容量契約エビデンス）

## 主要ステップ
1. 出力先パス決定（`$global:FabriqEvidenceBasePath` 有無で統一パスモード／レガシーモード分岐）
2. 実行確認（AutoPilot 時は自動 Y）
3. 各セクションを順次収集
4. 統合ファイル `_ALL_<PC名>_Log.txt` + 個別ファイル + CSV/HTML 出力
5. `manifest.json` 生成（`schemaVersion=1`、再実行時は `manifest.json.bak` に rotate）

## 注意点・運用メモ
- **管理者権限必須**（BitLocker / Firewall ルール / Defender 状態など admin 限定情報を含む）
- §22d の自動解釈は M365 sub の OSPP「飾りキー」問題（OSPP は常に Grace 表示になる）を内部で吸収し、Manifest の `Partial` / `Failed` 区別を出力
- §23 の各 probe は inner try/catch で個別退避（1 つ失敗してもセクションは Success）
- §24 の `gpresult` ユーザー側 RSoP は実行ユーザー（kitting admin01 等）視点であり、最終エンドユーザー視点ではない点に注意
- §26 はバッテリ非搭載で Skipped 完結
- 環境変数: `$env:SELECTED_NEW_PCNAME` / `$env:COMPUTERNAME`、グローバル `$global:FabriqUniqueId` / `$global:FabriqEvidenceBasePath`

## 検証
Post-Apply Verification は **概念的に不適用**。本モジュールは「読み出し系」であり、システム状態を変更しないため「適用後の読み返し検証」が論理的に存在しません。`-Verified` は未渡しで Verified 列は空欄。代わりに `manifest.json` のセクション単位 `status` フィールド（Success / Skipped / Failed / Partial の 4 値）が品質情報を担います。`kernel/EVIDENCE_MANIFEST.md` で公式契約として `schemaVersion=1` が定義されており、外部 evidence consumer はこの manifest を起点にエビデンスをパースする設計です。
