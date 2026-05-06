# network_profile_config (Extended)

**カテゴリ**: Network
**メニュー名**: Network Profile Config
**VERSION**: 1.0.0  / **REQUIRES_KERNEL**: 2.0.0
**Post-Apply Verification**: 実装あり（Step 5.5 で全対象プロファイルの Category を読み返して期待値と比較、`-Verified` を返却）
**サブスクリプト**: なし

## 目的
HKLM 配下のネットワークプロファイル `NetworkList\Profiles\{GUID}` に対し、ProfileName で GUID サブキーを動的解決し `Category` 値（0=Public / 1=Private / 2=Domain）を統一的に書き込むモジュール。
マシンごとに変わる GUID を意識せず、CSV で「ProfileName + MatchMode」を指定するだけで全マッチプロファイルへ一括適用できる。
特徴的な機能は **MatchMode=All** によるベースライン + 例外運用パターン: 「まず全プロファイルを Private に」+「Guest-WiFi だけ Public に上書き」を 1 CSV で表現可能（CSV 行順で後勝ち）。

## 入力 (CSV)
`network_profile_list.csv`:
- **Enabled**: 有効フラグ
- **ProfileName**: 検索対象プロファイル名（MatchMode=All では無視）
- **Category**: `0`=Public / `1`=Private / `2`=Domain
- **MatchMode**: `Exact`（デフォルト） / `Wildcard`（PowerShell `-like` 演算子） / `All`（全プロファイル）
- **Description**: 説明
- **Segment**: Segment フィルタ

## 主要ステップ
1. `Test-AdminPrivilege` で権限チェック
2. CSV 読み込み + MatchMode 空欄を `Exact` に正規化
3. `HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\Profiles` 存在確認 + Category/MatchMode 値検証
4. ドライラン表示（各行ごとにマッチプロファイル一覧、`[APPLY]`/`[CURRENT]`/`[NOT FOUND]` 色分け）
5. `Confirm-ModuleExecution`
6. 適用ループ（各プロファイルに対し冪等性チェック → `Set-ItemProperty Category` 書き込み）。`$finalExpected` ハッシュに後勝ちで最終期待値を記録
7. **Step 5.5 Post-Apply Verification**: `$finalExpected` の各エントリを `Test-CategoryValueMatch` で読み返し検証 → `verifyPass` / `verifyFail` 集計
8. `New-BatchResult` に `-Verified $verified` を渡して返却

## 注意点・運用メモ
- 反映には再起動不要（即時反映）
- Category=2 (Domain) は通常 NLA が自動管理する領域。手動設定後にドメイン参加状況等で Windows が再上書きする可能性あり
- 同じ ProfileName を複数行で指定した場合、CSV 行順で後勝ち（`$finalExpected` がオーバーライドルールを保持して Verification も後勝ち期待値で行う）
- 新規プロファイルの作成は行わず、既存プロファイルの Category 値変更のみ
- CSV は UTF-8 BOM 付き必須（日本語 Description / ProfileName 対応）
- SetupComplete.cmd 経由で OOBE 直後実行も可能（Guide.txt に記載例あり）

## 検証
**Post-Apply Verification 実装の好例モジュール**。Step 5.5 で書き込み後の Category を全対象プロファイルで読み返し、後勝ちルールを適用した `$finalExpected` と比較。失敗 0 件で `$verified=$true`、対象 0 件で `$null`、失敗ありで `$false` を返却。
