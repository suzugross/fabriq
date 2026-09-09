# プロファイル別データオーバーレイ 実装計画書

Status: **完了(レビュー待ち)** — Phase 1〜3 実装済み・VM E2E 済み / **Phase 4 は全項目が裁定により不実施**
(TM: t-0095【最重要】/ P1 = t-0097 / P2 = t-0098 / P3 = t-0100)
裁定: **Q1〜Q6 全件裁定済み(残裁定なし)** — Q1 = (a) メニュー単発は無効 / Q2 = 見送りで確定 /
Q3 = per-case / Q4 = Show-Warning / Q5 = strict mode 不採用 / Q6 = 役割分担 + Studio 起票(§9・§14)
作成: 2026-08-09 / 最終更新: 2026-09-09

---

## 1. 目的と設計原則

1 つのカーネル配備で複数の設定セット(顧客・案件・ロット別)を使い分けるため、
モジュール CSV・資材フォルダをプロファイル側フォルダに集約する。

### 原則: Profile-First(本体 CSV は非常用)

- **プロファイル実行時のデータの正はプロファイル側フォルダ**とする。
- モジュール本体側 CSV へのフォールバックは**存在するが非推奨**。
  「プロファイルに書かなければ本体のデフォルトが効く」という積極利用は**想定しない**。
  フォールバックは移行期・例外ケースの救済措置と位置づける。
- したがってフォールバックは**必ず可視化**する(無言フォールバック禁止)。
  可視化の手段は Show-Warning + テレメトリ `resolvedFrom` + HTML チェックリストの `Data Set` の 3 点で、
  **これを最終形とする**(strict mode = フォールバックの Error 化は 2026-09-09 に不採用裁定。Q5)。
- オーバーレイ不在(従来運用・メニュー単発実行)では**現状と完全同一の動作**を保証する。

### 背景となる実測インベントリ(2026-08-09 時点)

| 対象 | 実数 |
|---|---|
| Import-ModuleCsv 経由の CSV 読込 | 88 ファイル / 100 呼出 |
| CSV 読込前の Test-Path 前置き | 8 箇所(§7.1) |
| 資材フォルダ参照モジュール | 約 14(§7.2) |
| 多 CSV 列挙(`reg_*_list*.csv`) | 4 スクリプト(§7.3) |
| モジュール dir への書き込み(backup 系) | 8 スクリプト(§7.4) |
| クロスモジュール参照 | 2 件(§7.5) |
| hostlist 消費者 | kernel main + csv_editor + fabriq_ios + 衛星 2 リポジトリ |

---

## 2. 用語

| 用語 | 意味 |
|---|---|
| **プロファイルデータフォルダ (PDF: Profile Data Folder)** | `profiles/<name>/` — プロファイル `profiles/<name>.csv` に併設するオーバーレイルート |
| **オーバーレイ解決** | モジュールが読むデータパスを PDF 側 → 本体側の順で解決すること |
| **フォールバック** | PDF に該当ファイル/フォルダが無く本体側が使われること(非推奨・要可視化) |
| **データコンテキスト** | 「いまどの PDF が有効か」を表すランタイム状態(env + global) |

---

## 3. ディレクトリ構造仕様

```
profiles/
  Master_Config01.csv          ← プロファイル定義(現行のまま・identity 不変)
  Master_Config01/             ← PDF(新設)
    modules/
      driver_config/
        driver.csv             ← モジュール CSV(ファイル名は本体と同一)
        driver/                ← 資材フォルダ(フォルダ名は本体と同一)
      windows_license_config/
        license_key.csv
      reg_hklm_config/
        reg_hklm_list01.csv    ← 列挙系はモジュール単位 all-or-nothing(§4.4)
      ...
```

- **プロファイル定義は単一 CSV のまま**。フォルダ化しない。
  理由: Frex の Order 一級原則・operator アプリ・既存運用に非接触で導入できる。
- PDF 内の相対構造は `modules/<module名>/<本体と同じ相対パス>` の**完全ミラー**。
  写像規則が機械的になり、ツール(csv_editor 等)の追随も単純化する。
- `module.csv` / `preset.csv` / `VERSION` / `REQUIRES_KERNEL` / `Guide.txt` / `test.psd1` は
  **フレームワーク資産でありオーバーレイ対象外**(PDF に置かれても無視)。
- **PDF ルート直下(`modules/` 以外)の名前空間はツーリング用に予約**する。
  カーネルは `modules/` 配下しか見ない。FabriqStudio 等が案件メタデータ(顧客名・作成日等)を
  PDF ルートに置けるようにするための前方互換予約(予約は維持。ポリシーマーカーは Q5 不採用により不使用)。

---

## 4. 解決契約(正式仕様)

### 4.1 データコンテキストのライフサイクル

| タイミング | 動作 |
|---|---|
| プロファイル実行開始 | `profiles/<name>/` が存在すれば `FABRIQ_PROFILE_DATA_DIR` に絶対パスを設定。存在しなければ未設定(= 従来動作)で、その旨を 1 行表示 |
| プロファイル実行終了(finalize / Cancel / Error 終了) | 必ずクリア(finally 相当の位置で) |
| `__RESTART__` 跨ぎ | `Save-ResumeState` が `ProfileDataDir` を保存(空 = PDF 無し)。resume 二段目の直前に `Test-FabriqResumeDataDir` で「保存値が非空なのに現在フォルダが無い」を検査し、該当時は Show-Error して**残モジュールを実行しない**(state ファイルは温存 = フォルダ復旧後に再 resume 可能)。二段目のコンテキスト自体は `ProfilePath` から再導出される |
| メニュー単発実行 | コンテキスト無し = オーバーレイ無効 → **モジュール本体の CSV を使う**。Q1 裁定 (a)、2026-09-09 に**恒久確定**(データセット選択 UI は入れない。§14.2) |
| async(Runspace)実行 | env はプロセス共有のため追加対応不要(現行 FABRIQ_SEGMENT と同じ扱い) |
| WU 再起動ループ(`Invoke-WindowsUpdateLoop`) | 呼出元はメインメニュー操作と起動時 resume の 2 箇所のみ(プロファイルバッチからは起動されない = **メニューモード**)。Q1 (a) によりコンテキスト無しが正しい動作。P1 の「context unavailable」表示は事実の明示で、復元処理は不要(R11 は設計上クローズ)。プロファイル行として `windows_update.ps1` を置いた単発パスはバッチ内で実行されるため通常通り解決される |

### 4.2 パス解決アルゴリズム

公開 API `Resolve-ModuleDataPath -Path <absolute path>`(名称仮):

1. コンテキスト未設定 → 入力パスをそのまま返す(**恒等写像**。従来運用の完全互換)。
2. 入力パスが `<repo>\modules\<tier>\<module>\<rel>` の形でなければそのまま返す
   (モジュール外・評価不能パスには一切干渉しない)。
3. `<PDF>\modules\<module>\<rel>` が存在すればそれを返す(**採用元 = PROFILE**)。
4. 存在しなければ入力パスを返す(**採用元 = FALLBACK**)。

- 判定は「ファイル存在」のみ。中身の検査・マージは行わない(**ファイル単位 all-or-nothing**)。
- `<tier>`(standard/extended)は PDF 側パスに**含めない**(モジュール名は tier 間で一意)。
- 大文字小文字は Windows 準拠(不問)。`..\` を含む正規化は解決前に `GetFullPath` で行う。

### 4.3 可視化契約(Profile-First の中核)

プロファイル実行中の全データ解決について、読込直前に**採用元を必ず 1 行表示**する:

- PROFILE 採用: `[DATA] license_key.csv <- profile (Master_Config01)`(通常色)
- フォールバック: `[DATA] license_key.csv <- module dir (FALLBACK)` を **Show-Warning** で表示
  - 非推奨運用の顕在化が目的。移行期のノイズより「気づかない」ことのほうが害が大きい、
    という原則決定(2026-08-09 ユーザー方針)。
- 実行履歴/エビデンス(HTML チェックリスト)にモジュール毎の採用元サマリを記録する
  (誤設定セット出荷の事後検出手段)。

### 4.4 列挙系(reg_hklm / reg_hkcu)の特例

`Get-ChildItem -Filter "reg_*_list*.csv"` 型の列挙は**モジュール単位 all-or-nothing**:

- PDF 側に当該モジュールのフォルダが存在し、パターン一致ファイルが 1 つ以上ある
  → **PDF 側のみ**を列挙(本体側とのファイル合成はしない)。
- 無い → 本体側を列挙(= フォールバック。§4.3 の Warning 対象)。
- 理由: ファイル単位の部分マージは「どの行がどこから来たか」を操作者が追えなくなる。

### 4.5 資材フォルダの特例

driver/ certs/ xml/ file/ 等の資材は**フォルダ単位 all-or-nothing**(§4.4 と同じ理由)。
PDF 側にフォルダが存在すれば(空でも)PDF 側を使う。空フォルダは「資材ゼロ」の明示と解釈する。

### 4.6 書き込み系(backup / export)

- **コンテキストが有効な間の書き込み先は PDF 側**(`<PDF>\modules\<module>\backup\` 等)。
  backup→restore の対称性が PDF 内で閉じる(顧客 A のバックアップが base や顧客 B に混ざらない)。
- コンテキスト無しでは従来通り本体側。
- 実装は Phase 3。**Phase 1〜2 の間は書き込みは常に本体側**と契約に明記(過渡期仕様)。

### 4.7 直交性

- **Segment フィルタ**: 解決後の CSV に対して従来通り適用(オーバーレイと直交)。
- **ENC: 透過復号**: パス非依存のため無影響。PDF 側 CSV でも同様に機能する。
- **Deploy / framework patch**: patch は `profiles/` を保全する既存ルール
  (framework_overlay_rules.json)のため本構想と整合。site データが PDF に寄るほど
  patch 運用は単純化する。

---

## 5. 公開 API・カーネル変更案(Phase 1 時点)

| 項目 | 内容 | KERNEL_API.md |
|---|---|---|
| `Resolve-ModuleDataPath -Path <string>` | §4.2 の解決 + §4.3 の表示(表示は 1 実行 1 ファイル 1 回に抑制)。P1 実装済み | §1.2 追記済み(MINOR) |
| `FABRIQ_PROFILE_DATA_DIR` | データコンテキスト env。P1 実装済み | §3.2 追記済み |
| `Import-ModuleCsv` 内部フック | 受領パスを `Resolve-ModuleDataPath` に通す(88 ファイル/100 呼出を無改修で吸収)。csv.load テレメトリに `resolvedFrom` を付与。P1 実装済み | 契約文に解決順を追記済み |
| ResumeState 拡張 | `ProfileDataDir` フィールド追加(v1 追加フィールド)。P1 実装済み | — |
| 内部ヘルパ(§6 非公開) | `Get-FabriqProfileDataDir`(profile CSV → PDF 導出)/ `Set-` `Clear-FabriqProfileDataContext`(env + 表示 dedup リセット)/ `Test-FabriqResumeDataDir`(fail-closed 判定)。`Reset-FabriqState` の env クリア対象に `FABRIQ_PROFILE_DATA_DIR` を追加 | — |
| main.ps1 | `Invoke-BatchExecution` 冒頭でコンテキスト導出・設定・バナー、finally で解除。resume 二段目(Linear / Flex)直前に fail-closed 判定。WU ループに context 外表示。P1 実装済み | — |

- バージョン: **kernel MINOR(3.6.2 据置、次リリースで 3.7.0 に New-UiFont 分と合流)**。
- モジュール側の改修が必要なのは §7.1 の 13 箇所(Phase 1・実施済み)と §7.2 以降(Phase 2+)のみ。

---

## 6. フェーズ計画

### Phase 0: 契約凍結(1〜2 日)

- 本書 §3〜§4 をレビューし未決事項(§9)を裁定 → 契約凍結。
- 設計ゲート(フル版)提出: コンテキストのステートマシン + 敵対検証。
- **完了条件**: 本書の Status を「契約凍結」に更新。

### Phase 1: CSV 読みオーバーレイ(3〜5 日)

1. kernel: `Resolve-ModuleDataPath` + コンテキスト管理 + Import-ModuleCsv フック + 表示。
2. kernel: ResumeState 拡張(保存・復元・復元失敗 = Error)。
3. modules: Test-Path 前置き 13 箇所(§7.1)を解決 API 経由に修正(各 MINOR + REQUIRES_KERNEL 3.7.0)。
4. tests: 新規 Pester(解決順・恒等写像・restart 復元・表示抑制・Segment/ENC 直交)。
5. VM リグ E2E: PDF あり/なし/フォールバック混在プロファイルの 3 本。
- **完了条件**: run_tests 全緑 + VM E2E 3 本 + 「オーバーレイ不在で既存挙動と完全同一」の回帰確認。
- **ロールバック**: コンテキストを設定しなければ全コードパスが恒等写像 = 機能フラグ不要で即時無効化可能。
- **実施記録(2026-09-09)**: 1〜4 完了。run_tests **437 / 437 PASS**(既存 413 + 新規 24:
  ProfileDataOverlay 18 / ResumeState 1 / Invoke-BatchExecution 5)、パーサ 15 ファイル 0 エラー、
  check_ps1_encoding exit 0。VM リグ E2E(5)は未実施 — レビュー後に実施可否を判断。
  実装中の追加発見: Test-Path 前置きは 13 箇所(§7.1)、`GetFullPath` は 8.3 短縮名を展開するため
  テスト fixture 側も正規化が必要(TEMP が `SZK-WI~1` 形式の環境)。

### Phase 2: 資材フォルダ・列挙系(1〜1.5 週)

- **詳細設計は §12(着手可能な粒度で確定済み)**。wave 構成: W0 kernel(ディレクトリ解決 +
  列挙ヘルパ)→ W1 純読み取り資材 10 モジュール → W2 列挙系 reg ×4 → W3 パイプライン族
  (startlayout ×3 / sysprep source + taskbar クロス書き / printer_driver INF)→ W4 ツーリング
  (csv_editor / fabriq_ios)。
- 各 wave: 軽量版ゲート(W0 のみフル版)+ モジュール VERSION MINOR + REQUIRES_KERNEL 3.7.0 + 個別検証。
- **実施済み(2026-09-09)**: W0〜W3 を実装(kernel 1 + モジュール 16 件)。W4 はユーザー裁定で縮小
  — csv_editor は対象外(撤去候補)、fabriq_ios は追加実装なしで「従来どおり動く」ことを回帰検証。
  実施記録は §12.1 / §12.4 末尾 / §12.5。

### Phase 3: 書き込み系・クロスモジュール(1 週)— **実施済み(2026-09-09)**

- §7.4 の backup/restore 4 対 + driver_export / default_app export の計 6 対に §4.6 を実装
  (restore は「PDF に無ければ本体 backup を Warning 付きで参照」のフォールバック対称)。
- taskbar_config → sysprep_config/source の書き込み(§7.5)は **W3 で先行実施済み**。
- 公開 API `Resolve-ModuleDataPath -ForWrite` を追加(kernel MINOR)。実施記録は §13.4。

### Phase 4: 締め・運用移行 — **全項目が裁定により終了(2026-09-09)**

実装すべき項目は残っていない(§14)。

- ~~strict mode~~ → **不採用**(Q5 裁定)。可視化 3 点で十分と判断。
- ~~hostlist の per-profile 化~~ → **見送りで確定**(Q2 裁定)。
- ~~メニュー単発のデータセット選択~~ → **不採用**(Q1 恒久確定)。メニュー単発は常にモジュール
  デフォルト CSV を使う。
- ~~本体 CSV のサンプル縮退~~ → **不実施**。積極的な削除はせず、PDF 運用の定着に伴う自然消滅に委ねる。

→ **本計画書の実装スコープは Phase 1〜3 で完了。**

---

## 7. 改修対象インベントリ(実測・file:line)

### 7.1 Test-Path 前置き(Phase 1 で解決 API 経由に修正済み・13 箇所 / 9 ファイル / 8 モジュール)

放置すると「本体に無く PDF にある」構成で無言スキップ(2026-08 の license バグと同型)。
実装時の再 grep(`Test-Path` × CSV 系変数名すべて)で当初の 8 箇所から 13 箇所に増えた。
すべて `$x = Resolve-ModuleDataPath -Path (Join-Path $PSScriptRoot "...")` に統一済み。

| # | ファイル(変数) | 状態 |
|---|---|---|
| 1 | firewall_config/firewall_config.ps1(`$csvPath`) | 済 |
| 2 | firewall_rule_config/firewall_rule_export.ps1(`$csvPath` — Add-ImportEntry の書き戻し先も同パス) | 済 |
| 3-4 | local_user_config/local_user_config.ps1(`$csvPath` / `$hostCsvPath`) | 済 |
| 5-6 | local_user_config/local_user_delete.ps1(`$csvPath` / `$hostCsvPath`) | 済 |
| 7 | office_license_config/office_license_auth.ps1(`$csvPath`) | 済 |
| 8-9 | printer_delete/printer_delete.ps1(`$printerListCsv` クロスモジュール `..\` / `$csvPath`) | 済 |
| 10 | printer_driver_config/printer_config.ps1(`$csvPath`) | 済 |
| 11 | printer_driver_config/printer_driver_install.ps1(`$driverCsvPath`) | 済 |
| 12 | windows_license_config/windows_license_install.ps1(`$csvPath`) | 済 |
| 13 | extended/history_destroyer/history_destroyer.ps1(`$ssidCsvPath`) | 済 |

### 7.2 資材フォルダ読み(Phase 2)

| モジュール | フォルダ | 箇所 |
|---|---|---|
| driver_config | driver/ | driver_import_config.ps1:38(export は Phase 3) |
| cert_config | certs/ | cert_config.ps1:108 |
| odt_config | assets/ | odt_install.ps1:54 |
| default_app_config | xml/ | default_app_config.ps1:38 |
| app_config | file/ | app_config.ps1:31 |
| copyfile_config | source/ | copyfile_config.ps1:35 |
| wallpaper_config | wallpaper/ | wallpaper_config.ps1:237 |
| ppkg_config | file/ | ppkg_install_config.ps1:44 |
| printer_driver_config | INF/ | printer_driver_install.ps1:10(tools/7z.exe はフレームワーク資産・対象外) |
| startlayout_config | json/ xml/ ppkg/ | backup:47, build:83/98/99, import:44 |
| sysprep_config | source/ | sysprep_config.ps1:26 |
| manual_kitting_assistant | prompt/ | manual_kitting_assistant.ps1:99 |
| pianist | profiles/ | pianist.ps1:147(per-case 裁定済み 2026-08-09 → PDF 対象) |

### 7.3 多 CSV 列挙(Phase 2・§4.4 適用)

- reg_hklm_config/reg_hklm_config.ps1:14 / reg_hklm_delete.ps1:12
- reg_hkcu_config/reg_hkcu_config.ps1:24 / reg_hkcu_delete.ps1:21

### 7.4 書き込み系 backup/restore 対(Phase 3・§4.6 適用)

- acl_config: acl_backup.ps1:80 / acl_restore.ps1:76
- reg_template: reg_backup.ps1:23 / reg_import.ps1:25
- firewall_rule_config: firewall_rule_export.ps1:186 / firewall_rule_import.ps1:80
- desktop_icon_config: desktop_icon_backup.ps1:47 / desktop_icon_restore.ps1:19

### 7.5 クロスモジュール参照

- taskbar_config → `..\sysprep_config\source`(**書き**。taskbar_config.ps1:41)
- printer_delete → `..\printer_driver_config\printer_list.csv`(読み。解決 API が §4.2 で吸収)

---

## 8. リスク台帳と対策

| # | リスク | 重大度 | 対策(実装済み要件) |
|---|---|---|---|
| R1 | 誤った設定セットの無言適用(PDF 効いている/いないの取り違え) | **最重大** | §4.3 全読込の採用元表示 + エビデンス記録 + フォールバック Warning |
| R2 | `__RESTART__` 跨ぎでコンテキスト消失 → 再起動後だけ base 適用 | 高 | ResumeState 保存/復元 + 復元失敗 Error(§4.1)+ 専用テスト |
| R3 | Test-Path 前置きの取りこぼし(無言フォールバック同型) | 高 | §7.1 の 8 箇所を Phase 1 内で必須修正・grep 網羅を完了条件に |
| R4 | 列挙・資材の部分マージ混乱 | 中 | §4.4/§4.5 all-or-nothing 契約 |
| R5 | 書き込みが base を汚す / backup 迷子 | 中 | Phase 1〜2 は「書き込み常に本体側」と明記、Phase 3 で §4.6 |
| R6 | パス写像のエッジ(`..\`、`\\?\`、tier 判定) | 中 | GetFullPath 正規化 + 評価不能パスは恒等写像(§4.2-2)+ 単体テスト |
| R7 | 衛星(checksheet/backuper)の hostlist 自動発見が壊れる | 中 | hostlist は本計画から分離(§9)。動かすまで波及ゼロ |
| R8 | メニュー単発実行との混同 | 中 | メニュー = オーバーレイ無効を契約化 + 画面でコンテキスト非表示なら無効と分かる表示 |
| R9 | Segment / ENC との干渉 | 低 | 直交(§4.7)。テストで固定 |
| R10 | run_tests / CI の回帰 | 低 | 恒等写像デフォルトのため既存テスト無風のはず。全 Phase で run_tests 必須 |
| R11 | WU 再起動ループのレグがコンテキスト外で `windows_update_list.csv` を本体側から読む | 低(クローズ) | 調査の結果、WU ループはメニュー操作からのみ起動される(プロファイルバッチ外)= Q1 (a) の通りコンテキスト無しが正。P1 の明示表示のみで完結、復元処理は不要 |

---

## 9. 未決事項 → **全件裁定済み(2026-09-09)**

| # | 論点 | 選択肢 | 裁定 |
|---|---|---|---|
| Q1 | メニュー単発実行でのデータコンテキスト | (a) 常に無効(本体側のみ) / (b) 選択 UI を追加 | **裁定済み(2026-09-09): (a) で恒久確定**。再判断も行い「メニュー単発はモジュールデフォルト CSV を使う」で確定した(§14.2)。選択 UI は入れない |
| Q2 | hostlist の per-profile 化 | (a) 見送り(全案件共通のまま) / (b) Phase 4 で PDF へ | **裁定済み(2026-09-09): (a) 見送りで確定**(保留ではなく結論)。根拠: ①衛星 2 本(checksheet `common.ps1:1062` / backuper `common.ps1:753`)は `kernel\csv\hostlist.csv` の**存在によって fabriq ルートを発見**しており、さらに backuper は extended_hostlist と本家 hostlist の (OldPCname, NewPCname) **集合完全一致ゲート**を持つ。衛星は単独起動(別 PC のこともある)のため「実行中プロファイル」を原理的に知り得ない。②lifecycle が違う — hostlist は**案件ごとに作り直すジョブ入力**、PDF は**顧客ごとに安定するレシピ**。ジョブ入力をレシピフォルダに入れると、案件のたびに PDF を書き換えることになる |
| Q3 | pianist の profiles/(UI 操作プロファイル) | per-case 資材か framework 資産か | **裁定済み(2026-08-09): per-case 資材 = PDF 対象**(Studio の Pianist Profile Editor が案件コンテンツとして編集している実態とも整合) |
| Q4 | フォールバック Warning の表示強度 | Show-Warning / Show-Info | **裁定済み(2026-09-09): Show-Warning**(profile-first 原則の担保) |
| Q5 | strict mode(フォールバックの Error 化) | マーカーファイル / プロファイル CSV 列 / 不採用 | **裁定済み(2026-09-09): 不採用**。指定方法を選ぶ以前に機能自体を入れない。フォールバックの可視化は Show-Warning + テレメトリ `resolvedFrom` + HTML チェックリスト `Data Set` の 3 点で**最終形**とする(§1 の原則も更新済み) |
| Q6 | **FabriqStudio のワークスペースモデルとの関係** | (a) 役割分担(ワークスペース = 物理分離、PDF = 同一配備内の設定セット切替) / (b) ワークスペース切替を PDF 切替に統合 | **裁定済み(2026-09-09): (a) + Studio 側タスクを起票**。csv_editor 撤去(§12.5)により **PDF 側 CSV を編集するツールが現状ゼロ**のため優先度が上がった。Studio 側の起票内容は §14.3 参照 |

### FabriqStudio との関係(2026-08-09 調査)

Studio(E:\fabriq_studio, WPF/.NET8)には本構想と交差する機能が既にある:

- **ワークスペース切替** = 現行の複数顧客対応(fabriq 丸ごとコピー切替)。本構想は同じ問題への
  別解。**Q6 裁定(2026-09-09)= 役割分担**: ワークスペース = 物理分離が要るとき、PDF = 同一配備内の
  設定セット切替。統合はしない(§14.3)。
- **モジュール CSV への直接書き込み**: 端末管理 / レジストリ辞書エクスポート / INF→hostlist 転記。
  データが PDF に移ると Studio の書き込み先が変わる → Studio 側改修は本計画の**非スコープ**(Studio の
  TM に別途起票)。Studio は `IWorkspaceService.RootPath` + ルート相対パスで CSV を読み書きするため
  差し込み口はほぼ 1 箇所だが、PDF は `<tier>` を落とすので**「ルート挿し替え」では足りず §4.2 の
  写像関数のミラーが要る**(§14.3 に実装形を記載)。
- **Pianist Profile Editor**: `modules/extended/pianist/profiles/` を案件コンテンツとして編集
  → Q3 は per-case 側(PDF 対象)に倒す根拠。
- **fabriq オーバーレイ更新**(SemVer 比較付きコード上書き): データが PDF に分離されるほど
  除外ルールが単純化する強シナジー。フレームワーク更新戦略の保留(コード/データ分離待ち)は
  本構想の実現で実質解消する。

---

## 10. 非スコープ(明示)

- プロファイル定義 CSV のフォルダ化・スキーマ変更(identity は現行のまま)
- module.csv / preset.csv / Guide.txt 等フレームワーク資産のオーバーレイ
- NextGProfile 的な実行制御の変更(線形 Profile 維持の既決事項に非接触)
- 衛星リポジトリ(fabriq_checksheet / fabriq_backuper / evidence_manager)の改修
  (Q2 = 見送り確定により**恒久的に発生しない**)
- FabriqStudio の PDF 対応実装(Q6 裁定により Studio 側 TM へ起票済み。
  本計画は Studio が追随可能な契約構造 — ミラー構造 + 写像規則の明文化 — の担保まで)
- evidence/ ツリーの per-profile 化(実行履歴・エビデンスは従来通り全体共有)

---

## 11. 進行管理

- TM: t-0095(本計画全体)。Phase 着手ごとに子タスクを切る(P1 = t-0097)。
- 各 Phase の実装は CLAUDE.md 設計ゲート(P1 / P2-W0 / P3 はフル版、その他 wave は軽量版)を通す。
- kernel touched の全コミットで `powershell.exe -File ./dev/run_tests.ps1` を必須実行。
- 本書は実装と同コミットで随時更新し、契約変更は必ず §4 に反映してから実装する。

---

## 12. Phase 2 詳細設計(着手可能・2026-09-09 確定)

### 12.1 W0: kernel 拡張(フル版ゲート対象)

**A. `Resolve-ModuleDataPath` のディレクトリ対応(既存 API の後方互換な拡張)**

- 現行(P1)は候補が**ファイル**として存在する時だけ PDF 側を返す。P2 で候補が**ディレクトリ**として
  存在する場合も PDF 側を返す(§4.5 フォルダ単位 all-or-nothing。空フォルダ = 資材ゼロの明示)。
- 表示ラベルは末尾 `/` 付き(`[DATA] driver_config/driver/ <- profile (X)`)。dedup キーは同じ流儀。
- 判定順: `Test-Path -PathType Leaf` → `-PathType Container` → フォールバック。
- 副作用ゼロ: P1 で恒等写像だった「ディレクトリ候補」が解決されるようになるだけで、
  コンテキスト無し・PDF にフォルダ無しの挙動は不変。既存 Pester に「ディレクトリ候補」3 ケースを追加。

**B. 列挙ヘルパ(新公開 API、§1.2 追記・MINOR)**

```
Get-ModuleDataFiles -Directory <string> -Filter <string>   (FileInfo[] 返却、Name 昇順)
```

- `Directory`(モジュール dir の絶対パス)を §4.2 の写像で PDF 側に写す。PDF 側のモジュールフォルダに
  `Filter` 一致が **1 件以上あれば PDF 側の一致のみ**を返す(本体側と合成しない = §4.4)。
- 一致ゼロまたはフォルダ無し → 本体側を列挙してフォールバック(コンテキスト有効時は Show-Warning、
  ラベル例 `[DATA] reg_hklm_config/reg_hklm_list*.csv <- module dir (FALLBACK)`)。
- コンテキスト無し → 本体側を列挙(表示なし)。0 件時の扱いは呼出側の既存ロジック(Error)に委ねる。
- 返す FileInfo の `FullName` は PDF 側パスなので、続く `Import-ModuleCsv` は恒等写像(冪等)。

**C. テスト(W0 完了条件)**

- `ProfileDataOverlay.tests.ps1` に追加: ディレクトリ候補あり/なし/空フォルダ、`Get-ModuleDataFiles` の
  PDF 優先・合成禁止・フォールバック警告・コンテキスト無し・0 件、PDF 側パスの冪等。
- run_tests 全緑 + 既存 P1 ケース不変。

**W0 実施記録(2026-09-09・完了)**

設計時に計画外の項目 2 件を追加した(いずれも設計ゲートで提示・承認済み):

| 追加 | 理由 |
|---|---|
| **モジュールルート写像**(`<rel>` 空) | `Get-ModuleDataFiles -Directory $PSScriptRoot` が渡すパスは `<repo>\modules\<tier>\<module>` で `<rel>` が無く、P1 の `Split('\',3)` ガード(3 要素未満は恒等)に弾かれて列挙が一切 PDF に写らなかった。`<rel>` 空を許可し、ただし**ディレクトリとしてのみ照合**する(`modules\<tier>\<file>` が `<PDF>\modules\<file>` を誤って掴むのを防ぐ)。末尾 `\` 付きも同経路に正規化 |
| **`resolvedFrom` テレメトリの是正**(バグ修正) | P1 の判定が「`Resolve-ModuleDataPath` でパスが変化したか」だったため、`Get-ModuleDataFiles` が返した**PDF 側パスをそのまま `Import-ModuleCsv` に渡すと `module` と誤記録**される。§4.3 の採用元エビデンスが列挙経路で嘘をつくため、判定を「最終パスが PDF 配下か」(内部 `Get-FabriqDataOrigin`)に変更。P1 の全経路では結果は同値 |

併せて §4.3 の未了分だった **HTML チェックリストへの採用元記録**を実装(Meta に `Data Set` = PDF 名 /
`(module defaults)`。env が無い `[cl]` 再生成でも `ProfilePath` から導出)。内部ヘルパ
`Get-FabriqOverlayCandidate`(純写像)/ `Write-FabriqDataResolution`(dedup 表示)/ `Get-FabriqDataOrigin` を
抽出し、`Resolve-ModuleDataPath` と `Get-ModuleDataFiles` が同一規則を共有する形にした(KERNEL_API.md §6 に追認)。

検証: run_tests **455/455 PASS**(P1 の 437 + Phase 2 ブロック 18。既存 P1 ケースは無改変)/ パーサ OK /
`check_ps1_encoding` exit 0 / `check_version` exit 0。`KERNEL_VERSION` は 3.6.2 据置(§I)。

**表示の非対称について(仕様)**: 列挙が PDF を採用した場合は 1 行(後続 `Import-ModuleCsv` は PDF 側パス
= 恒等写像で無音)、フォールバック時は「glob 1 行 + ファイル N 行」になる。Profile-First の原則
(冗長 > 無言)に照らして許容する。

### 12.2 W1: 純読み取り資材フォルダ(10 モジュール・軽量版ゲート)

編集パターンは全件同一: `$dir = Resolve-ModuleDataPath -Path (Join-Path $PSScriptRoot "<folder>")`。
既存の「フォルダ不在 → Error」「ファイル不在 → NOT FOUND 表示」ロジックは**無改修**で PDF 側に対して働く
(= 部分欠落は各モジュールの既存表示で顕在化する。無言経路なし)。

| モジュール | 変数(行) | 備考 |
|---|---|---|
| driver_config / driver_import_config.ps1 | `$driverDir`(:38) | export 側は P3 |
| cert_config / cert_config.ps1 | `$certsDir`(:108) | |
| odt_config / odt_install.ps1 | `$AssetsDir`(:54) | `setup.exe` + 構成 XML を同フォルダから解決 |
| default_app_config / default_app_config.ps1 | `$xmlDir`(:38) | export_app_associations(書き)は P3 |
| app_config / app_config.ps1 | `$fileDir`(:31) | インストーラ実体 |
| copyfile_config / copyfile_config.ps1 | `$sourceDir`(:35) | |
| wallpaper_config / wallpaper_config.ps1 | `$wallpaperDir`(:237) | 相対 FileName 行のみ対象、絶対パス行は無影響 |
| ppkg_config / ppkg_install_config.ps1 | `$fileDir`(:44) | uninstall は CSV のみ(P1 で解決済み) |
| manual_kitting_assistant / manual_kitting_assistant.ps1 | `$promptDir`(:99) | |
| pianist / pianist.ps1 | `$script:profilesRoot`(:147) | Q3 裁定 per-case。Studio の Pianist Editor は本体側を編集するため、PDF 運用時は Studio 側の追随(Q6 起票済み)まで手コピー |

各モジュール VERSION MINOR / REQUIRES_KERNEL 3.7.0 / CHANGELOG 1 行。検証 = パーサ + 「PDF にフォルダあり
/なし/空」の 3 状態を dev 機でロジック確認(資材の実適用は VM)。

### 12.3 W2: 列挙系(reg_hklm / reg_hkcu × config / delete・軽量版ゲート)

- 4 スクリプトの `Get-ChildItem -Path $PSScriptRoot -Filter "reg_*_list*.csv" -File | Sort-Object Name` を
  `Get-ModuleDataFiles -Directory $PSScriptRoot -Filter "reg_*_list*.csv"` に置換(1 行)。
- 以降のループ(`Import-ModuleCsv -Path $csvFile.FullName`)は無改修。
- csv_editor の Registry 系エントリは W4 で PDF 認知。

### 12.4 W3: パイプライン族(読み書きが同一フォルダで閉じる例外群・軽量版ゲート)

「解決したフォルダを読みにも書きにも使う」ことで、PDF 内で入出力が閉じる。§4.6 の書き込み原則
(P1〜P2 は本体側)に対する**承認済み例外**として本節に列挙する(理由: 入力だけ PDF・出力だけ本体に
分かれると同一モジュール族の前後工程が食い違う)。

| 族 | 変更 | 閉じ方 |
|---|---|---|
| startlayout_config(backup / build / import) | `$jsonDir` `$xmlDir` `$ppkgDir` を解決 | backup が json を PDF に書き、build が json を読み xml/ppkg を PDF に書き、import が ppkg を PDF から読む |
| sysprep_config + taskbar_config | sysprep `$sourceDir`(:26)を解決。taskbar の `$sysprepSourceDir`(`..\sysprep_config\source`)も解決 | taskbar が PDF の sysprep source に書き、sysprep がそこからステージング |
| printer_driver_config / printer_driver_install.ps1 | `$INF_DIR`(:10)を解決 | アーカイブ展開先(`Join-Path $INF_DIR <BaseName>`)も PDF 内 |

- 解決先フォルダが PDF に無い場合は本体側(フォールバック警告)。**PDF 側にフォルダだけ作れば**
  以後の入出力は PDF に閉じる、という運用を Guide に明記する。
  → W3 実施時に README「プロファイル別データオーバーレイ」節(オペレータ向け)へ記載。

**W1〜W3 実施記録(2026-09-09・完了)**

| wave | 対象 | 検証 |
|---|---|---|
| W1 | 資材フォルダ 10 モジュール 13 箇所。計画書比 +3 箇所(odt の行別相対 `AssetsFolder` :83/:216/:271)、taskbar の行番号は :41 → :71 に移動していた | 実モジュール dir に対する 4 状態 × 10 = 40 チェック PASS(PDF 無し → 本体 / PDF 空 → PDF / PDF 実体 → PDF / コンテキスト無し → 恒等) |
| W2 | reg 列挙 4 スクリプト(1 行置換) | 4 状態 × 2 モジュール: 件数・ソート順・`resolvedFrom` 分類・`Import-ModuleCsv` 恒等・行読み出し PASS |
| W3 | startlayout(json/xml/ppkg)/ sysprep source + taskbar クロス書き / printer INF | クロスモジュール両側が同一 PDF パスに解決すること、展開先と §8 封じ込めガードが解決後ルートで整合すること、コンテキスト無しの恒等を実測 PASS |

各 wave で run_tests 455/455・パーサ・`check_ps1_encoding` exit 0。モジュールは全件 MINOR /
`REQUIRES_KERNEL` 3.7.0。

### 12.5 W4: ツーリング追随(apps)— 2026-09-09 裁定で縮小

当初案は csv_editor に「データセット選択」UI を、fabriq_ios にセッション単位のデータセット選択を
追加するものだった。着手前調査と**ユーザー裁定(2026-09-09)**により次のとおり縮小:

- **csv_editor: 対象外(恒久)**。ユーザー裁定「CSV エディターは正直もういらない」。実運用の CSV 編集は
  FabriqStudio が担っている。着手前調査でも `$script:CsvRegistry` の静的 20 エントリのうち **6 件が
  実在しないパス**(`modules\standard\reg_config\*` 4 件 = 現行は reg_hklm_config / reg_hkcu_config、
  `gyotaku_template\task_list.csv`、`autokey_template\recipe.csv`)を指しており、オーバーレイ以前に
  レジストリ自体が陳腐化していた。**撤去候補として TM に起票**(オーバーレイ対応はしない)。
- **fabriq_ios: 追加実装なし。「今まで通り使える」ことを回帰検証で担保**(ユーザー要件)。
  fabriq_ios は `kernel/common.ps1` を**丸ごと** dot-source するため新公開 API は自動的に可視であり、
  データコンテキスト(`FABRIQ_PROFILE_DATA_DIR`)を一切設定しないので全解決が恒等写像になる
  = P2 以前と完全に同一挙動。検証: 新 2 関数の可視性 / コンテキスト未設定 / 資材パス恒等 /
  `Get-ModuleDataFiles` の返却が旧 `Get-ChildItem ... | Sort-Object Name` と完全一致、を実測。
  加えて IOS 自身のテスト(`module_schema` / `do` / `_phase*_smoke` ≒ 339 アサーション)が
  run_tests 455 の中で全緑。
  - PDF 認知(データセット選択)が必要になった場合の設計は上記当初案のまま保留。IOS は編集系のため
    実行時の誤設定リスクが無く、優先度は低い。

### 12.6 P2 敵対検証(事前)

| # | 攻撃 | 封じ方 |
|---|---|---|
| 1 | PDF に空フォルダだけ置いた → 資材ゼロで実行される | 仕様(§4.5)。各モジュールの既存 NOT FOUND / Error 表示で顕在化し、無言ではない |
| 2 | PDF フォルダに一部ファイルだけ → 残りが本体から補われると誤解 | フォルダ単位 all-or-nothing。ファイル単位の補完はしない(§4.5)。Guide に明記 |
| 3 | 列挙系で PDF に 1 ファイル、本体に 3 ファイル → 合成されて 4 ファイル分適用 | `Get-ModuleDataFiles` は PDF 側一致が 1 件以上なら PDF のみ(§4.4)。Pester でピン留め |
| 4 | printer INF の展開・startlayout の生成物が PDF を「汚す」 | 意図した閉じ込め(§12.4)。PDF は案件データの一部 |
| 5 | Studio(Pianist Editor / レジストリ辞書)が本体側を編集し続け PDF と乖離 | Q6 裁定(2026-09-09)で Studio 側にデータセット選択タスクを起票。それまでは既知の運用制約(PDF の CSV は本体と同じ構造なので Excel 等で直接編集可)。csv_editor は撤去方針のため追随させない(§12.5) |
| 6 | ディレクトリ対応で P1 の挙動が変わる | 変わるのは「PDF にフォルダが存在する」時だけ。コンテキスト無し・フォルダ無しは恒等写像のまま。Pester の P1 ケースを不変のまま維持 |
| 7 | 大文字小文字違いのフォルダ名 | Windows FS は大文字小文字不問。`Test-Path` で判定するため問題なし |

### 12.7 P2 変更スコープ宣言(雛形・着手時にそのまま使う)

```
【変更スコープ宣言】(W0)
- 対象: kernel
- 公開 API サーフェスへの影響: あり(Resolve-ModuleDataPath のディレクトリ対応 = 後方互換拡張 /
  Get-ModuleDataFiles 新設)
- 予想バージョン影響: kernel MINOR(3.7.0 節に追記) / modules 変更なし
- 影響テスト: tests/kernel/ProfileDataOverlay.tests.ps1(追加)
- 実行予定: run_tests.ps1(必須)

【変更スコープ宣言】(W1〜W3 各 wave)
- 対象: module:<wave の対象>
- 公開 API サーフェスへの影響: なし
- 予想バージョン影響: 各 MINOR / REQUIRES_KERNEL 3.7.0(新 API 依存)
- 影響テスト: なし(パーサ + ロジック確認、VM 任意)
```

---

## 13. Phase 3 設計(書き込み系・クロスモジュール)

### 13.1 契約(§4.6 の実装形)

- 公開 API 拡張: `Resolve-ModuleDataPath -Path <string> -ForWrite`
  - コンテキスト有効時は**存在に関係なく PDF 側パス**を返し、親ディレクトリを作成する
    (`<PDF>\modules\<module>\backup\...`)。表示は `[DATA] <label> -> profile (X) [write]`。
  - コンテキスト無しは恒等写像。フォールバック概念は書き込みに存在しない(書く先は常に一意)。
- restore 側は通常の読み解決(PDF 優先・欠落時は本体 backup へ Warning 付きフォールバック)。
  backup(書き)→ restore(読み)の対称性が PDF 内で閉じる。

### 13.2 対象と編集点

| モジュール | 書き側 | 読み側 |
|---|---|---|
| acl_config | acl_backup.ps1 `$backupBaseDir`(:80) `-ForWrite` | acl_restore.ps1 `$backupBaseDir`(:76) |
| reg_template | reg_backup.ps1 `$backupDir`(:23) `-ForWrite` | reg_import.ps1 `$backupDir`(:25) |
| firewall_rule_config | firewall_rule_export.ps1 `$defaultBackupRoot`(:186) `-ForWrite` + Add-ImportEntry の CSV 書き戻し(P1 で解決済みパス) | firewall_rule_import.ps1 `$resolved`(:80)の backup 基点 |
| desktop_icon_config | desktop_icon_backup.ps1 `$backupDir`(:47) `-ForWrite` | desktop_icon_restore.ps1 `$backupDir`(:19) |
| driver_config | driver_export_config.ps1 `$driverDir`(:38) `-ForWrite` | (import は W1 済み) |
| default_app_config | export_app_associations.ps1 `$xmlDir`(:38) `-ForWrite` | (import は W1 済み) |

### 13.3 敵対検証(事前)

| # | 攻撃 | 封じ方 |
|---|---|---|
| 1 | backup は PDF に書いたが restore がコンテキスト無しで走り本体 backup を読む | メニュー単発は Q1 (a) で常に本体側(恒久確定)。README に「backup と restore は同じプロファイルの中で対にする」を明記。なお**本体側に古い backup が残っていない限りは fail-closed**(見つからず Error)であり、誤ったデータの復元が起きるのは本体側に旧案件の捕捉物が残っている場合に限られる |
| 2 | `-ForWrite` が PDF 外に書く | 写像規則は読みと同一。モジュール外パスは恒等写像(本体側に書く) |
| 3 | 親ディレクトリ作成の失敗 | 作成失敗は例外 → 呼出側の既存 try/catch で Error |
| 4 | 削除系(restore の cleanup 等)が PDF 側を再帰削除 | CLAUDE.md §8 ガード(`Test-FabriqSafe*`)を PDF パスにも適用する。既存ガードは containment 検証なので PDF ルートを許容ベースに追加 |

### 13.4 Phase 3 実施記録(2026-09-09・完了)

**§13.1 からの変更 2 点**(設計ゲートで提示・承認済み):

| 変更 | 理由 |
|---|---|
| **`-ForWrite` はディレクトリを作成しない** | 当初案は「親ディレクトリを作成する」だったが、実コードを読むと**書き側 6 本は全て自前で生成している**(`acl_backup.ps1:212` が `-Force` で連鎖生成 / `reg_backup.ps1:26` / `firewall_rule_export.ps1:252` / `desktop_icon_backup.ps1:50` / `driver_export_config.ps1:41` / `export_app_associations.ps1:38`)。解決しただけで `<PDF>\...\backup\` が生えると、**プレビューやキャンセルで終わった実行が空フォルダを残し、後続の読み解決が「PDF にフォルダあり」に化ける**遠隔作用が入る。リゾルバは純関数のままにした |
| **§13.3 #4(§8 ガードに PDF ルートを許容追加)は不要だった** | 6 モジュールの再帰削除は 2 箇所(`acl_backup.ps1:210` / `driver_export_config.ps1:179`)だけで、いずれも**解決後のベースに対する inline containment 判定**のため解決先に自動追随する。`Test-FabriqProtectedPath` はこの 6 本のどこからも呼ばれていない。**カーネルのガード変更はゼロ** |

**restore を fail-closed にしなかった判断**: 「コンテキスト有効だが PDF に backup を作ったことがない
→ 本体側の別案件の捕捉物を復元しうる」経路は残る。これは Show-Warning で可視化するに留めた。
理由は Q5(strict mode 不採用)と同じ論理で、可視化 3 点で足り、ハード停止は移行途中の全停止を招くため。
**オーバーレイ以前と同じ挙動 + 警告**であり退行ではない。当初はメニュー単発のデータセット選択(§14.2)で
構造的に解消する想定だったが、それも 2026-09-09 に不採用が確定したため、**本項は恒久的な運用上の制約**として
README に明記する形で決着した。

**列挙を `Get-ModuleDataFiles` にしなかった判断**: restore 側の列挙(`DesktopIcons_*.reg` /
`*_<name>.reg`)は**フォルダ解決 + 既存 `Get-ChildItem`** のままにした。PDF の backup フォルダが
「有るが空」は「このデータセットにバックアップ無し」の明示(= Error)であるべきで、
他データセットのファイルへ落ちてはいけないため。

**編集点(実測)**

| モジュール | 書き側(`-ForWrite`) | 読み側(通常解決) | VERSION |
|---|---|---|---|
| acl_config | acl_backup.ps1 `$backupBaseDir`:80 | acl_restore.ps1 `$backupBaseDir`:76 | 1.0.2 → 1.1.0 |
| reg_template | reg_backup.ps1 `$backupDir`:23 | reg_import.ps1 `$backupDir`:25 | 1.1.0 → 1.2.0 |
| firewall_rule_config | firewall_rule_export.ps1 `$defaultBackupRoot`:186 | firewall_rule_import.ps1 の backup アンカー:80 | 1.1.0 → 1.2.0 |
| desktop_icon_config | desktop_icon_backup.ps1 `$backupDir`:47 | desktop_icon_restore.ps1 `$backupDir`:19 | 1.0.1 → 1.1.0 |
| driver_config | driver_export_config.ps1 `$driverDir`:38 | (import は W1 済み) | 1.2.0 → 1.3.0 |
| default_app_config | export_app_associations.ps1 `$xmlDir`:38 | (適用側は W1 済み) | 1.1.0 → 1.2.0 |

全件 MINOR / `REQUIRES_KERNEL` 3.7.0。firewall は `Add-ImportEntry` が **`BackupRoot` 相対**で
`SourcePath` を書き、import 側が同じ解決後ルートを相対の起点にするため、export の書き先と
import のアンカーが一致する(実測確認済み)。

**検証**

- run_tests **464/464 PASS**(P2 の 455 + Phase 3 ブロック 9)。
- dev 機ロジックチェック **37 件全 PASS**: 6 対 × (コンテキスト無しで書き読みとも恒等 / 空 PDF で
  write→PDF・read→本体・**何も生成されない** / モジュールが実際に書いた後は対が PDF 内で閉じる)
  + firewall の export root == import anchor。
- VM リグ(clean-base): `reg_template` と `firewall_rule_config` に overlay シナリオを追加し
  **4/4 PASS / exit 0**。

  | Module | Scenario | Verdict | Oracle |
  |---|---|---|---|
  | reg_template | backup(コンテキスト無し) | PASS(SELF) | 後方互換 |
  | reg_template | overlay-backup | **PASS** | `True/True` = .reg が PDF にあり本体 backup は空 |
  | firewall_rule_config | export(コンテキスト無し) | PASS | 後方互換 |
  | firewall_rule_config | overlay-export | **PASS** | `True/False` = policy.wfw が PDF にあり本体側に無い |

- 実 VM の表示目視:

  ```
  [INFO]    [DATA] reg_template/backup/ -> profile (_p3_display) [write]
  [WARNING] [DATA] reg_template/backup/ <- module dir (FALLBACK: not in profile data folder)
  created by resolving? : False          ← 解決では何も生成されない
  read source (after write) : ...\profiles\_p3_display\modules\reg_template\backup
  pair closed in the PDF : True
  no context -> write / read とも本体側(恒等)
  ```

  書きと読みが**同一バッチ内で両方表示**される(dedup キーが独立)ことを実機で確認。
---

## 14. Phase 4(締め・運用移行)— **全項目が裁定により終了(2026-09-09)**

Q1〜Q6 の全裁定を経て、Phase 4 に**実装すべき項目は残っていない**。4 項目の決着は次のとおり:

| 項目 | 決着 |
|---|---|
| 14.1 strict mode | **不採用**(Q5) |
| 14.2 メニュー単発のデータセット選択 | **不採用**(Q1 恒久確定 — メニュー単発はモジュールデフォルト CSV) |
| 14.3 hostlist / Studio | hostlist = **見送り確定**(Q2)/ Studio = **役割分担 + 別リポジトリへ起票済み**(Q6) |
| 14.4 本体 CSV のサンプル縮退 | **不実施**(積極的な削除はせず自然消滅に委ねる) |

→ **本計画書の実装スコープは Phase 1〜3 で完了**。

### 14.1 strict mode — **不採用**(Q5 裁定 2026-09-09)

当初案は PDF 直下のマーカー `overlay_policy.txt`(`warn` | `strict`)でフォールバックを Error 化し、
移行完了プロファイルから順に締めていくものだった。**機能自体を入れないと裁定した。**

- 判断理由: フォールバックは既に 3 経路で可視化されている — 実行画面の `Show-Warning`(1 バッチ
  1 ラベル 1 回)/ csv.load テレメトリの `resolvedFrom` / HTML チェックリストの `Data Set`。
  R1(誤った設定セットの無言適用)に対する対策としてはこれで足り、ハード停止を足すのは
  「移行途中のプロファイルが 1 ファイルの置き忘れで全停止する」副作用のほうが大きい。
- 帰結: §1 の原則「最終形では strict に締められること」は撤回し、**可視化 3 点を最終形**とした。
  `Resolve-ModuleDataPath` / `Get-ModuleDataFiles` に strict 分岐は入れない(現行実装が最終形)。
- PDF ルート直下の予約名前空間(§3)は維持する(案件メタデータ用)。ポリシーマーカーは置かない。

### 14.2 メニュー単発実行のデータコンテキスト選択 — **不採用**(Q1 恒久確定 2026-09-09)

当初案はメインメニューに「データセット: なし / <profile 名>」の切替を追加し、単発実行の前後で
`Set-FabriqProfileDataContext` を set/clear するものだった。**入れないと裁定した。**

- 確定した仕様: **メニュー単発実行は常にモジュールデフォルトの CSV を使う**(コンテキスト無し =
  恒等写像)。プロファイル実行だけがデータセットを切り替える、という一本の線を維持する。
- 帰結: Q1 の「Phase 4 で再判断」は再判断を実施したうえで (a) のまま確定。以後この論点は開かない。
- 副次: §13.3 #1(メニュー単発の restore が本体側 backup を読む)は**恒久的な運用上の制約**となる。
  README に「backup と restore は同じプロファイルの中で対にする」を明記済み。実害は本体側に
  旧案件の捕捉物が残っている場合に限られ、残っていなければ「見つからず Error」の fail-closed。

### 14.3 hostlist(Q2)と Studio(Q6)の裁定内容

**Q2 = (a) 見送りで確定**(保留ではなく結論)。本体 `kernel/csv/hostlist.csv` を正のまま維持する。

- 衛星 2 本は `kernel\csv\hostlist.csv` の**存在によって fabriq ルートを発見**する構造
  (`fabriq_checksheet/checksheet/common.ps1:1062` / `fabriq_backuper/backuper/common.ps1:753`)。
  さらに backuper は `extended_hostlist` と本家 hostlist の (OldPCname, NewPCname) **集合完全一致
  ゲート**を持つ。衛星は単独起動(別 PC のこともある)で「実行中プロファイル」を知り得ない。
- lifecycle が違う: hostlist は**案件ごとに作り直すジョブ入力**、PDF は**顧客ごとに安定するレシピ**。
- 将来 real pain が出た場合の逃げ道としてのみ、片方向同期案(本体を正としたまま PDF 側 hostlist を
  マージ元にする)を記録として残す。**現時点では実装しない。**

**Q6 = (a) 役割分担 + Studio 側タスクを起票**。

- 役割: ワークスペース = **物理分離**が要るとき(別現場・別持出し PC)、PDF = **同一配備内**の
  設定セット切替。統合はしない。
- 起票理由: csv_editor を撤去対象にした(§12.5)結果、**PDF 側 CSV を編集するツールが現状ゼロ**。
  Studio は本体側を編集し続けるため、PDF 運用を始めると Studio の編集先と fabriq の読み先が乖離する
  (§12.6 敵対検証 #5 が現実になる)。
- Studio 側の実装形(調査済み): Studio は `IWorkspaceService.RootPath` + **ルートからの相対パス**で
  CSV を読み書きする(`ICsvService` / `Services/Master/MasterTargetResolver.cs:26`)ため、差し込み口は
  ほぼ 1 箇所。ただし計画書が当初書いていた「ルートを 1 つ挿し替えるだけ」は**不正確**で、
  PDF レイアウトは `<tier>` を落とす(`modules/<module>/<rel>`)ため、**§4.2 の写像関数を Studio 側に
  ミラーする**必要がある(15 行程度)。tier を落とす設計自体は維持する — 操作者が standard/extended を
  意識せずに PDF を作れること、モジュールの tier 移動が PDF を壊さないこと、が理由。
- 起票先: `E:\fabriq_studio/.tm/tasks.json`(Studio 独自 TM)。
- **ツール実装者向けの要約は `dev/PROFILE_DATA_OVERLAY_FOR_TOOLING.md`** に切り出した
  (写像規則の擬似コード・粒度契約・PDF 対象/フレームワーク資産の判断表・やらないこと一覧)。
  Studio 側 CLAUDE.md が「必ず `E:\fabriq` を読んで準拠する」と定めているため、契約の置き場は
  fabriq 側に集約する。

### 14.4 本体 CSV のサンプル縮退 — **不実施**(裁定 2026-09-09)

当初案は全 wave 完了後に本体側 CSV を dev/template 同等の「サンプル」(Enabled=0 の例示行のみ)へ
縮退させるものだった。**積極的な削除はしない。**

- 確定した方針: 本体側 CSV は**使わないが消しもしない**。PDF 運用が定着するにつれて参照されなくなり、
  **自然消滅していく**ことを想定する。
- 理由: 一斉縮退は「移行が済んでいないプロファイル」を一度に壊しうる作業で、得られるのは見た目の
  整理だけ。フォールバックは可視化済み(Q5 の判断と同型)であり、放置のコストは低い。
- 副次: §13.3 #1 の実害条件「本体側に旧案件の捕捉物が残っている」は、この方針では時間とともに
  自然に減っていく(新規の捕捉は PDF 側に書かれるため本体側は増えない)。

---

## 15. Phase 1 VM E2E 手順(リグ準備後に実施)

### 15.1 前提と制約

- リグ(`dev/test_rig`)はモジュール単位の headless 実行で、`Invoke-BatchExecution` と `__RESTART__` は
  対象外(run_scenario.ps1 ヘッダ参照)。よって E2E は **(a) リグ headless(解決の実動作)** と
  **(b) VM 上の実 Fabriq 手動実行(コンテキスト寿命と resume)** の 2 本立て。

### 15.2 (a) リグ headless — envelope 拡張(dev ツーリング、kernel 非接触)

- `run_module_tests.ps1` / `run_scenario.ps1` の envelope に `profileDataDir` キーを追加し、
  VM 側で `$env:FABRIQ_PROFILE_DATA_DIR` に設定する(既存の `segment` と同じ配管、
  run_module_tests.ps1:86-87 と :112 の隣)。
- シナリオ: `test_harness_config` の apply に `profileDataDir='C:\fabriq\profiles\_test_harness'` を与え、
  VM 上に `profiles/_test_harness/modules/test_harness_config/test_harness_list.csv`
  (本体と異なる Description を持つ行)を置く。
  期待: 出力に `[DATA] test_harness_config/test_harness_list.csv <- profile (_test_harness)` と
  PDF 行の Description。PDF CSV を消す(フォルダは残す)→ `[WARNING] ... (FALLBACK ...)` + 本体行。
  `profileDataDir` を空にする → `[DATA]` 行なし + 本体行(P1 以前と同一出力)。
- 追加: `windows_license_config`(license_key.csv を PDF に置いた場合の門番通過)は
  実キー投入を伴うため **VM スナップショット前提**で任意。

### 15.3 (b) 実 Fabriq 手動実行(VM コンソール) — 操作者向け手順

**専用フィクスチャ(リポジトリ同梱・VM へ同期済み 2026-09-09)**:
- `profiles/_test_overlay.csv`: `__AUTOPILOT__`(WaitSec=1)→ test_harness ×2 → `__RESTART__`(Order 40)→
  test_harness ×2。全行 Segment 空。
- `profiles/_test_overlay/modules/test_harness_config/test_harness_list.csv`: 1 行、Description が
  `OVERLAY ROW - read from profiles/_test_overlay (profile data folder)`。本体 CSV の既定行
  (`Default scenario - single success with Verified PASS`)と表示で区別できる。

**前提**: VM の `C:\fabriq` に kernel / apps / commands / test_harness_config / 上記フィクスチャが同期済み
(dev 機から `vm_sync_full.ps1` 相当で実施)。`kernel\json\resume_state.json` が無いこと。
VM のコンソール(対話セッション)で操作する。各シナリオ前にスナップショットは不要(フィクスチャは
非破壊)。所要時間の目安: 合計 20〜30 分。

| # | 操作 | 期待される画面 | 判定 |
|---|---|---|---|
| S1 | `Fabriq.exe` 起動 → パスフレーズ/作業者/ホスト選択 → プロファイル実行(Linear)で `_test_overlay` を選ぶ | バッチ開始直後: `[INFO] [DATA] Profile data folder: C:\fabriq\profiles\_test_overlay`。Order 20 の実行中に `[INFO] [DATA] test_harness_config/test_harness_list.csv <- profile (_test_overlay)` が出て、モジュール表示の Description が `OVERLAY ROW - ...`。Order 30 では `[DATA]` 行が**再表示されない**(1 バッチ 1 回) | 両方見えたら PASS |
| S2 | Order 40 の `__RESTART__` で 5 秒カウントダウン後に再起動 → ログオン後 Fabriq が自動起動し「Profile Resume Detected」→ 60 秒カウントダウン(**Enter で即再開可、Esc は押さない**) | 再開後のバッチで再び `[DATA] Profile data folder: ...` と `<- profile (_test_overlay)`、Order 50/60 も `OVERLAY ROW - ...`。完走後に HTML チェックリスト生成 | PASS |
| S3 | (fail-closed)S1 をもう一度実行。再起動後の**60 秒カウントダウン中**に Explorer で `C:\fabriq\profiles\_test_overlay` フォルダを `_test_overlay_bak` にリネームし、カウントダウンを満了(または Enter) | `[ERROR] Profile data folder recorded before the restart is missing: C:\fabriq\profiles\_test_overlay` → `[ERROR] Resume aborted ...` → `[WARNING] Remaining modules were NOT executed. Restore the profile data folder and relaunch Fabriq to resume.` Order 50/60 は実行されず、`kernel\json\resume_state.json` が**残っている** | PASS |
| S3' | フォルダ名を `_test_overlay` に戻して Fabriq を再起動 | 「Profile Resume Detected」→ Order 50/60 が `OVERLAY ROW - ...` で実行され完走 | PASS |
| S4 | (後方互換)`_test_overlay` フォルダを `_test_overlay_off` にリネームしたまま S1 を実行 | バッチ開始直後に `[INFO] [DATA] No profile data folder for this profile (module CSVs in use)` の 1 行のみ。以降 `[DATA]` 行は一切出ず、Description は `Default scenario - single success with Verified PASS`。`__RESTART__` 後の resume も従来通り | PASS(終了後フォルダ名を戻す) |
| S5 | (Flex)operator(FlexProfile ダッシュボード)で `_test_overlay` を開き、Order 20 を **Run This**、続けて 20/30 を **Run Selected** | 各バッチの開始時に `[DATA] Profile data folder: ...` バナー、`<- profile (_test_overlay)` はバッチごとに 1 回。ダッシュボードを閉じてメインメニューからスクリプトメニュー [S] で test_harness_config を単発実行 → `[DATA]` 行が**出ない**(コンテキスト漏れなし) | PASS |

**記録**: 各行の PASS/FAIL と気づいた点(表示の読みづらさ等)を t-0097 の claudeNote 用にメモ。
FAIL があれば画面のメッセージをそのまま控える(ログは `logs\` のトランスクリプトにも残る)。

### 15.4 記録

- 結果は t-0097 の claudeNote と本節に追記。全項目 PASS で P1 完了条件を満たす。

**実施記録 (a) リグ headless — 2026-09-09 実施・全 PASS**(VM 10.1.10.8 / PS 5.1.26100 / winrm Session 0):

- リグ envelope に `profileDataDir` を追加(`run_module_tests.ps1` の inproc 経路 + session-b request、
  `_rig_interactive_runner.ps1`)。`test_harness_config/test.psd1` に `overlay-profile` / `overlay-fallback`
  シナリオを追加(fixture が PDF CSV を生成、teardown で PDF フォルダを丸ごと除去)。
- `run_module_tests.ps1 -SyncRepo -Module test_harness_config -Idempotency`:

  | シナリオ | 結果 | 意味 |
  |---|---|---|
  | simulate(コンテキスト無し) | PASS(SELF) Success / Verified=True | P1 以前と同一(後方互換) |
  | overlay-profile | PASS Skipped / Idem OK / teardown undone | PDF 側 CSV(Behavior=skip)が本体 CSV に勝った |
  | overlay-fallback | PASS Success / Verified=True / Idem OK | PDF にフォルダのみ → 本体 CSV へフォールバック |

- アドホック実行(コンソール中継)で表示契約を目視確認:
  - profile-copy: `[INFO] [DATA] test_harness_config/test_harness_list.csv <- profile (_test_harness)` → Skipped
  - fallback: `[WARNING] [DATA] test_harness_config/test_harness_list.csv <- module dir (FALLBACK: not in profile data folder)` → Success
  - no-context: `[DATA]` 行なし・出力は従来と同一 → Success
  - 3 ケースとも実行後に `FABRIQ_PROFILE_DATA_DIR` が空(コンテキスト漏れなし)

**(b) 実 Fabriq 手動実行 — 2026-09-09 ユーザー実施・ログ照合済み・S1〜S5 PASS**(VM コンソール、
Fabriq.exe、DefaultAsync ON = モジュールは監視 Runspace で実行):

| # | ログ上の証跡(logs/ トランスクリプト + logs/telemetry csv.load `resolvedFrom` + history export) | 判定 |
|---|---|---|
| S1 | 10:20 セッション: `[DATA] Profile data folder: C:\fabriq\profiles\_test_overlay` バナー → Order 20/30 の csv.load が `resolvedFrom=profile` → `__RESTART__`(ResumeAfter 40) | PASS |
| S2 | 10:51 セッション: `Profile Resume Detected` → バナー再表示 → Order 50/60 も `resolvedFrom=profile` → history export に 20/30/[RESTART]/50/60 全 Success/Verified=True、HTML チェックリスト生成 | PASS |
| S3 | 10:53 セッション: `Profile Resume Detected` → `[ERROR] Profile data folder recorded before the restart is missing` → `[ERROR] Resume aborted ...` → `[WARNING] Remaining modules were NOT executed ...`。モジュール telemetry ゼロ(= 未実行)、resume_state 温存 | PASS |
| S3' | 10:55 セッション: フォルダ復旧後 `Profile Resume Detected` → バナー → Order 50/60 `resolvedFrom=profile` → 完走(history export に S3 一段目〜S3' 二段目まで累積、[RESTART] 2 件は同一セッション ID 復元による正しい累積) | PASS |
| S4 | 10:56/10:57 セッション: `[DATA] No profile data folder for this profile (module CSVs in use)` の 1 行のみ、全 4 実行が `resolvedFrom=none`、resume も従来通り完走 | PASS(後方互換) |
| S5 | 10:58 セッション(Flex): Run This + Run Selected の各バッチ開始に `[DATA] Profile data folder:` バナー、3 実行とも `resolvedFrom=profile`、execution_history に 3 行。メニュー単発実行の漏れ確認は未実施(Pester でカバー) | PASS |

副次的な確認事項:
- **監視 Runspace(async)内でも解決が効く**ことを実 OS で確認(env はプロセス共有 — 敵対検証 #9 の実証)。
- Runspace 実行のためモジュール内の `<- profile` 行はトランスクリプトに載らない(画面と telemetry のみ)。
  §4.3 の「エビデンスへの採用元記録」は csv.load テレメトリの `resolvedFrom` で満たしており、
  HTML チェックリストへの表示は P2 の改善項目として維持。
- トランスクリプト中の `_***_overlay` は秘密マスクがマスターパスフレーズ(テスト用)と一致する
  部分文字列を伏せたもの。仕様通りで実害なし(本番の長いパスフレーズでは起きない)。
- **改善(P1 ポリッシュ、2026-09-09 修正済み)**: S3 の abort 経路で残モジュールを実行しなかった後、
  既存の完了バナー `Profile Execution Completed`(緑)がそのまま出ていた。Linear / Flex 両 abort 経路で
  `[RESTART]` の Error 結果 + 実行履歴行を記録するようにし(RunOnce 失敗時と同じ記録パターン)、
  完了バナーは `Completed with Errors` 側、チェックリスト/履歴にも abort が残るようにした。
  検証は parse + run_tests(main.ps1 トップレベルのため単体テスト外)。

---

## 16. Phase 2 VM E2E(2026-09-09 実施・全 PASS)

VM を clean-base に revert した状態から実施(10.1.10.8 / DESKTOP-M8FC97M / PS 5.1.26100)。
P1 と異なり **実 Fabriq の手動操作は不要**と判断した。P2 が足したのは「解決の種類」(ファイル →
フォルダ／列挙)だけで、コンテキストのライフサイクル・`__RESTART__` 跨ぎ・resume fail-closed といった
**プロセス側の面は P1 の S1〜S5 で検証済み**かつ P2 で未変更のため。P2 の新面はリグ headless で
モジュール実挙動まで到達できる。

### 16.1 リグ記述子(新規 3 シナリオ)

Phase 2 の 3 つの新面それぞれに、**PDF が勝ったことを本体側の痕跡の不在で証明する** oracle を置いた
(「PDF 側が使われた」だけでなく「本体側が使われていない」ことを同時に見る)。

| wave | モジュール / シナリオ | 仕掛け | oracle |
|---|---|---|---|
| W1 | `copyfile_config` / `overlay-source` | PDF 側 `source\test.txt` に**本体と同名・別内容**(`PDF-SOURCE-W1`)を置く | コピー先の**内容**が `PDF-SOURCE-W1` = 本体の同名ファイルは使われていない(フォルダ単位 all-or-nothing) |
| W2 | `reg_hklm_config` / `overlay-enumeration` | PDF 側に 1 本だけ `reg_hklm_list_overlay.csv`(固有マーカー) | マーカー適用済 **かつ** 本体 CSV 固有の `DisableCAD` が**不在** = 合成されていない |
| W3 | `taskbar_config` / `overlay-sysprep-source` | PDF 側に空の `modules\sysprep_config\source\` だけ作る | `LayoutModification.xml` が **PDF 側に存在し本体側に不在** = クロスモジュール書き込みが PDF に閉じた |

### 16.2 実行結果(`run_module_tests.ps1 -SyncRepo -Idempotency`)

| Module | Scenario | Verdict | Status | Idem | Oracle |
|---|---|---|---|---|---|
| copyfile_config | apply | PASS | Success | OK | present 3/3(**後方互換**: コンテキスト無しで従来どおり) |
| copyfile_config | overlay-source | **PASS** | Success | OK | `PDF-SOURCE-W1`(W1) |
| reg_hklm_config | apply | PASS | Success | OK | registry 6/6(**後方互換**) |
| reg_hklm_config | overlay-enumeration | **PASS** | Success | OK | `True/True`(W2 = マーカー有 / 本体行 無) |
| taskbar_config | apply | PASS | Success | OK | present 1/1(**後方互換**) |
| taskbar_config | overlay-sysprep-source | **PASS** | Success | OK | `True/False`(W3 = PDF 有 / 本体 無) |
| test_harness_config | simulate | PASS(SELF) | Success | – | P1 回帰 |
| test_harness_config | overlay-profile | PASS | Skipped | OK | P1 回帰 |
| test_harness_config | overlay-fallback | PASS | Success | OK | P1 回帰 |

`Summary: 9 scenario(s) | FAIL/ERROR: 0 | manual-revert: 0`、exit 0。teardown は全件 undone
(PDF テストフォルダ・レジストリマーカー・生成物とも自動撤去)。

### 16.3 表示契約の実 VM 目視(アドホック中継)

```
=== case 1: copyfile_config (PDF に copy_list.csv と source/ の両方) ===
[INFO] [DATA] copyfile_config/copy_list.csv <- profile (_p2_display)
[INFO] [DATA] copyfile_config/source/ <- profile (_p2_display)      ← フォルダは末尾 / 付き
=== case 2: reg_hklm_config (PDF に reg_hklm_list*.csv が 1 本) ===
[INFO] [DATA] reg_hklm_config/reg_hklm_list*.csv <- profile (_p2_display)   ← 1 行のみ
=== case 3: 空の PDF -> フォールバック ===
[WARNING] [DATA] reg_hklm_config/reg_hklm_list*.csv <- module dir (FALLBACK: ...)
[WARNING] [DATA] reg_hklm_config/reg_hklm_list.csv <- module dir (FALLBACK: ...)
[WARNING] [DATA] copyfile_config/copy_list.csv <- module dir (FALLBACK: ...)
=== case 4: コンテキスト無し ===
[DATA] line count: 0 / context cleared after run: True
```

- フォルダのラベルが末尾 `/`、列挙のラベルが `<module>/<Filter>` であることを実機で確認。
- **PDF 採用時の列挙は 1 行のみ**(後続の `Import-ModuleCsv` は PDF 側パス = 恒等写像で無音)、
  **フォールバック時は glob 1 行 + ファイル N 行**。§12.1 に記した非対称が実機でもそのとおり出る。
- case 3 で `copyfile_config/source/` のフォールバック行が出ないのは早期 return のため
  (本体 CSV の行は Segment 付きで、`FABRIQ_SEGMENT` 空では 0 行 → `$sourceDir` に到達する前に Skipped)。
  フォルダのフォールバック経路自体は dev 機の 40 チェックと Pester で被覆済み。
- コンテキスト無しでは `[DATA]` 行ゼロ、実行後に env がクリアされていることも確認(コンテキスト漏れなし)。

### 16.4 未被覆(記録)

- W1 の残り 9 モジュール(driver / cert / odt / default_app / app / wallpaper / ppkg /
  manual_kitting_assistant / pianist)は**資材の実体**(ドライバ INF・証明書・インストーラ・PPKG 等)を
  要するため VM 常設リグには載せない。解決ロジックは全件同一イディオムで、dev 機の実モジュール dir に
  対する 40 チェック(4 状態 × 10)と Pester で被覆している。
- W3 の startlayout(ADK / `Export-StartLayout` が Win11 26200 で破綻 = TM t-0083)と
  printer_driver INF(実ドライバ要)は同様に未被覆。sysprep source への書き込みは taskbar 経由で被覆済み。
