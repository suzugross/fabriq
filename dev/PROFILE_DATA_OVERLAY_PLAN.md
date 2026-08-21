# プロファイル別データオーバーレイ 実装計画書

Status: **計画確定前ドラフト・保留中**(実装未着手・着手時期は未定 / TM: t-0095【最重要】)
裁定済み: Q3 = per-case(2026-08-09)。残裁定: Q1 / Q2 / Q4 / Q5 / Q6
作成: 2026-08-09 / 最終更新: 2026-08-09

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
  最終形ではプロファイル単位で strict mode(フォールバック = Error)に締められること。
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
  カーネルは `modules/` 配下しか見ない。FabriqStudio 等が案件メタデータ
  (顧客名・作成日・strict ポリシーマーカー等)を PDF ルートに置けるようにするための
  前方互換予約(§9 Q5 のマーカーファイルもここに置く)。

---

## 4. 解決契約(正式仕様)

### 4.1 データコンテキストのライフサイクル

| タイミング | 動作 |
|---|---|
| プロファイル実行開始 | `profiles/<name>/` が存在すれば `FABRIQ_PROFILE_DATA_DIR` に絶対パスを設定。存在しなければ未設定(= 従来動作)で、その旨を 1 行表示 |
| プロファイル実行終了(finalize / Cancel / Error 終了) | 必ずクリア(finally 相当の位置で) |
| `__RESTART__` 跨ぎ | ResumeState に PDF パスを保存し、resume 時に**復元してから**後続モジュールを実行(既存 profileName 保存の隣に追加)。復元失敗は続行せず Error(base 設定の無言適用は絶対に許さない) |
| メニュー単発実行 | コンテキスト無し = オーバーレイ無効(現状動作)。Phase 4 で「データコンテキスト選択」を検討(§9) |
| async(Runspace)実行 | env はプロセス共有のため追加対応不要(現行 FABRIQ_SEGMENT と同じ扱い)。ただしテストで固定する |

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
| `Resolve-ModuleDataPath -Path <string>` | §4.2 の解決 + §4.3 の表示(表示は 1 実行 1 ファイル 1 回に抑制) | §1.2 追記(MINOR) |
| `FABRIQ_PROFILE_DATA_DIR` | データコンテキスト env | §環境変数 追記 |
| `Import-ModuleCsv` 内部フック | 受領パスを `Resolve-ModuleDataPath` に通す(88 ファイル/100 呼出を無改修で吸収) | 契約文に解決順を追記 |
| ResumeState 拡張 | `profileDataDir` フィールド追加(state JSON スキーマ変更 = kernel PATCH 相当だが、公開 API 追加と合わせて **MINOR 1 回**に束ねる) | — |
| main.ps1 | プロファイル実行開始/終了/resume でのコンテキスト設定・クリア | — |

- 予想バージョン: **kernel MINOR(3.6.1 → 次リリースで 3.7.0 に New-UiFont 分と合流)**。
- モジュール側の改修が必要なのは §7.1 の 8 箇所(Phase 1)と §7.2 以降(Phase 2+)のみ。

---

## 6. フェーズ計画

### Phase 0: 契約凍結(1〜2 日)

- 本書 §3〜§4 をレビューし未決事項(§9)を裁定 → 契約凍結。
- 設計ゲート(フル版)提出: コンテキストのステートマシン + 敵対検証。
- **完了条件**: 本書の Status を「契約凍結」に更新。

### Phase 1: CSV 読みオーバーレイ(3〜5 日)

1. kernel: `Resolve-ModuleDataPath` + コンテキスト管理 + Import-ModuleCsv フック + 表示。
2. kernel: ResumeState 拡張(保存・復元・復元失敗 = Error)。
3. modules: Test-Path 前置き 8 箇所(§7.1)を解決 API 経由に修正(各 PATCH)。
4. tests: 新規 Pester(解決順・恒等写像・restart 復元・表示抑制・Segment/ENC 直交)。
5. VM リグ E2E: PDF あり/なし/フォールバック混在プロファイルの 3 本。
- **完了条件**: run_tests 全緑 + VM E2E 3 本 + 「オーバーレイ不在で既存挙動と完全同一」の回帰確認。
- **ロールバック**: コンテキストを設定しなければ全コードパスが恒等写像 = 機能フラグ不要で即時無効化可能。

### Phase 2: 資材フォルダ・列挙系(1〜1.5 週)

- §7.2 の各モジュールを wave 分割(3〜4 モジュール/wave)で `Resolve-ModuleDataPath` 化。
  読み取り系を先行(driver_import, cert, odt, default_app, app_config, copyfile, wallpaper,
  ppkg_install, printer_driver INF, startlayout_import/build, sysprep source, mka prompt)。
- §7.3 列挙系(reg_hklm/hkcu × config/delete)に §4.4 を実装。
- csv_editor / fabriq_ios の PDF 認知(編集対象の切替 UI)は本 Phase の後半。
- 各 wave: 軽量版ゲート + モジュール VERSION MINOR + 個別検証。

### Phase 3: 書き込み系・クロスモジュール(1 週)

- §7.4 の backup/restore 4 対に §4.6 を実装(restore は「PDF に無ければ本体 backup を
  Warning 付きで参照」のフォールバック対称)。
- taskbar_config → sysprep_config/source の書き込み(§7.5)を PDF 経由に。
- driver_export / firewall_rule_export / startlayout_backup / default_app export 等の出力先。

### Phase 4: 締め・運用移行(逐次)

- strict mode: PDF 直下のマーカー(例: `overlay_policy.txt` = `strict`)で
  フォールバックを Error 化(移行完了プロファイルから順次)。
- メニュー単発実行への「データコンテキスト選択」追加の要否判断(§9)。
- hostlist の per-profile 化の要否判断(衛星 2 リポジトリへの波及があるため独立判断。§9)。
- 本体側 CSV をサンプル最小化(dev/template 同等の位置づけへ)。

---

## 7. 改修対象インベントリ(実測・file:line)

### 7.1 Test-Path 前置き(Phase 1 で解決 API 経由に修正・8 箇所)

放置すると「本体に無く PDF にある」構成で無言スキップ(2026-08 の license バグと同型)。

| # | ファイル:行 |
|---|---|
| 1 | modules/standard/firewall_config/firewall_config.ps1:55 |
| 2 | modules/standard/firewall_rule_config/firewall_rule_export.ps1:88 |
| 3 | modules/standard/local_user_config/local_user_config.ps1:16 |
| 4 | modules/standard/local_user_config/local_user_delete.ps1:16 |
| 5 | modules/standard/office_license_config/office_license_auth.ps1:66 |
| 6 | modules/standard/printer_delete/printer_delete.ps1:96 |
| 7 | modules/standard/printer_driver_config/printer_config.ps1:49 |
| 8 | modules/standard/windows_license_config/windows_license_install.ps1:65 |

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

---

## 9. 未決事項(Phase 0 で裁定)

| # | 論点 | 選択肢 | 暫定推奨 |
|---|---|---|---|
| Q1 | メニュー単発実行でのデータコンテキスト | (a) 常に無効(本体側のみ) / (b) 選択 UI を追加 | (a) で開始、Phase 4 で再判断 |
| Q2 | hostlist の per-profile 化 | (a) 見送り(全案件共通のまま) / (b) Phase 4 で PDF へ | (a)。衛星 2 リポジトリ波及が対価に見合うか要実運用データ |
| Q3 | pianist の profiles/(UI 操作プロファイル) | per-case 資材か framework 資産か | **裁定済み(2026-08-09): per-case 資材 = PDF 対象**(Studio の Pianist Profile Editor が案件コンテンツとして編集している実態とも整合) |
| Q4 | フォールバック Warning の表示强度 | Show-Warning / Show-Info | Show-Warning(profile-first 原則の担保) |
| Q5 | strict mode の指定方法 | マーカーファイル / プロファイル CSV 列 | PDF 直下マーカーファイル(プロファイル CSV スキーマ非接触) |
| Q6 | **FabriqStudio のワークスペースモデルとの関係** | (a) ワークスペース = 物理分離が必要な時のみ(別現場・別持出し PC)、PDF = 同一配備内の設定セット切替、と役割分担 / (b) 将来的にワークスペース切替を PDF 切替に統合 | (a) で開始。Studio 側の PDF 対応(編集先の切替 UI・レジストリ辞書等のエクスポート先)は fabriq P1 完了後に Studio 側タスクとして起票 |

### FabriqStudio との関係(2026-08-09 調査)

Studio(E:\fabriq_studio, WPF/.NET8)には本構想と交差する機能が既にある:

- **ワークスペース切替** = 現行の複数顧客対応(fabriq 丸ごとコピー切替)。本構想は同じ問題への
  別解であり、関係の裁定(Q6)が Phase 0 の凍結条件。
- **モジュール CSV への直接書き込み**: 端末管理 / レジストリ辞書エクスポート / INF→hostlist 転記。
  データが PDF に移ると Studio の書き込み先が変わる → Studio 側改修は本計画の**非スコープ**だが、
  §3 の完全ミラー構造 + PDF ルート名前空間予約により、Studio は「ルートを 1 つ挿し替えるだけ」で
  追随できる形を契約側で担保する。
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
  (Q2 を (b) にした場合のみ発生)
- FabriqStudio の PDF 対応実装(fabriq P1 完了後に Studio 側リポジトリで別途起票。
  本計画は Studio が追随可能な契約構造 — 完全ミラー + ルート名前空間予約 — の担保まで)
- evidence/ ツリーの per-profile 化(実行履歴・エビデンスは従来通り全体共有)

---

## 11. 進行管理

- TM: t-0095(本計画全体)。Phase 着手ごとに子タスクを切る(t-00xx)。
- 各 Phase の実装は CLAUDE.md 設計ゲート(P1 はフル版、P2 以降の各 wave は軽量版)を通す。
- kernel touched の全コミットで `powershell.exe -File ./dev/run_tests.ps1` を必須実行。
- 本書は実装と同コミットで随時更新し、契約変更は必ず §4 に反映してから実装する。
