# プロファイル別データオーバーレイ — 外部ツール向け契約ブリーフ

対象読者: fabriq のデータを読み書きする**外部ツール**（FabriqStudio ほか）の実装者。
正式仕様は [`dev/PROFILE_DATA_OVERLAY_PLAN.md`](PROFILE_DATA_OVERLAY_PLAN.md)（§3 構造 / §4 解決契約）と
[`kernel/KERNEL_API.md`](../kernel/KERNEL_API.md) §1.2。本書はそこからツールに関係する部分だけを抜いた要約。

最終更新: 2026-09-09（fabriq 側 Phase 1〜3 実装完了・Phase 4 は裁定により不実施）

---

## 1. 何が変わったか（3 行）

- プロファイル CSV `profiles/<name>.csv` に**同名フォルダ** `profiles/<name>/` を併設できるようになった。
- そのプロファイルの実行中だけ、**モジュールの設定 CSV・資材フォルダ・backup/export の出力先**が
  そちらへ切り替わる（無ければ従来どおりモジュール本体側。切り替わったかは実行画面に必ず表示される）。
- フォルダが無いプロファイル・メニューからの単発実行は**従来と完全に同一の動作**。既存運用は無改修。

以下このフォルダを **PDF (Profile Data Folder)** と呼ぶ。

---

## 2. 写像規則 — ここだけは正確に

**PDF レイアウトは `<tier>`（standard / extended）を落とす。** よって「ルートを差し替えるだけ」では
足りず、この写像そのものを実装する必要がある。

```
本体: <fabriqRoot>\modules\<tier>\<module>\<rel>
PDF : <fabriqRoot>\profiles\<name>\modules\<module>\<rel>
                                    ^^^^^^^ tier は入らない
```

例:

| 本体側 | PDF 側 |
|---|---|
| `modules\standard\gpo_config\gpo_list.csv` | `profiles\Cust_A\modules\gpo_config\gpo_list.csv` |
| `modules\standard\printer_driver_config\INF\` | `profiles\Cust_A\modules\printer_driver_config\INF\` |
| `modules\extended\pianist\profiles\` | `profiles\Cust_A\modules\pianist\profiles\` |

C# での実装イメージ（fabriq 側 `kernel/common.ps1` の `Get-FabriqOverlayCandidate` と等価）:

```csharp
// relPath は fabriqRoot からの相対パス（例 @"modules\standard\gpo_config\gpo_list.csv"）
// dataSetName が null/空 = データセット未選択 → 恒等写像（本体側をそのまま使う）
static string MapToDataSet(string fabriqRoot, string dataSetName, string relPath)
{
    if (string.IsNullOrWhiteSpace(dataSetName)) return relPath;

    var parts = relPath.Split('\\');
    // "modules" / <tier> / <module> [/ <rel...>] の形でなければ干渉しない
    if (parts.Length < 3 ||
        !parts[0].Equals("modules", StringComparison.OrdinalIgnoreCase)) return relPath;

    var module = parts[2];
    var rest   = string.Join("\\", parts.Skip(3));   // 空でもよい（= モジュールルート）

    var mapped = Path.Combine("profiles", dataSetName, "modules", module);
    return string.IsNullOrEmpty(rest) ? mapped : Path.Combine(mapped, rest);
}
```

補足:

- **`modules\` 配下以外は一切写像しない**（`kernel\csv\hostlist.csv` などは常に本体側。§5 参照）。
- 大文字小文字は Windows 準拠で不問。`..\` を含むパスは正規化してから判定する。
- **tier を落とす設計は確定**。理由は (1) 操作者が standard/extended を意識せずに PDF を組めること、
  (2) モジュールを tier 間で移動しても既存 PDF が壊れないこと。モジュール名は tier 間で一意。

---

## 3. 粒度契約 — 「部分的に混ぜる」は禁止

fabriq 側の解決は次の 3 粒度で、いずれも **all-or-nothing**。ツール側の UI もこれに合わせること
（「PDF に無いものは本体から補われる」と読める導線を作らない）。

| 対象 | 粒度 | 意味 |
|---|---|---|
| 設定 CSV | **ファイル単位** | PDF に同名 CSV があればそれを使う。無ければ本体側（警告表示） |
| 資材フォルダ（`INF/` `driver/` `certs/` `xml/` `file/` `source/` `wallpaper/` `assets/` `prompt/` `json/` `ppkg/` `backup/` 等） | **フォルダ単位** | PDF に**フォルダが存在すれば空でも**そちらを使う。本体側のファイルで穴埋めはしない（空フォルダ = 「資材ゼロ」の明示） |
| 複数 CSV 列挙（`reg_hklm_list*.csv` / `reg_hkcu_list*.csv`） | **モジュール単位** | PDF 側に 1 件でも一致があれば PDF 側のみ。本体側とは合成しない |

**帰結（Studio の UI で効いてくる点）**: 資材フォルダを PDF に作るなら、そのモジュールの資材は
**全部** PDF に入れる必要がある。「CSV だけ PDF に移してフォルダを忘れる」は、実行時に
「フォルダ不在 → Error」または「ファイル不在 → NOT FOUND」として現れる（無言では壊れない）。
移行導線を作るなら CSV と資材フォルダをセットで扱うこと。

---

## 4. 何が PDF 対象で、何がフレームワーク資産か

Studio が書き込む先を決めるときの判断表。**PDF ルート直下の `modules\` 以外の名前空間はツール用に
予約**してあるので、案件メタデータ（顧客名・作成日等）を `profiles\<name>\` 直下に置いてよい
（fabriq カーネルは `modules\` 配下しか見ない）。

| ファイル | 扱い |
|---|---|
| `<module>\*_list.csv` など**設定 CSV 全般** | **PDF 対象** |
| 資材フォルダ一式 | **PDF 対象** |
| `backup/` `driver/`（export 出力先） | **PDF 対象**（§6） |
| `modules\extended\pianist\profiles\` | **PDF 対象**（案件コンテンツと裁定済み。Pianist Profile Editor の保存先） |
| `module.csv` / `preset.csv` | **フレームワーク資産**（PDF に置かれても無視される） |
| `VERSION` / `REQUIRES_KERNEL` / `Guide.txt` / `test.psd1` / `*.ps1` / `tools\`(7z.exe 等) | **フレームワーク資産** |
| `kernel\csv\hostlist.csv` | **PDF 対象外（恒久）**。§5 |
| `profiles\<name>.csv`（プロファイル定義そのもの） | 従来どおり。フォルダ化しない |
| `evidence\` / `logs\` | PDF 対象外（実行履歴・エビデンスは全体共有のまま） |

---

## 5. hostlist は PDF に入れない（恒久確定）

per-profile hostlist は 2026-09-09 に**見送りで確定**した（保留ではなく結論）。理由:

1. 衛星ツール 2 本（fabriq_checksheet / fabriq_backuper）が **`kernel\csv\hostlist.csv` の存在によって
   fabriq ルートを発見**しており、さらに backuper は自分の `extended_hostlist` と本家 hostlist の
   `(OldPCname, NewPCname)` **集合完全一致ゲート**を持つ。衛星は単独起動（別 PC のこともある）なので
   「いまどのプロファイルが実行中か」を原理的に知り得ない。
2. ライフサイクルが違う。hostlist は**案件ごとに作り直すジョブ入力**、PDF は**顧客ごとに安定するレシピ**。
   ジョブ入力をレシピフォルダに入れると、案件のたびに PDF を書き換えることになる。

→ Studio の端末管理（hostlist 編集）は、**データセット選択の影響を受けない**。常に
`kernel\csv\hostlist.csv` を見る。

---

## 6. 書き込み側（backup / export）

fabriq 側は `Resolve-ModuleDataPath -ForWrite` で、**データセットが有効なら存在に関係なく PDF 側へ書く**
（backup → restore の対が PDF 内で閉じ、ある案件の捕捉物が別案件に混ざらない）。対象は
acl_config / reg_template / firewall_rule_config / desktop_icon_config / driver_config / default_app_config。

Studio がエクスポート系（レジストリ辞書 → `gpo_list.csv` 等）を書くときも同じ考え方でよい:
**データセットが選ばれていれば書き先は PDF 側**、選ばれていなければ本体側。

fabriq 側は「解決しただけではディレクトリを作らない」設計にしてある（プレビューやキャンセルで
終わった実行が空フォルダを残すと、後続の**読み**解決が「PDF にフォルダあり = all-or-nothing 発動」に
化けてしまうため）。Studio 側も**実際に保存するときだけ**ディレクトリを作ること。

---

## 7. Studio に実装してほしいこと（TM: fabriq_studio t-0022）

- **編集先データセットの選択**: 「なし（モジュール本体）」/ `profiles\<name>`。既定は「なし」= 現行動作。
- 選択中は**常時どこかに表示**する。取り違えて別顧客の設定を編集するのが最大のリスクなので、
  fabriq 本体も実行画面に採用元を必ず 1 行出す設計にしてある（同じ思想で）。
- 差し込み口は `IWorkspaceService.RootPath` + ルート相対パスで CSV を読み書きしている箇所
  （`ICsvService` / `Services\Master\MasterTargetResolver.cs` 付近）で、ほぼ 1 箇所に集約されている。
  そこに §2 の写像を挟むのが最小改修。
- 「PDF 側に無いので本体からコピーして作る」導線を出す場合は、**§3 の粒度**に合わせること
  （CSV はファイル単位でよいが、資材フォルダはフォルダごと）。
- CSV を新規作成する場合のエンコーディングは既存と同じ規約に従う（日本語を含むなら **UTF-8 BOM + CRLF**。
  fabriq 側は `-Encoding Default` で読むため、CP932 / UTF-8 BOM のどちらでも読めるが BOM 無し UTF-8 は不可）。
- `ENC:` 暗号化値はパス非依存なので、PDF 側 CSV でもそのまま機能する（追加対応不要）。

**なぜ急ぐ必要があるか**: fabriq 側の `apps/csv_editor` は撤去方針になったため、
**PDF 側 CSV を編集する GUI が現状ゼロ**。それまでは Excel / エクスプローラでの直接編集になる。

---

## 8. 決着済みで「やらない」こと（再提案しないでほしい一覧）

| 論点 | 決着 |
|---|---|
| ワークスペース切替と PDF の統合 | **しない**。ワークスペース = 物理分離が要るとき（別現場・別持出し PC）、PDF = 同一配備内の設定セット切替、と役割分担する |
| hostlist の per-profile 化 | **しない**（§5） |
| strict mode（フォールバックを Error 化） | **不採用**。可視化は Show-Warning + テレメトリ + HTML チェックリストの 3 点で最終形 |
| メニュー単発実行でのデータセット選択 | **不採用**。メニュー単発は常にモジュールデフォルト CSV を使う |
| 本体側 CSV の一斉縮退 | **しない**。使わないが消しもせず、自然消滅に委ねる |
| csv_editor のオーバーレイ対応 | **しない**（撤去候補） |

---

## 9. うれしい副作用（更新運用）

`profiles/` は `dev/framework_overlay_rules.json` の `excludeDirsRecursive` に入っており、
**フレームワーク更新・モジュール更新のどちらでも「いかなる場合も保全」される唯一の名前空間**。
モジュール側の `*_list.csv` はパッチ生成時に strip されるので保全されるが、**資材フォルダは
strip 対象ではない**（パッチ適用は追加コピーなので消えはしないが、同名ファイルは上書きされるし
保全の契約はない）。

→ 案件データが PDF に寄るほど、Studio の「fabriq オーバーレイ更新」機能は安全かつ単純になる。

---

## 10. FabriqStudio 側の調査への回答（2026-09-09）

Studio 側 TM `t-0022` の調査（プリンタドライバ検出への影響ほか）を読んだうえでの、fabriq 側からの回答。

### 10-1. 【訂正】「CSV は PDF・INF は本体」でも**警告は出る**

> Studio の記載: *「driver_list.csv（ファイル単位）と INF（フォルダ単位）で解決粒度が違うため、
> 『CSV は PDF・INF は本体』というちぐはぐが fabriq 的には正常系として成立する（警告が出ない）」*

**これは誤り。この状態では実行時に警告が出る。** 実測（dev 機、PDF に `printer_list.csv` のみを置いた状態）:

```
[INFO]    [DATA] printer_driver_config/printer_list.csv <- profile (Cust_A)
[WARNING] [DATA] printer_driver_config/INF/ <- module dir (FALLBACK: not in profile data folder)
```

解決は**採用元がどちらであっても必ず 1 行表示**する契約（§4.3）で、本体側へ落ちた場合は必ず
`Show-Warning`。「粒度が違う」ことと「無言になる」ことは別で、無言経路は存在しない。

ただし**これは実行時の可視化**であり、Studio の編集時点では見えない。編集画面で
「この CSV は PDF・この資材は本体」を示す不整合表示には依然として価値がある（実行前に気づける）。
compensating control としてではなく、前倒しの UX として実装するのが正しい位置づけ。

### 10-2. INF の all-or-nothing（Studio 指摘 3）— 設計意図と運用の答え

指摘のとおり **PDF 側に `INF/` を作った瞬間、本体側の INF は 1 件も読まれない**（空でも PDF が勝つ）。
これは仕様であって回避策は用意していない。理由は「どのドライバがどこから来たか」を操作者が
追えなくなることを避けるため（§3 の設計理由と同一）。

fabriq 側が想定している運用は次のどちらか。**中間はない。**

| 選択 | 意味 |
|---|---|
| PDF に `INF/` を作る | その案件で使うドライバを**すべて** PDF に入れる（案件のドライバセットが PDF 内で完結） |
| PDF に `INF/` を作らない | 本体側の共通ドライバ庫をそのまま使う（フォールバック警告が毎回出る＝Profile-First 的には非推奨だが**正常に動く**） |

70 MB 全コピーが非現実的という指摘は妥当で、その場合の答えは「**PDF に `INF/` を作らない**」。
案件専用ドライバが 1 機種だけ要るなら、その 1 機種を PDF に入れるのではなく、
**本体庫に足して PDF には `INF/` を作らない**ほうが破綻しない。

→ Studio の「本体の INF（n 件）は使われなくなる。必要な分を一緒に取り込むか」を聞く導線は
**この二択を操作者に見せるもの**として設計するのが正確。「一部だけ取り込む」を既定の選択肢にしないこと。

### 10-3. 2 プロファイル間の CSV 二重化（Studio P4 の裁定事項）

**fabriq 側に PDF 間の共有・継承の仕組みは無い**（意図的に作っていない。all-or-nothing の provenance が
崩れるため）。マスタ設計が 2 プロファイル（マスタ / sysprep）を吐く場合、選択肢は次の 3 つで、
**どれを採るかは Studio 側の裁定**。fabriq 側はいずれも受け入れる。

1. 共有したい CSV は **PDF に置かない**（本体側に残す＝フォールバック。警告は出るが動作は正常）
2. Studio が**両方の PDF に同じ内容を書く**（Studio のモデル上は単一ソース、ディスク上は二重化）
3. 2 プロファイルを 1 プロファイルに統合する（可能なら最も単純）

fabriq 側の推奨は **2**。ディスク上の二重化は「その PDF だけ見れば何が適用されるか完結して分かる」
という PDF の一番の価値を保つため。1 は共有部分が Profile-First から外れる。

### 10-4. マスタ設計のスナップショット走査（`MasterProfileGeneratorService.cs:95`）

「PDF 合成が要る」という読みで正しい。走査が「いま何が設定されているか」を報告するものである以上、
**カーネルと同じ解決規則を適用した結果**を見せる必要がある（§2 の写像 + §3 の粒度）。
特に「INF に対応するフォルダが無い」系の警告は、PDF 側を見ずに判定すると誤爆する。

### 10-5. printer_driver INF は fabriq 側で実機未検証（Studio 指摘 8）

そのとおり。fabriq の VM リグでは実ドライバを要するため未被覆と明記している
（`PROFILE_DATA_OVERLAY_PLAN.md` §16.4）。解決ロジックは他 12 モジュールと同一イディオムで
Pester と dev 機チェックが被覆しているが、**実ドライバでの通し確認は未実施**。
Studio が最初の実利用者になる可能性が高いという読みは妥当で、P3 で実機確認を 1 回入れる方針に賛成。
その際の確認点は 3 つ:

1. PDF 側 `INF\` からドライバが認識・インストールされること
2. `.exe` / `.zip` をアーカイブで置いた場合、**展開先も PDF 内**になること
   （`printer_driver_install.ps1` は `INF\<BaseName>\` に展開し、`:285` の封じ込めガードも
   解決後ルートに追随する）
3. `tools\7z.exe` は本体側のままでよいこと（フレームワーク資産）

### 10-6. Studio の方針で fabriq 側と整合が取れている点（確認）

- 「プロファイル画面経由 = PDF / モジュール編集メニュー経由 = 本体」— 妥当。fabriq 側の
  「プロファイル実行だけがデータセットを切り替える」（Q1 恒久確定）と対称になっている。
- 「モジュール追加時に本体から as-is コピー、全件一括コピーはしない」— 妥当。
  §3 の粒度と矛盾せず、「取り込みで実行結果が変わらない」不変条件も正しい
  （as-is コピーなら PDF が勝っても内容が同一なので挙動不変）。
- 「資材フォルダは実際に置くときだけ作る」— **重要かつ正しい**。fabriq 側も同じ理由で
  `-ForWrite` がディレクトリを作らない設計にしてある（§6）。空フォルダを先に作ると
  all-or-nothing が発動して本体側の資材が落ちる。
- 「テンプレートは過去案件の PDF を複製で足りる」— 妥当。fabriq 側に雛形機構は無いし、作る予定もない。

---

## 11. 参照

- 正式仕様: `dev/PROFILE_DATA_OVERLAY_PLAN.md`（§3 構造 / §4 解決契約 / §9 裁定一覧 / §12〜§14 実施記録）
- 公開 API: `kernel/KERNEL_API.md` §1.2（`Resolve-ModuleDataPath` / `Get-ModuleDataFiles`）、§3.2（`FABRIQ_PROFILE_DATA_DIR`）
- 実装本体: `kernel/common.ps1` の `Get-FabriqOverlayCandidate` / `Resolve-ModuleDataPath` / `Get-ModuleDataFiles`
- 操作者向け説明: `README.md`「プロファイル別データオーバーレイ」節
