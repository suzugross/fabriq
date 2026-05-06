# Fabriq Brochure Materials — Index

**生成日**: 2026-05-06
**fabriq バージョン**: 3.2.2
**目的**: 別 Claude（チャットベース）が技術パンフレット化するための原材料一式

このディレクトリには、fabriq Windows キッティング自動化フレームワーク（PowerShell + WinForms）の技術仕様を細粒度で文書化した素材を配置した。チャットベース Claude にこれを丸ごと渡せば、パンフレット・スライド・営業資料・技術提案書のいずれの体裁にも編集可能。

---

## 運用方式：2 ファイル受け渡し（INDEX.md + SOURCE.md）

claude.ai のチャットには次の **2 ファイルだけ** 添付すれば全素材が届く：

| ファイル | サイズ | 役割 |
|---|---|---|
| **`INDEX.md`** | ~17KB | 本ファイル。全章の目次 / パンフレット化推奨構成 / 哲学・運用方針 / caveats |
| **`SOURCE.md`** | ~488KB | 全 100 章を 1 ファイルに連結したコンテンツ本体 |

### SOURCE.md 内の章の探し方

`SOURCE.md` 内の各章は以下の見出しで始まる：

```
# === <path/to/file.md> ===
```

例:
- `# === kernel/02_public_api.md ===` — 公開 API §1-§5 解説
- `# === contracts/module_result.md ===` — ModuleResult 契約
- `# === modules/hostname_config.md ===` — hostname_config モジュール

下の「構成」節のファイルパス記述がそのまま `SOURCE.md` 内の章見出しに対応。チャットベース Claude には「`SOURCE.md` の `# === modules/evidence_config.md ===` 章を読んで...」のように指示できる。

### チャット Claude への定型指示（テンプレ）

```
添付の INDEX.md と SOURCE.md は fabriq Windows キッティング自動化フレームワークの
技術仕様素材です。INDEX.md の「構成」節と「パンフレット化のおすすめ構成」を読み、
SOURCE.md 内の対応する `# === <path> ===` 章を必要に応じて参照して、
<目的の体裁> を作成してください。
```

---

## 構成

```
e:\tmp\fabriq_brochure_materials\
├── INDEX.md                              ← 本ファイル
├── kernel/                               ── カーネル詳細（11 ファイル）
│   ├── 01_overview.md                    ── カーネル全体像 / 責務分担 / 保管場所マップ
│   ├── 02_public_api.md                  ── 公開 API §1-§5 解説
│   ├── 03_orchestration.md               ── main.ps1 / Invoke-BatchExecution / FlexProfile
│   ├── 04_csv_encryption.md              ── CSV 駆動 + AES-256-CBC + パスフレーズ検証
│   ├── 05_resume_restart.md              ── __RESTART__ / RunOnce / DPAPI / resume_state.json
│   ├── 06_status_monitor.md              ── 別プロセス WinForms モニタ + ART 演出
│   ├── 07_evidence_history.md            ── スクリーンショット / 履歴 / HTML チェックリスト
│   ├── 08_async_execution.md             ── __ASYNC__ + Runspace + Skip / Timeout
│   ├── 09_versioning.md                  ── SemVer / KERNEL_VERSION / REQUIRES_KERNEL / リリース手順
│   ├── 10_function_index.md              ── 90+ 関数の公開/内部分類インデックス
│   └── 11_directory_layout.md            ── 全ディレクトリ構成 + 配備境界
│
├── modules/                              ── モジュール 75 件（per-module 1 ファイル + 全体図）
│   ├── 00_modules_overview.md            ── カテゴリ別の全 75 モジュール一覧 + 統計
│   └── <module_name>.md (× 75)           ── 各モジュールの目的/CSV/ステップ/注意/検証
│
├── apps/                                 ── apps + commands + dev tooling（5 ファイル）
│   ├── 00_apps_overview.md               ── apps/ ディレクトリ全体図
│   ├── 01_fabriq_operator_dashboard.md   ── メインダッシュボード深掘り
│   ├── 02_fabriq_ios.md                  ── Cisco IOS 風シェル（独立 SemVer）
│   ├── 03_other_apps.md                  ── csv_editor / system_launcher 等 7 件
│   ├── 04_dev_template_and_tooling.md    ── dev/template / build_framework_patch / seed_module_versions 等
│   └── 05_commands.md                    ── commands/ 配下 6 ファイル
│
├── profiles/                             ── プロファイル一覧
│   └── 00_profiles_overview.md           ── Master_Pre / Master_Config / sysprep / easy_template 等
│
└── contracts/                            ── 公開契約 5 ファイル
    ├── profile_csv_schema.md             ── Profile CSV スキーマ + 特殊マーカー（KERNEL_API.md §4）
    ├── module_result.md                  ── ModuleResult 契約（KERNEL_API.md §5）
    ├── special_markers.md                ── 5 マーカー詳細仕様 + 削除 4 マーカーの経緯
    ├── host_environment.md               ── SELECTED_* / FABRIQ_* 環境変数（KERNEL_API.md §3）
    ├── overlay_contract.md               ── 更新オーバーレイ契約（KERNEL_API.md §9）
    └── evidence_manifest_contract.md     ── manifest.json 公開契約（KERNEL_API.md §10 / EVIDENCE_MANIFEST.md）
```

---

## パンフレット化のおすすめ構成

### A. 営業向け 1 ページ A4

- 表紙: 概要（README L1 のキャッチ + ver / 「**Manifeste du Surkitinisme**」）
- 機能ハイライト: kernel/01_overview.md の責務テーブル + modules/00_modules_overview.md の数字
- 安全性: kernel/04 の暗号化仕様 + kernel/07 のエビデンス
- 導入実績の枠: 空（顧客データ別途）

### B. 技術提案書（10–20 ページ）

1. **概要章**: kernel/01_overview.md
2. **アーキテクチャ章**: kernel/03_orchestration.md（フローチャート図表化）
3. **機能カタログ章**: modules/00_modules_overview.md → 主要 10 モジュールのみ抜粋
4. **セキュリティ章**: kernel/04_csv_encryption.md + contracts/host_environment.md
5. **エビデンス章**: kernel/07_evidence_history.md + contracts/evidence_manifest_contract.md
6. **運用章**: kernel/05_resume_restart.md + kernel/06_status_monitor.md
7. **管理ツール章**: apps/01_fabriq_operator_dashboard.md（FlexProfile を強調）
8. **拡張性章**: kernel/09_versioning.md + contracts/overlay_contract.md
9. **API リファレンス**: kernel/02_public_api.md + kernel/10_function_index.md
10. **付録**: contracts/profile_csv_schema.md / contracts/module_result.md / contracts/special_markers.md

### C. 開発者向け技術ホワイトペーパー（30–50 ページ）

すべての kernel/* と contracts/* を順序通り編集 → 各モジュールのうち代表的な 20 件を deep-dive として展開。

---

## 重要メタデータ

| 項目 | 値 |
|---|---|
| **エントリ言語** | PowerShell 5.1（Windows 標準）+ WinForms GUI + C# ランチャ |
| **対応 OS** | Windows 11 |
| **必要権限** | 管理者（Fabriq.exe が UAC で自動昇格） |
| **コードライン数** | kernel/common.ps1: 4,371 行 / kernel/main.ps1: 1,913 行 |
| **公開 API 関数** | 約 18 関数（KERNEL_API.md §1） |
| **内部関数** | 約 75 関数 |
| **モジュール総数** | 75（標準 60 + 拡張 15） |
| **Post-Apply Verification 実装率** | 33%（25/75）+ 意図的除外 15 件 |
| **暗号化** | AES-256-CBC + PBKDF2-SHA256 (100,000 iter, 固定ソルト) |
| **Resume 機構** | __RESTART__ + RunOnce + DPAPI passphrase + resume_state.json (v1/v2) |
| **公開契約** | 4 つ: Public API（§1-§5）/ Profile CSV（§4）/ Overlay（§9）/ Manifest（§10） |
| **依存関係** | Fabriq Studio（パスフレーズ生成）に依存 / fabriq_evidence_manager（外部 consumer） |
| **同梱サードパーティ** | 7-Zip 25.01（LGPL）— printer_driver_config 用 |

---

## 哲学・運用方針（パンフレットの世界観で必ず触れたい論点）

### 1. **B2B/B2G キッティング業務の置換**

fabriq は顧客ごとに「使い捨ての batch 群」を毎回作っていた業務を、共通フレームワークで置き換える設計（user memory `user_business_context`）。「カスタマイズ可能な汎用基盤 + プロファイルで案件別シナリオ」のパターン。保守スコープは含まない（user 提供のみで運用は顧客 / 作業者）。

### 2. **保守性・堅牢性・柔軟性が最優先**

CLAUDE.md の最初の行に明記。`dev/template` テンプレート厳守 + `kernel/common.ps1` 共通関数の徹底活用 + 既存パターンの踏襲（車輪の再発明禁止）が三大ルール。

### 3. **AutoPilot は完全無人ではない**

「**確認スキップ + auto-resume**」という運用上重要な区別（feedback memory `feedback_autopilot_wording`）。operator は脇で見ていて Esc できる前提。AutoPilot は「人間レス」ではなく「介入レス」を意味する。

### 4. **公開契約と内部実装の分離**

`KERNEL_API.md` で `§1-§5` の公開 API のみ宣言。それ以外（90+ の内部関数）は PATCH バージョンでも自由に変更可能。これにより kernel 開発の自由度を確保しつつ、モジュール側は最小限の依存だけで済む。

### 5. **エビデンス自動取得 + 公開 manifest 契約**

すべてのモジュール実行はスクリーンショット PNG + 履歴 CSV 行 + HTML チェックリストへ反映。`evidence_config` は 22 セクションのシステム情報を `manifest.json`（公開契約 schemaVersion=1）と共に出力し、外部 evidence consumer（`fabriq_evidence_manager`）が前方互換に消費可能。

### 6. **fabriq_studio との疎結合**

fabriq 本体は Studio のバージョン・機能に依存しない。Studio はパスフレーズ生成・CSV 編集・プロファイル管理・更新オーバーレイを担当するが、本体側は契約（`framework_overlay_rules.json` schemaVersion=1）越しにしか Studio を見ない。

### 7. **演出文化（Manifeste du Surkitinisme）**

ART pulse / sentences / silence.flag / manifesto / fabriq_ios（Cisco IOS 風シェル）など、業務ツールには珍しい演出機能を持つ。「surkitinisme」は **キッティング超克主義** の造語。fabriq の名前と哲学の核。`silence.flag` で演出を止められる opt-out 設計が運用配慮。

### 8. **3 つの入り口の収斂**

`tonebender`（左シフト：MDM 連携の入口）+ `PPKG`（プロビジョニング側）+ `fabriq`（本体）が converge する設計（user memory `project_solo_dev_scope`）。Right shift（fleet/SaaS）はスコープ外、ソロ開発の意思決定。

### 9. **画像認識 RPA の拒絶**

冪等性 / 検証可能性が壊れるため image-recognition RPA は受け入れない（user memory `project_rpa_rejected`）。UIA + AutomationId のみが許容形式。fabriq の信頼性の根幹に関わる方針。

---

## 補足: ナレッジベース参照（user memory より）

このパンフレット化作業に関連する周辺情報：

- **クリプト見直し未着手**: `project_crypto_security_review` — A 群（単独で直せる衛生）と B 群（Studio 連携要する format migration）に分類済、未実施
- **Effect Domain (α/β/γ) 設計済 / 採用見送り**: `project_effect_domain_deferred` 2026-04-25
- **NextGProfile（並列/ループ/IF）見送り**: `project_nextgprofile_deferred` 2026-04-26（過剰機能と判断）
- **fabriq は機能完成・feature-hunting フェーズ**: `feedback_feature_hunting_phase` — 新機能提案は保守的に
- **FrexProfile 設計**: `project_frex_design` — 実行=AutoPilot常時 / finalize=手動常時、Order が一級識別子
- **Pianist は extended モジュール**: `project_pianist_module_decision` 2026-05-02 — apps/ から昇格、fabriq 統合完成

---

## モジュールカタログの読み方

各 `modules/<name>.md` は以下の標準構造：

```markdown
# <module_name> (Standard|Extended)

**カテゴリ**: ...
**メニュー名**: ...（複数の場合あり）
**VERSION**: x.y.z / **REQUIRES_KERNEL**: x.y.z
**Post-Apply Verification**: 実装あり / なし / 不可
**サブスクリプト**: 補助 .ps1 ファイルとその役割

## 目的
（Guide.txt から抽出した 2–4 文の要約）

## 入力 (CSV)
（_list.csv の列スキーマ）

## 主要ステップ
（.ps1 の Step 1..7 概要）

## 注意点・運用メモ
（権限・再起動要 / 副作用 / 既知の制約）

## 検証
（Post-Apply Verification の中身、または除外理由）
```

このフォーマットで 75 ファイルが揃っているため、パンフレット化時は欲しいモジュールだけ pluck して使える。

---

## 既知の caveats（編集時に注意）

1. **`looper_list.csv` の日本語化け**: `modules/extended/script_looper/looper_list.csv` に文字化けセル（`NIC設定リトライ�E�EHCP解放征E���E�E`）あり。CSV BOM/encoding 事故の疑い。誰かが触る時にフラグを立てておく
2. **README の Extended 件数 (14)**: 現状は 15 件（pianist が 2026-05-02 に昇格して以降の数字に追従していない）
3. **kernel 3.0.0 で削除されたマーカー**: `__SHUTDOWN__` / `__PAUSE__` / `__STOPLOG__` / `__STARTLOG__` を旧プロファイル例として書く際は「graceful degradation」で動くが廃止と明記
4. **Verification 統計**: agents が "実装あり" と書いた件数は 25 と記載しているが、`-Verified` 引数を渡しているかどうかで微妙に解釈差があり得る（office_license_config / office_update のように内部検証はあるが `-Verified` 渡していない例）
5. **commands/system_launcher.ps1 が apps/system_launcher/system_launcher.ps1 と重複**: 同一スクリプトが 2 箇所に存在、起動経路が異なる。集約候補
6. **commands/temp_command.ps1 は意図的な空 stub**: operator が ad-hoc per-site 作業のために埋める scratch slot（"未実装" ではない）
7. **fabriq_ios のモード数**: SPEC.md は 4 モード記載だが現状は **5 モード**（ModuleConfig が後から追加）。SPEC は更新漏れの可能性
8. **Fabriq_IOS.exe は独立ランチャ**: `dev/launcher/` に `Launcher.cs` と `Launcher_IOS.cs` の **両方**あり。fabriq_ios は Fabriq.exe ではなく専用 .exe を持つ。トップレベル exe 2 つの構成
9. **fabriq_ios VERSION = 0.3.5**: kernel から独立した SemVer。pre-1.0 の作家性プロジェクト位置付け
10. **profile CSV の path 区切りが混在**: `Master_Pre02.csv` 等は `/`、`Custom Plan.csv` は `\`。kernel は両対応するが、shipped templates 内で statement のスタイル不統一あり

---

## 連絡先

- **作者**: yuki.suzuki@suzugross.com
- **ライセンス**: MIT（fabriq 本体）
- **リポジトリ**: e:\fabriq

---

このディレクトリのすべてのファイルを単純にチャットベース Claude に投げるか、必要な節だけ編集してパンフレット化してください。素材は粒度を意図的に大きく取ってあるので、削るのは簡単に / 足すのは元のソースファイルから直接、という運用が想定されています。
