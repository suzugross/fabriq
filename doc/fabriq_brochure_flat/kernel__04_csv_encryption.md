# CSV 駆動 + 暗号化アーキテクチャ

fabriq 全体は CSV 駆動で動く。ホスト情報・モジュール設定・プロファイル定義・カテゴリマスタ・作業者・ログ配送先 — すべて CSV。機密値はフィールド単位で暗号化して同じ CSV に共存させる。

---

## 暗号化仕様（全モジュール共通）

| 項目 | 値 |
|---|---|
| 鍵導出 | PBKDF2-HMAC-SHA256, **100,000 iterations**, 固定ソルト `fabriq-fixed-salt-2024` |
| 暗号アルゴリズム | AES-256-CBC, PKCS7 padding |
| エンコード | UTF-8（平文）→ Base64（暗号文） |
| プレフィックス | `ENC:<Base64>` |
| 鍵長 | 32 bytes (AES-256) |
| IV 長 | 16 bytes (AES block size) — PBKDF2 から派生（鍵と同じ KDF stream を再利用） |

### `Unprotect-FabriqValue` の実装ポイント（common.ps1 §445）

- `ENC:` プレフィックスが無い値はそのまま返す（`Import-ModuleCsv` で混在容認）
- `Rfc2898DeriveBytes` で key と iv を**一回の KDF 流れから連続取得**（C# `CryptoPoC` と完全互換）
- `MemoryStream` → `CryptoStream` → `StreamReader` で UTF-8 復号
- すべての CryptoStream / Aes / KDF オブジェクトは `Dispose` 確実

### パスフレーズ検証トークン（`kernel/txt/passphrase_verify.txt`）

- 中身: `ENC:<Base64>` 形式で、平文 `"surkitinisme"` を当該パスフレーズで暗号化したもの
- Fabriq Studio がパスフレーズ初回設定時に生成
- `Test-MasterPassphrase` が起動時にこのトークンを復号 → 結果が `"surkitinisme"` ならパスフレーズ正解
- このファイルが無いと Fabriq は起動できない（exit 1）。Studio が必ず先行する設計

### Resume 時のパスフレーズ持ち回り（DPAPI）

再起動跨ぎは AES とは別経路：

- `Protect-PassphraseForResume`: パスフレーズを **DPAPI LocalMachine** で暗号化して Base64 文字列化、`resume_state.json` の `ProtectedPassphrase` フィールドに格納
- `Unprotect-PassphraseFromResume`: 同マシンの DPAPI で復号
- LocalMachine スコープなのでマシンが変わると復号できない（盗難 PC で resume が走らない安全弁）
- DPAPI 復号失敗時は手動パスフレーズ再入力にフォールバック（最大 3 回）

---

## Import-ModuleCsv（モジュールが必ず通す統合パイプライン）

```powershell
$rows = Import-ModuleCsv -Path $listCsv `
                        -FilterEnabled `
                        -RequiredColumns @("Enabled","Path","Type") `
                        -Segment $env:FABRIQ_SEGMENT
```

### 4 段階パイプライン

```
1. Import-CsvSafe
   ├── Test-Path 不在 → Show-Error + return $null
   └── Import-Csv -Encoding Default で文字化け回避（PS5.1 のシステム既定 = SJIS / UTF-8 自動判別）

2. 透過復号
   ├── 各 row の各 prop を走査
   ├── 文字列値で `ENC:` 始まりなら Unprotect-FabriqValue
   └── 失敗時は Show-Warning + 元の暗号文のまま温存（モジュール側で判断可能）

3. RequiredColumns 検証（指定時のみ）
   ├── 1 行目の PSObject.Properties.Name で列存在チェック
   └── 欠落あれば Show-Error + return $null（モジュール先頭で空 return パターン）

4. FilterEnabled + Segment フィルタ
   ├── Enabled='1' で絞る（FilterEnabled 指定時のみ）
   ├── Segment 列があれば $env:FABRIQ_SEGMENT と厳密マッチ（空 vs 空もマッチ）
   └── ヒット 0 件なら Show-Skip + return @() （モジュールは 0 件を Skipped として返す）
```

### Segment の使い分け例

```csv
# wallpaper_list.csv
Enabled,Segment,WallpaperPath
1,office,\\share\wallpapers\office.png
1,home,\\share\wallpapers\home.png
1,,                                       ← Segment 空 = "デフォルトグループ"
```

```csv
# profile.csv
Order,ScriptPath,Enabled,Description,Segment
10,standard/wallpaper_config/wallpaper_config.ps1,1,オフィス用壁紙,office
20,standard/wallpaper_config/wallpaper_config.ps1,1,ホーム用壁紙,home
30,standard/wallpaper_config/wallpaper_config.ps1,1,既定壁紙,
```

3 行とも同じモジュールを呼ぶが、`FABRIQ_SEGMENT` 経由で異なる `_list.csv` 行に分岐する。

---

## CSV エンコーディング規約

| ファイル | エンコード | 改行 | 備考 |
|---|---|---|---|
| `kernel/csv/*.csv` | UTF-8 BOM | CRLF | hostlist / workers / categories / log_destinations / manifesto |
| `modules/*/module.csv` | UTF-8 BOM | CRLF | カーネルが `Import-Csv -Encoding Default` で読む |
| `modules/*/_list.csv` | UTF-8 BOM | CRLF | 各モジュール固有の設定 CSV |
| `modules/*/preset.csv` | UTF-8 BOM | CRLF | Studio 用ドロップダウン UI 定義 |
| `profiles/*.csv` | Default (SJIS or UTF-8 BOM) | CRLF | `Import-Csv -Encoding Default` で読み書き |

PS5.1 + Windows の制約で「日本語含む CSV は UTF-8 BOM が安全」が運用ルール（feedback memory `feedback_ps1_utf8_bom`）。

---

## CSV ファイルの分類（更新オーバーレイから見た 3 区分）

### 1. **Framework CSV**（テンプレ → ターゲットへ overlay 対象）

| ファイル | 役割 | 配置 |
|---|---|---|
| `module.csv` | モジュールメニュー定義（MenuName, Category, Order, Script, Enabled） | `modules/{type}/<name>/module.csv` |
| `preset.csv` | Studio 用ドロップダウン UI（Column, Value, Label） | `modules/{type}/<name>/preset.csv` |

`dev/framework_overlay_rules.json` の `moduleCsvWhitelist` に列挙。

### 2. **Site-Specific CSV**（絶対保護、上書き禁止）

| ファイル | 中身 | 保護理由 |
|---|---|---|
| `kernel/csv/hostlist.csv` | 対象 PC マスタ（`ENC:` 暗号化済の機密フィールド含む） | 顧客固有 |
| `kernel/csv/workers.csv` | 作業者マスタ | 顧客固有 |
| `kernel/csv/log_destinations.csv` | ログ配送先 + 認証情報（`ENC:`） | 顧客固有 |
| `modules/*/_list.csv` | 各モジュールの設定データ（reg_template, app_config 等） | キッティング案件ごとに作る |
| `profiles/*.csv` | 実行プロファイル | 案件ごとに作る |

`framework_overlay_rules.json` の `excludeFilesKernelLevel` + `excludeDirsRecursive` ("profiles") + ホワイトリスト除外で守られる。

### 3. **Runtime Artifact**（実行時生成、配備からは除外）

| ファイル | 役割 |
|---|---|
| `kernel/json/status.json` | Status Monitor のライブ状態（atomic write） |
| `kernel/json/session.json` | 現セッション情報 |
| `kernel/json/resume_state.json` | 再起動跨ぎの状態スナップショット |
| `kernel/json/art_pulse.txt` | Show-* が +1 する鼓動カウンタ |
| `kernel/json/skip_request.flag` | async モジュール強制スキップ要求 |
| `kernel/txt/passphrase_verify.txt` | パスフレーズ検証トークン |
| `kernel/txt/silence.flag` | ART 演出抑制 flag |

---

## hostlist.csv の構造（運用上もっとも重要）

```csv
AdminID,OldPCName,NewPCName,
EthernetIP,EthernetSubnet,EthernetGateway,
WifiIP,WifiSubnet,WifiGateway,
DNS1,DNS2,DNS3,DNS4,
Pin,
Printer1Name,Printer1Driver,Printer1Port,
Printer2Name,...,Printer10Port
```

- `AdminID`: 管理 ID（一級識別子。実行履歴 / HTML チェックリスト / Restore-ExecutionHistory のフィルタキー）
- `OldPCName` / `NewPCName`: 旧 PC 名 / 新 PC 名（hostname_config が NewPCName へリネーム）
- 機密フィールドは Studio で `ENC:<Base64>` に暗号化可能（Pin / DNS / 各 Printer 等）
- ホスト選択時 `Set-SelectedHostEnvironment` がすべて env vars へ流し込み（ENC は復号して入る）

---

## 暗号化設計の哲学

> 「**鍵を分散させない、ファイルベースの透過復号で鍵管理を 1 つに集約**」

- 鍵は単一のマスターパスフレーズのみ（PBKDF2 で派生）
- パスフレーズの検証はトークン 1 つ（`passphrase_verify.txt`）
- すべての CSV は同じ鍵で読まれるため、operator の心理的負荷が小さい
- Studio が暗号化 / 復号 / 検証トークン生成を一括で担当する → fabriq 本体はパスフレーズを「使う」だけ
- 固定ソルトの選択は意図的（鍵をマスターパスフレーズだけに依存させ、複数 PC 間で同じ ENC 値が同じ平文に復号できるようにするため。Salt をユニークにすると配布フェーズで矛盾が起きる）

セキュリティ詰めの甘さの棚卸しは `project_crypto_security_review.md` で別途追跡されている（A: 単独で直せる衛生 / B: Studio 連携要する format migration）。
