# Evidence Manifest 契約（External Evidence Consumer ↔ fabriq）

`KERNEL_API.md §10` および `kernel/EVIDENCE_MANIFEST.md` で公式宣言された、**外部 evidence consumer ツール**（代表: `fabriq_evidence_manager`、別プロジェクト C#/WPF/.NET8）が前方互換に消費するための公開契約。

`evidence_config` モジュール v1.3.0+ が `pc_information/<dir>/manifest.json` を出力。kernel 2.2.2 で contract 公開化。

---

## ファイル配置

```
{evidenceBaseDir}/pc_information/{collectionDir}/manifest.json
```

`{collectionDir}` は `{date}_{uid}_{pc}` 命名（legacy / unified どちらも同じ）。

- 1 evidence_config 実行 = 1 manifest.json
- 再実行時は既存 manifest を `manifest.json.bak` に rename してから上書き（**1 世代のみ保持**、`.bak.bak` は作らない）
- manifest 内のファイルパスは **manifest.json 自身からの相対パス**（self-contained）

---

## schemaVersion=1 のスキーマ

```json
{
  "schemaVersion": 1,
  "manifestType": "fabriq-evidence-manifest",
  "evidenceConfigVersion": "1.3.0",
  "fabriqKernelVersion": "3.2.2",
  "collectedAt": "2026-04-25T13:28:39+09:00",
  "computerName": "NEW-PC-01",
  "hardwareUniqueId": "T2NXCV06Y22208C",
  "selectedNewPcName": "NEW-PC-01",
  "workerName": "suzuki",
  "sections": [
    {
      "id": "01",
      "title": "System Basic Info",
      "files": ["01_SystemInfo.txt"],
      "status": "Success",
      "reason": null,
      "elapsedMs": 145
    },
    {
      "id": "14",
      "title": "Server Roles & Features (CSV)",
      "files": [],
      "status": "Skipped",
      "reason": "Client OS detected (Server-only section)",
      "elapsedMs": 12
    },
    {
      "id": "22",
      "title": "Office License / Activation Status",
      "files": ["22_OfficeLicense.txt"],
      "status": "Success",
      "reason": null,
      "elapsedMs": 8200
    }
  ],
  "summary": {
    "sectionCount": 23,
    "successCount": 21,
    "skippedCount": 2,
    "failedCount": 0,
    "partialCount": 0
  }
}
```

### トップレベルフィールド

| Field | Type | Required | Description |
|---|---|---|---|
| `schemaVersion` | int | yes | 現行 `1`、破壊的変更時に `2` |
| `manifestType` | string | yes | 固定値 `"fabriq-evidence-manifest"`（type discrimination 用） |
| `evidenceConfigVersion` | string | yes | manifest を書いた evidence_config モジュールの SemVer |
| `fabriqKernelVersion` | string | yes | manifest 書き込み時点の `kernel/KERNEL_VERSION` |
| `collectedAt` | string (ISO 8601) | yes | 収集開始日時。タイムゾーンオフセット付き |
| `computerName` | string | yes | `$env:COMPUTERNAME`（実 OS 上のコンピュータ名） |
| `hardwareUniqueId` | string | yes | `Get-HardwareUniqueId` 戻り値（BIOS Serial 由来） |
| `selectedNewPcName` | string | yes | `$env:SELECTED_NEW_PCNAME`（無ければ `computerName` と同値） |
| `workerName` | string \| null | no | `$env:FABRIQ_WORKER_NAME`（profile 実行外では null） |
| `sections` | array<Section> | yes | セクション結果配列 |
| `summary` | Summary | yes | 集計値 |

### Section オブジェクト

| Field | Type | Required | Description |
|---|---|---|---|
| `id` | string | yes | セクション ID。`"01"`〜`"22"`、`"8b"` 等の副 ID も許容 |
| `title` | string | yes | セクション名（例: `"System Basic Info"`） |
| `files` | string[] | yes | manifest 親ディレクトリからの相対パス配列。`/` で終わる文字列はディレクトリを意味する。書き込みファイルが無ければ空配列 |
| `status` | enum | yes | `"Success"` / `"Skipped"` / `"Failed"` / `"Partial"` のいずれか |
| `reason` | string \| null | yes | 非 Success 時の理由文字列。Success 時は `null` |
| `elapsedMs` | int | yes | セクション処理経過時間（ミリ秒） |

### Status セマンティクス

- **Success**: セクション完了、`files[]` のすべてが書き込まれた
- **Skipped**: 意図的にスキップ（例: Server 専用セクションを Client OS で実行 / Office 未インストール / Defender 未稼働）。`reason` 必須、`files` は通常空
- **Failed**: 例外発生でセクションが完了できなかった。`reason` に exception メッセージ。`files` は **常に空配列**（途中まで書かれた壊れたファイルを manifest に載せないため）
- **Partial**: 単一セクション内で複数の独立処理（例: §11 DesktopApps + StoreApps）の一部だけ成功した状態。`reason` に詳細

### Summary オブジェクト

| Field | Type | Required | Description |
|---|---|---|---|
| `sectionCount` | int | yes | `sections.length` |
| `successCount` | int | yes | status=Success の数 |
| `skippedCount` | int | yes | status=Skipped の数 |
| `failedCount` | int | yes | status=Failed の数 |
| `partialCount` | int | yes | status=Partial の数 |

**不変条件**: `successCount + skippedCount + failedCount + partialCount === sectionCount`

---

## 前方互換ルール

### 外部ツール（manager 等）の責任

1. **schemaVersion チェック必須**: 未知 major 版を検知したら警告を出し、legacy mode（manifest 無視 + ファイル列挙）にフォールバックする。silent な部分動作は禁止
2. **未知 section ID は raw 表示**: パーサが知らない `id` のセクションは raw text/CSV としてそのまま提示する。クラッシュさせない
3. **未知 status enum 値は Failed 扱い**: 将来 `"InProgress"` 等が追加されても安全側に倒す
4. **追加フィールドは無視**: schemaVersion=1 内での後方互換な field 追加は manager 側で無視可能

### evidence_config 側の責任

1. **schemaVersion を上げない限り破壊しない**: フィールドの削除・改名・型変更は schemaVersion=2 への昇格を伴う
2. **新 section 追加は schemaVersion=1 内で OK**: 既存 manager は未知 ID として raw 表示するので clash しない
3. **status enum 拡張は schemaVersion=2**: 既存 4 値以外を追加する場合のみ schemaVersion を上げる
4. **任意フィールドの追加は schemaVersion=1 内で OK**: required は変えない

---

## 再実行時の挙動

evidence_config を同じディレクトリで再実行する際:

1. 既存 `manifest.json` が存在すれば `manifest.json.bak` に rename（既存 `.bak` は削除して上書き）
2. 新しい収集を実行し、新 manifest を atomic に書き出し（収集完了時に一括）
3. 中断時の半端な manifest を防ぐため、incremental write は採用しない

---

## ディレクトリ表現

`files[]` の要素が `/` で終わる場合、それはディレクトリを意味する：

```json
"files": ["20_TempBackup.txt", "20_TempBackup/"]
```

manager は `/` で終わる要素を「opaque な forensic dump dir」として扱い、内部のファイルは個別パースしない（必要なら raw として一覧化のみ）。

---

## manifest 不在の旧 evidence

manifest.json 不在の旧形式 evidence（kernel 2.2.1 以前 / evidence_config 1.2.0 以前で収集されたもの）も外部ツールはサポートし続けることが期待される。manifest が無ければファイル列挙ベースで動作する従来挙動を維持する。

---

## 監査ギャップとの関連

`evidence_config` v1.2.0 は 22 セクションを出力するが、官公庁向け監査では追加項目が必要（project memory `project_evidence_audit_gaps` 参照）：

- **§23–§30 候補**: TPM / SecureBoot / gpresult / secedit / 証明書 / NTP / AccessControl

これらの追加は schemaVersion=1 内で OK（後方互換 section 追加）。manifest 契約は将来 audit gap を埋める追加に耐えられる設計になっている。

evidence_manager 側の対応では、§21（License）/§22（Office License）の重要 section 表示が manifest schema 経由で必須化されることで、過去の "files only" 表示モードでの section 取りこぼしを防ぐ（project memory `project_evidence_manifest_gap` 参照）。
