# Fabriq Internal Telemetry — Design Notes

**Status**: Internal implementation. **Not** part of `KERNEL_API.md` public surface (yet).
**Schema Version**: `1` (field `telemetrySchemaVersion` in every emitted file)
**Introduced**: kernel `3.2.3` (PATCH bump pending — released on user instruction)

---

## 1. Purpose

Capture per-module + per-event behavior of fabriq kitting runs as a **machine-readable
corpus** for AI agents working on fabriq development. Not for human consumption.
Not for customer-facing audit (evidence/HTML checklist remain the audit channel).

Goals:
- Cross-session pattern detection ("module X fails 12 times with HRESULT 0x80072030")
- Decision-branch tracing (which CSV row triggered which Skip path)
- Full `ErrorRecord` capture (catch-and-swallowed exceptions included)
- Step-level timing
- Idempotency-violation detection via before/after intent

Non-goals:
- Real-time observability dashboard
- External consumer ingest (no public contract; AI reads JSONL directly)
- Human readability (verbose / structured / no formatting effort)

## 2. Why internal (not `KERNEL_API.md` §11)

User explicitly scoped this as developer-internal. Public-contract status is reserved
for things external tools `consume` (overlay rules, evidence manifest). Promoting
this layer to §11 is a future MINOR change once a stable consumer exists
(e.g. fabriq_studio diagnostic viewer).

Implication: fields can be added/renamed/removed within `schemaVersion=1`
without coordinating external tools. Schema bump only signals to the AI
corpus reader.

## 3. Privacy contract

**Inputs treated as sensitive** (auto-detected from `Set-SelectedHostEnvironment`
output and `$global:FabriqUniqueId`):

| Source | Treatment | Token |
|---|---|---|
| `$env:SELECTED_PIN` | Hard redact | `[REDACTED]` |
| `$env:SELECTED_KANRI_NO` | Salted SHA-256 (12 hex) | `KANRI:<hash>` |
| `$env:SELECTED_OLD_PCNAME` / `NEW_PCNAME` | Same | `PC:<hash>` |
| `$env:SELECTED_ETH_*` / `WIFI_*` / `DNS1-4` | Same | `IP:<hash>` |
| `$env:SELECTED_PRINTER_<N>_NAME/DRIVER/PORT` (N=1..10) | Same | `PRINTER:<hash>` |
| `$env:FABRIQ_WORKER_NAME` | Same | `WORKER:<hash>` |
| `$env:COMPUTERNAME` | Same | `HOST:<hash>` |
| `$global:FabriqUniqueId` | Same | `HW:<hash>` |

**Salt** is generated on first run and persisted at
`kernel/json/telemetry_salt.txt` (256 bits, base64). Same environment ⇒ same
hashes (cross-session correlation works). Different environments ⇒ no
correlation possible. Salt file is `.gitignore`d.

**Application**: at module envelope start, a redact map is built from current
env vars. Every Show-* message + every error message + every script stack
trace is post-processed by replacing literal occurrences of any sensitive
value with its token. Replacements applied longest-first to avoid partial
overlap.

**Limitations**:
- Direct `Write-Host` calls in modules are not captured (covers ~80-90%
  of fabriq log surface; the rest is `Write-Host` for layout / decoration).
- Child-process stdout (winget/robocopy/dism) not captured.
- Module-internal variables not in env are not redacted (e.g. an
  intermediate computed string that contains an IP). Acceptable for v1
  given the redact map covers the canonical sources.

## 4. File layout

```
logs/telemetry/
└── {SessionID}/                              # e.g. 20260509_143000
    ├── _meta.json                            # session-level metadata (written once)
    └── modules/
        ├── 0001_hostname_config.jsonl        # NN = sequence (1-based, zero-padded 4)
        ├── 0002_ipaddress_config.jsonl
        ├── ...
```

The leading number is a **chronological sequence**, not the Profile `Order`.
Profile `Order` is recorded inside the JSONL `envelope.start` event (when
known by the caller). Sequence keeps filename ordering monotonic even in
single-module reruns / out-of-order Flex execution.

`logs/telemetry/` is excluded from `log_uploader` (so it never leaves the
kitting PC) and excluded from `.gitignore` (whole `logs/*` already covered).

## 5. `_meta.json` schema

```json
{
  "telemetrySchemaVersion": 1,
  "sessionId": "20260509_143000",
  "kernelVersion": "3.2.3",
  "startedAt": "2026-05-09T14:30:00.123+09:00",
  "redactionPolicy": "hash-and-redact-v1",
  "saltDigest": "sha256:a3f2c891"
}
```

Written once per session by the first envelope of that session.
`saltDigest` is the first 4 bytes (8 hex) of `sha256(salt)` — non-secret
indicator that two sessions share a salt (and thus their hashes correlate).

## 6. JSONL events (one event per line)

### 6.1 `envelope.start`

```json
{"ts":"...","type":"envelope.start","module":"hostname_config","sequence":1,"order":10,"segment":"","errorMode":"","group":"","isAsync":false}
```

- `sequence`: 1-based chronological index within the session
- `order` / `segment` / `errorMode` / `group`: passed by caller when known;
  empty/0 otherwise (e.g. single-module GUI execution)
- `isAsync`: `true` when running through `Invoke-SafeCommandAsync`

### 6.2 `show.<function>`

One event per `Show-Info` / `Show-Success` / `Show-Warning` / `Show-Error`
/ `Show-Skip` call inside the module envelope.

```json
{"ts":"...","type":"show.info","tag":"info","msg":"Loaded reg_hklm_list.csv (15 items)"}
{"ts":"...","type":"show.warning","tag":"verifyFail","msg":"[VERIFY FAILED] Expected: PC:def456, Pending: PC:abc123"}
```

`tag` is inferred from message prefix:
| Prefix | tag |
|---|---|
| `[APPLY]` | `apply` |
| `[SKIP]` / `[NOT FOUND]` | `skip` / `notFound` |
| `[VERIFIED]` | `verifyPass` |
| `[VERIFY FAILED]` | `verifyFail` |
| `[AUTOPILOT]` | `autopilot` |
| `[ASYNC]` | `async` |
| `[RESTART]` | `restart` |
| (none of the above) | function name (`info` / `success` / etc.) |

### 6.3 `error`

Emitted at `envelope.end` time, **one event per `$Error` entry** captured
during the envelope (catch-and-swallowed entries included — `$Error.Clear()`
runs at envelope start to scope the capture).

```json
{
  "ts":"...","type":"error",
  "errorType":"System.UnauthorizedAccessException",
  "hresult":-2147024891,
  "msg":"Access to the path is denied",
  "scriptStack":"at <ScriptBlock>, e:\\fabriq\\modules\\standard\\reg_hklm_config\\reg_hklm_config.ps1: line 142",
  "categoryInfo":"PermissionDenied: (HKLM:\\...:String) [Set-ItemProperty], UnauthorizedAccessException",
  "targetObject":"HKLM:\\..."
}
```

`scriptStack` is preserved as-is (fabriq paths are not sensitive). Message
content is run through the redact map.

### 6.4 `envelope.end`

```json
{
  "ts":"...","type":"envelope.end",
  "status":"Partial",
  "verified":false,
  "message":"Success: 13, Skip: 1, Fail: 1",
  "duration_ms":1234,
  "errorCount":1,
  "showCounts":{"info":3,"success":12,"warning":1,"error":0,"skip":2}
}
```

`verified` is `null` when the module did not implement Post-Apply
Verification.

## 7. Implementation surface

| File | Change |
|---|---|
| `kernel/common.ps1` | `Telemetry Layer` section (~250 lines): `Get-TelemetrySalt`, `New-TelemetryRedactMap`, `Invoke-TelemetryRedact`, `Write-TelemetryEvent`, `Start-ModuleTelemetry`, `Complete-ModuleTelemetry`, `_Get-ShowTag`, `_Track-ShowEvent` |
| `kernel/common.ps1` | Show-* family: 1 line added (call to `_Track-ShowEvent`) |
| `kernel/common.ps1` | `Invoke-SafeCommand`: wrap `& $ScriptBlock` with `Start-ModuleTelemetry` + `Complete-ModuleTelemetry` |
| `kernel/common.ps1` | `Invoke-SafeCommandAsync`: same + inject envelope into runspace |
| `modules/extended/log_uploader/log_uploader.ps1` | robocopy `logs/` invocation gets `/XD` switch for `logs\telemetry` |
| `modules/extended/log_uploader/VERSION` | `1.0.0` → `1.1.0` (MINOR: new exclusion behavior) |
| `.gitignore` | Add `kernel/json/telemetry_salt.txt` (explicit; `/logs/*` already covers `logs/telemetry/`) |

**Module side: 0 changes** (all 74 modules pick up telemetry transparently via
Show-* and Invoke-SafeCommand wrapping).

## 8. Reentrancy / failure isolation

- `$script:_TelemetryWriting` flag: set during JSONL append. Show-* and
  the writer itself check it to avoid recursion (Show-Error called from
  inside a telemetry write would otherwise loop).
- Every telemetry path is wrapped in `try { } catch { }` with **no**
  fallback message to console. **Telemetry never affects kitting outcomes
  or surfaces errors to the operator.**
- File creation is best-effort: directory creation failure ⇒ envelope
  silently disabled for the run, kitting continues normally.
- Salt generation failure ⇒ all values get `[REDACTED]` (safe-by-default).

## 9. Future hooks (not implemented in v1)

- Step-level markers inside module scripts (`Mark-TelemetryStep -Name "Step5"`).
  Modules would opt-in. Adds richer timing.
- Cross-session aggregation index (`logs/telemetry/index.jsonl`).
  Skipped per user request — manual aggregation handled outside fabriq.
- Promote to public `KERNEL_API.md §11` if external consumer materializes.
