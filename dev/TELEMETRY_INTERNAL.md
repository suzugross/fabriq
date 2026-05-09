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
    ├── _kernel.jsonl                         # session-lifecycle events (profile/restart/finalize)
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
  "saltDigest": "sha256:a3f2c891",
  "host": {
    "os":       { "caption": "Microsoft Windows 11 Pro", "version": "10.0.26200", "build": "26200" },
    "hardware": { "manufacturer": "LENOVO", "model": "ThinkCentre M75q Gen 2", "ram_gb": 16 },
    "powershell": "5.1.26200.0"
  }
}
```

Written once per session by the first telemetry-relevant event (module envelope
or kernel event, whichever fires first). `saltDigest` is the first 4 bytes
(8 hex) of `sha256(salt)` — non-secret indicator that two sessions share a salt
(and thus their hashes correlate).

`host.hardware.manufacturer` / `host.hardware.model` are fleet-level identifiers
(e.g. ThinkCentre / Latitude / OptiPlex) and intentionally NOT redacted —
they enable cross-session correlation of "this PC model has X-class issues".
PC-individual identifiers (serial number, MAC, hostname) are still redacted
elsewhere and never appear in `_meta.json`.

## 6. JSONL events (one event per line)

### 6.1 `envelope.start`

```json
{
  "ts":"...","type":"envelope.start",
  "module":"hostname_config","sequence":1,"order":10,
  "segment":"","errorMode":"retry","group":"Network","isAsync":false,
  "profileName":"Master_Config01","profileOrder":10,"executionMode":"Linear",
  "prevModuleName":"windows_license_install","prevModuleStatus":"Success"
}
```

**Module-intrinsic fields** (always present):
- `sequence`: 1-based chronological index within the session
- `order` / `segment` / `errorMode` / `group`: passed by caller when known;
  empty/0 otherwise (e.g. single-module GUI execution)
- `isAsync`: `true` when running through `Invoke-SafeCommandAsync`

**Profile context fields** (present when fired from `Invoke-BatchExecution`,
i.e., a profile or batch run; absent for single-module GUI clicks):
- `profileName`: name of the profile being executed (Master_Config01 etc.)
- `profileOrder`: row Order in the profile CSV (often equal to top-level `order`)
- `executionMode`: `"Linear"` (Execute Profile) or `"Flex"` (FlexProfile)
- `prevModuleName` / `prevModuleStatus`: the previous module in the same batch
  and its outcome. Cross-module dependency analysis ("module X failed when
  prev was Skipped"). Absent for the first module of a batch.

### 6.2 `csv.load`

Emitted by `Import-ModuleCsv` after a successful CSV load. Captures only
**structural metadata** (no row values). Decision-trace value:
"module reg_hklm_config loaded reg_hklm_list.csv with 14 total / 1 enabled rows
under segment 'success_verified'".

```json
{
  "ts":"...","type":"csv.load",
  "fileName":"reg_hklm_list.csv",
  "path":"e:\\fabriq\\modules\\standard\\reg_hklm_config\\reg_hklm_list.csv",
  "totalRows":14,"returnedRows":1,
  "filterEnabled":true,"segment":"success_verified",
  "columns":["Enabled","TargetName","Description","Segment"]
}
```

### 6.3 `show.<function>`

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

### 6.4 `error`

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

### 6.5 `envelope.end`

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

## 6.6 Kernel events channel (`_kernel.jsonl`)

Session-lifecycle events are written to a **separate JSONL file** at the session
root (parallel to `modules/`). Same line format (one event per line) as module
envelopes but no envelope-scoped state. Emitted by `Write-KernelTelemetryEvent`
helper hooked into kernel control points.

```
logs/telemetry/{SessionID}/_kernel.jsonl
```

### 6.6.1 `profile.start`

Fired by `Invoke-BatchExecution` after the user confirms the batch.

```json
{
  "ts":"...","type":"profile.start",
  "profileName":"Master_Config01","profilePath":"profiles/Master_Config01.csv",
  "executionMode":"Linear","moduleCount":17,"autoPilot":true,
  "selectedOrders":[]
}
```

`selectedOrders` is populated only for FlexProfile group/batch runs.

### 6.6.2 `profile.end`

Fired at natural completion of `Invoke-BatchExecution` (after the foreach loop,
before `Complete-ProfileExecution`). NOT fired when `__RESTART__` early-exits
the loop — telemetry consumers should pair `profile.start` with the next
`resume.consumed` to span a restart cycle.

```json
{
  "ts":"...","type":"profile.end",
  "profileName":"Master_Config01","executionMode":"Linear",
  "modulesRun":17,"successCount":15,"errorCount":1,"skippedCount":1,
  "partialCount":0,"cancelledCount":0,"outcome":"WithErrors"
}
```

`outcome` enum: `"Success"` / `"WithErrors"` / `"Partial"`.

### 6.6.3 `restart.invoked`

Fired inside `Save-ResumeState` — the canonical "we are about to reboot"
boundary. Covers both `__RESTART__` markers (Profile-internal) and FlexProfile
[Restart Now] (`resumeAfterOrder=-1` sentinel).

```json
{
  "ts":"...","type":"restart.invoked",
  "profileName":"Master_Config01","executionMode":"Linear",
  "resumeAfterOrder":40,"completedCount":4,"schemaVersion":1
}
```

### 6.6.4 `resume.consumed`

Fired inside `Load-ResumeState` whenever `resume_state.json` is read into memory.
May fire multiple times per session (defensive callers re-read).

```json
{
  "ts":"...","type":"resume.consumed",
  "profileName":"Master_Config01","schemaVersion":1,
  "executionMode":"Linear","resumeAfterOrder":40,"completedCount":4
}
```

### 6.6.5 `finalize.start` / `finalize.end`

Wraps `Complete-ProfileExecution` (HTML checklist + log_uploader pipeline).
`mode` echoes the function's `Auto` / `Manual` flag (Auto = Linear auto-finalize,
Manual = `[cl]` regenerate or Flex `[Complete]`).

```json
{"ts":"...","type":"finalize.start","profileName":"...","mode":"Auto","elapsedMs":482300,"definedModules":17}
{"ts":"...","type":"finalize.end","profileName":"...","mode":"Auto","durationMs":3120,"checklistGenerated":true}
```

## 7. Implementation surface

| File | Change |
|---|---|
| `kernel/common.ps1` | `Telemetry Layer` section: `Get-TelemetrySalt`, `New-TelemetryRedactMap`, `Invoke-TelemetryRedact`, `Write-TelemetryEvent`, `Start-ModuleTelemetry`, `Complete-ModuleTelemetry`, `_GetShowTag`, `_TrackShowEvent`, `Get-TelemetryHostInfo`, `_WriteTelemetryMeta`, `Write-KernelTelemetryEvent` |
| `kernel/common.ps1` | Show-* family: 1 line added per fn (call to `_TrackShowEvent`) |
| `kernel/common.ps1` | `Invoke-SafeCommand`: wrap with `Start-ModuleTelemetry` + `Complete-ModuleTelemetry` |
| `kernel/common.ps1` | `Invoke-SafeCommandAsync`: same + inject envelope into runspace |
| `kernel/common.ps1` | `Import-ModuleCsv`: emit `csv.load` event before successful return |
| `kernel/common.ps1` | `Save-ResumeState` / `Load-ResumeState`: emit `restart.invoked` / `resume.consumed` |
| `kernel/common.ps1` | `Complete-ProfileExecution`: wrap with `finalize.start` / `finalize.end` |
| `kernel/main.ps1` | `Invoke-BatchExecution`: set `$global:_FabriqCurrentProfileContext` per module + emit `profile.start` / `profile.end` |
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
