# Fabriq IOS

A Cisco IOS-style command shell layered over the Fabriq kitting framework.
This is an art object, not a production tool. See `SPEC.md` for the design
intent and `CLAUDE_CODE_PROMPT.md` for the original brief.

## Status

Phase 1 skeleton. Directory layout, function signatures, and data files are
in place. The self-spawning subprocess guard works end-to-end. Function
bodies are stubs (`throw 'Not implemented: ...'`) and tests are `-Skip`
placeholders. Implementation lands in subsequent phases.

## Launch paths

- **From Fabriq Operator**: open the **Settings** tab, click **[FabriqApps]**,
  select `fabriq_ios` from the list, click **Launch**. The app appears
  automatically because `apps_dialog.ps1` discovers any `apps/<name>/<name>.ps1`.
- **Direct**: invoke `apps\fabriq_ios\fabriq_ios.ps1` from any shell.

In both cases, `fabriq_ios.ps1` self-spawns an isolated `powershell.exe`
subprocess (`-NoProfile -ExecutionPolicy Bypass -File ...`) so that
PSReadLine key handlers, environment-variable mutations, and global-scope
state never leak back into the caller. The subprocess is detected via
`$env:FABRIQ_IOS_SUBPROCESS`.

## Internal API coupling

`fabriq_ios` (planned, mutation phase) depends on the following kernel
internal-implementation symbols, which are listed under KERNEL_API.md
section 6 ("Internal implementation, may change in PATCH releases"):

- `Invoke-BatchExecution`
- `Initialize-ModuleSystem`
- `Resolve-ProfileModules`
- `Test-MasterPassphrase`

This coupling is accepted as the price of the joke. Re-validate this app
after any kernel PATCH bump. The shell stays inside `apps/fabriq_ios/` and
does not modify the kernel surface.

## Tests

Tests use Pester 5. Install with:

    Install-Module Pester -Force -SkipPublisherCheck -Scope CurrentUser

Run with:

    Invoke-Pester apps\fabriq_ios\tests\
