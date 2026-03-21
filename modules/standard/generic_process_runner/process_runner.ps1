# ========================================
# Generic Process Runner Script
# ========================================
# 任意の EXE ファイルを CSV 定義に基づいて実行する汎用モジュール。
#
# [NOTES]
# - ExecutablePath は絶対パス、環境変数、WorkingDirectory からの相対パスに対応
# - WorkingDirectory 未指定時は $PSScriptRoot を基準にパスを解決する
# ========================================

Write-Host ""
Show-Separator
Write-Host "Generic Process Runner" -ForegroundColor Cyan
Show-Separator
Write-Host ""


# ========================================
# Step 1: CSV 読み込み
# ========================================
$csvPath = Join-Path $PSScriptRoot "process_list.csv"

$processList = Import-ModuleCsv -Path $csvPath -FilterEnabled `
    -RequiredColumns @("Enabled", "Description", "ExecutablePath", "Arguments", "WorkingDirectory", "TimeoutSec", "SuccessCodes", "NoNewWindow")

if ($null -eq $processList) {
    return (New-ModuleResult -Status "Error" -Message "Failed to load process_list.csv")
}
if ($processList.Count -eq 0) {
    return (New-ModuleResult -Status "Skipped" -Message "No enabled process entries")
}
Write-Host ""


# ========================================
# Step 3: 実行前の確認表示（ドライラン）
# ========================================
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host "Process Execution List" -ForegroundColor Cyan
Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

$missingCount = 0

foreach ($proc in $processList) {
    # --- パス解決 ---
    # 環境変数を展開
    $exePath = [Environment]::ExpandEnvironmentVariables($proc.ExecutablePath)
    $workDir = if (-not [string]::IsNullOrWhiteSpace($proc.WorkingDirectory)) {
        [Environment]::ExpandEnvironmentVariables($proc.WorkingDirectory)
    } else {
        ""
    }

    # 絶対パス / 相対パスの解決
    if ([System.IO.Path]::IsPathRooted($exePath)) {
        $fullPath = $exePath
    } elseif (-not [string]::IsNullOrWhiteSpace($workDir)) {
        $fullPath = Join-Path $workDir $exePath
    } else {
        $fullPath = Join-Path $PSScriptRoot $exePath
    }

    # 表示用パラメータの整形
    $exists  = Test-Path $fullPath
    $timeout = if ([string]::IsNullOrWhiteSpace($proc.TimeoutSec) -or $proc.TimeoutSec -eq '0') { "None" } else { "$($proc.TimeoutSec)s" }
    $codes   = if ([string]::IsNullOrWhiteSpace($proc.SuccessCodes)) { "0" } else { $proc.SuccessCodes }
    $window  = if ($proc.NoNewWindow -eq '1') { "NoNewWindow" } else { "NewWindow" }
    $dispWorkDir = if (-not [string]::IsNullOrWhiteSpace($workDir)) { $workDir } else { "(default: $PSScriptRoot)" }

    if ($exists) {
        Write-Host "  $($proc.Description)" -ForegroundColor Yellow
        Write-Host "    EXE:      $fullPath"
        Write-Host "    WorkDir:  $dispWorkDir"
        Write-Host "    Timeout:  $timeout / SuccessCodes: $codes / Window: $window"
        if (-not [string]::IsNullOrWhiteSpace($proc.Arguments)) {
            Write-Host "    Args:     $($proc.Arguments)"
        }
        if (-not [string]::IsNullOrWhiteSpace($proc.WaitProcessName)) {
            Write-Host "    WaitFor:  $($proc.WaitProcessName) (poll until process exits)"
        }
    }
    else {
        Write-Host "  $($proc.Description) [NOT FOUND]" -ForegroundColor Red
        Write-Host "    EXE:      $fullPath"
        $missingCount++
    }
    Write-Host ""
}

Write-Host "----------------------------------------" -ForegroundColor White
Write-Host ""

if ($missingCount -gt 0) {
    Show-Warning "$missingCount executable(s) not found"
    Show-Info "Missing executables will be skipped during execution"
    Write-Host ""
}


# ========================================
# Step 4: 実行確認
# ========================================
$cancelResult = Confirm-ModuleExecution -Message "Proceed with process execution?"
if ($null -ne $cancelResult) { return $cancelResult }

Write-Host ""


# ========================================
# Step 5: プロセス実行ループ
# ========================================
$successCount = 0
$skipCount    = 0
$failCount    = 0

foreach ($proc in $processList) {
    $desc = $proc.Description

    # --- パス解決（Step 3 と同一ロジック） ---
    $exePath = [Environment]::ExpandEnvironmentVariables($proc.ExecutablePath)
    $workDir = if (-not [string]::IsNullOrWhiteSpace($proc.WorkingDirectory)) {
        [Environment]::ExpandEnvironmentVariables($proc.WorkingDirectory)
    } else {
        ""
    }

    if ([System.IO.Path]::IsPathRooted($exePath)) {
        $fullPath = $exePath
    } elseif (-not [string]::IsNullOrWhiteSpace($workDir)) {
        $fullPath = Join-Path $workDir $exePath
    } else {
        $fullPath = Join-Path $PSScriptRoot $exePath
    }

    Write-Host "----------------------------------------" -ForegroundColor White
    Write-Host "Executing: $desc" -ForegroundColor Cyan
    Write-Host "----------------------------------------" -ForegroundColor White

    # --- 存在チェック ---
    if (-not (Test-Path $fullPath)) {
        Show-Skip "Executable not found: $fullPath"
        Write-Host ""
        $skipCount++
        continue
    }

    # --- パラメータ整形 ---
    $successCodes = if ([string]::IsNullOrWhiteSpace($proc.SuccessCodes)) { "0" } else { $proc.SuccessCodes }
    $timeoutSec   = if ([string]::IsNullOrWhiteSpace($proc.TimeoutSec) -or $proc.TimeoutSec -eq '0') { 0 } else { [int]$proc.TimeoutSec }
    $successList  = @($successCodes -split ',' | ForEach-Object { [int]$_.Trim() })

    # --- 作業ディレクトリの決定 ---
    $execWorkDir = if (-not [string]::IsNullOrWhiteSpace($workDir)) {
        $workDir
    } else {
        $PSScriptRoot
    }

    # --- Start-Process 引数の組み立て ---
    $spParams = @{
        FilePath         = $fullPath
        WorkingDirectory = $execWorkDir
        PassThru         = $true
    }

    if (-not [string]::IsNullOrWhiteSpace($proc.Arguments)) {
        $spParams.ArgumentList = $proc.Arguments
    }

    if ($proc.NoNewWindow -eq '1') {
        $spParams.NoNewWindow = $true
    }

    try {
        $process = Start-Process @spParams

        Show-Info "Process started (PID: $($process.Id))"

        # --- タイムアウト監視 ---
        $timedOut = $false
        if ($timeoutSec -gt 0) {
            $null = $process | Wait-Process -Timeout $timeoutSec -ErrorAction SilentlyContinue
            if (-not $process.HasExited) {
                $timedOut = $true
                Show-Error "Timeout ($($timeoutSec)s exceeded). Killing process..."
                $process | Stop-Process -Force -ErrorAction SilentlyContinue
            }
        } else {
            $process | Wait-Process
        }

        # --- Post-execution process polling (WaitProcessName) ---
        $waitProcessName = if (
            ($proc.PSObject.Properties.Name -contains 'WaitProcessName') -and
            (-not [string]::IsNullOrWhiteSpace($proc.WaitProcessName))
        ) { $proc.WaitProcessName } else { "" }

        if (-not [string]::IsNullOrWhiteSpace($waitProcessName) -and -not $timedOut) {
            Show-Info "Waiting for process '$waitProcessName' to complete..."
            $pollInterval = 5
            $elapsed = 0
            while ($true) {
                $running = Get-Process -Name $waitProcessName -ErrorAction SilentlyContinue
                if (-not $running) { break }
                if ($timeoutSec -gt 0) {
                    $elapsed += $pollInterval
                    if ($elapsed -ge $timeoutSec) {
                        $timedOut = $true
                        Show-Error "Timeout ($($timeoutSec)s exceeded) while waiting for '$waitProcessName'"
                        $running | Stop-Process -Force -ErrorAction SilentlyContinue
                        break
                    }
                }
                Start-Sleep -Seconds $pollInterval
            }
            if (-not $timedOut) {
                Show-Success "Process '$waitProcessName' has completed"
            }
        }

        $exitCode = $process.ExitCode

        # --- 結果判定 ---
        if ($timedOut) {
            Show-Error "$desc : Timed out after $($timeoutSec)s"
            $failCount++
        }
        elseif ($exitCode -in $successList) {
            Show-Success "$desc (ExitCode: $exitCode)"
            $successCount++
        }
        else {
            Show-Error "$desc (ExitCode: $exitCode, Expected: $successCodes)"
            $failCount++
        }
    }
    catch {
        Show-Error "$desc : $($_.Exception.Message)"
        $failCount++
    }

    Write-Host ""
}


# ========================================
# Step 6: 結果集計・返却
# ========================================
return (New-BatchResult -Success $successCount -Skip $skipCount -Fail $failCount `
    -Title "Process Runner Results")
