# ========================================
# Fabriq Test Helpers - CSV / Mock Modules
# ========================================
# Builds temporary Profile CSV files and mock module objects for kernel
# unit tests. ASCII encoding is used to match Resolve-ProfileModules'
# Import-Csv -Encoding Default behavior on Japanese-locale Windows
# (CP932), where a UTF-8 BOM would be misread as garbage in the header.
# ========================================

function New-TestProfileCsv {
    param(
        [Parameter(Mandatory)][object[]]$Rows,
        [string[]]$Columns = @('Order','ScriptPath','Enabled','Description','Segment','ErrorMode','Group')
    )
    $csvPath = Join-Path $env:TEMP ("fabriq-test-{0}.csv" -f ([guid]::NewGuid().ToString('N')))

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add($Columns -join ',')

    foreach ($row in $Rows) {
        $values = foreach ($col in $Columns) {
            $val = if ($row -is [hashtable]) {
                if ($row.ContainsKey($col)) { $row[$col] } else { '' }
            } else {
                if ($row.PSObject.Properties.Name -contains $col) { $row.$col } else { '' }
            }
            $valStr = "$val"
            if ($valStr -match '[,"]') {
                '"' + ($valStr -replace '"','""') + '"'
            } else {
                $valStr
            }
        }
        $lines.Add($values -join ',')
    }

    [System.IO.File]::WriteAllLines($csvPath, $lines, [System.Text.Encoding]::ASCII)
    return $csvPath
}

function Remove-TestProfileCsv {
    param([string]$Path)
    if ($Path -and (Test-Path $Path)) {
        Remove-Item $Path -Force -ErrorAction SilentlyContinue
    }
}

function New-MockModule {
    param(
        [Parameter(Mandatory)][string]$RelativePath,
        [Parameter(Mandatory)][string]$MenuName,
        [string]$ModuleDir,
        [string]$Category = 'Test'
    )
    if (-not $ModuleDir) {
        $parts = $RelativePath -split '[\\/]'
        $ModuleDir = if ($parts.Count -ge 2) { $parts[1] } else { '' }
    }
    [PSCustomObject]@{
        RelativePath = $RelativePath
        MenuName     = $MenuName
        ModuleDir    = $ModuleDir
        Category     = $Category
        Script       = (Split-Path -Leaf $RelativePath)
    }
}
