# Module category definitions loader.
# Reads apps/fabriq_ios/data/module_categories.json (schemaVersion 1)
# and exposes lookup helpers used by the GlobalConfig dispatcher
# (verb -> category), the CategoryConfig sub-mode (modules listed
# under a category), and the completer (filtered candidates).
#
# Phase 9b: each item in a category's `modules[]` array is either
#   - a string (e.g. "reg_hklm_config")    : implicit defaults
#       Name   = string itself
#       Dir    = string itself
#       Script = "<dir>.ps1"
#       Csv    = first *_list.csv in the directory, or (if none) the
#                single non-special *.csv (module.csv / preset.csv
#                excluded) - see Get-ModuleCsvSchema in module.ps1
#       Label  = string itself
#   - an object with explicit overrides:
#       { name, dir?, script?, csv?, label? }
# Resolve-ModuleEntry normalises both forms into a uniform hashtable.
#
# The exclusion list in lib/commands/module.ps1
# ($script:FabriqIosExcludedModules) takes precedence: a name listed
# in JSON but globally excluded is never returned.

$script:FabriqIosCategoryCache = $null

function Get-FabriqIosCategories {
    if ($null -ne $script:FabriqIosCategoryCache) {
        return $script:FabriqIosCategoryCache
    }
    $jsonPath = Join-Path $script:FabriqIosRoot 'data\module_categories.json'
    if (-not (Test-Path $jsonPath)) {
        Write-Host ("% module_categories.json not found at: {0}" -f $jsonPath) -ForegroundColor Red
        $script:FabriqIosCategoryCache = @{}
        return $script:FabriqIosCategoryCache
    }
    try {
        $raw = Get-Content -Path $jsonPath -Raw -Encoding UTF8
        $obj = $raw | ConvertFrom-Json
    } catch {
        Write-Host ("% Failed to parse module_categories.json: {0}" -f $_.Exception.Message) -ForegroundColor Red
        $script:FabriqIosCategoryCache = @{}
        return $script:FabriqIosCategoryCache
    }

    if ($obj.schemaVersion -ne 1) {
        Write-Host ("% Unsupported module_categories.json schemaVersion: {0}" -f $obj.schemaVersion) -ForegroundColor Red
        $script:FabriqIosCategoryCache = @{}
        return $script:FabriqIosCategoryCache
    }

    $byId = @{}
    foreach ($cat in $obj.categories) {
        $byId[$cat.id] = $cat
    }
    $script:FabriqIosCategoryCache = $byId
    return $byId
}

function Get-FabriqIosCategoryByVerb {
    param([string]$Verb)
    $cats = Get-FabriqIosCategories
    foreach ($cat in $cats.Values) {
        if ($cat.verb -eq $Verb) { return $cat }
    }
    return $null
}

function Get-FabriqIosCategoryById {
    param([string]$Id)
    $cats = Get-FabriqIosCategories
    if ($cats.ContainsKey($Id)) { return $cats[$Id] }
    return $null
}

function Get-FabriqIosEntryName {
    # Internal helper: extract the display name from either form.
    param($Item)
    if ($null -eq $Item) { return $null }
    if ($Item -is [string]) { return $Item }
    if ($Item.PSObject -and $Item.PSObject.Properties['name']) {
        return $Item.name
    }
    return $null
}

function Resolve-ModuleEntry {
    # Looks up (CategoryId, Name) and returns a fully-populated entry
    # hashtable, or $null if not found / globally excluded.
    param(
        [string]$CategoryId,
        [string]$Name
    )
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    if ($Name -in @($script:FabriqIosExcludedModules)) { return $null }

    $cat = Get-FabriqIosCategoryById -Id $CategoryId
    if (-not $cat) { return $null }

    foreach ($item in $cat.modules) {
        $itemName = Get-FabriqIosEntryName -Item $item
        if ($itemName -ne $Name) { continue }

        if ($item -is [string]) {
            return @{
                Name     = $item
                Dir      = $item
                Script   = ('{0}.ps1' -f $item)
                Csv      = $null   # auto-detect first *_list.csv
                Label    = $item
                Category = $cat.id
            }
        }
        # Object form
        $dir    = if ($item.PSObject.Properties['dir'])    { $item.dir    } else { $item.name }
        $script = if ($item.PSObject.Properties['script']) { $item.script } else { ('{0}.ps1' -f $dir) }
        $csv    = if ($item.PSObject.Properties['csv'])    { $item.csv    } else { $null }
        $label  = if ($item.PSObject.Properties['label'])  { $item.label  } else { $item.name }
        return @{
            Name     = $item.name
            Dir      = $dir
            Script   = $script
            Csv      = $csv
            Label    = $label
            Category = $cat.id
        }
    }
    return $null
}

function Find-ModuleEntryAcrossCategories {
    # Scans every category for a name match. Returns the resolved
    # entry (with Category field populated) or $null. Useful when
    # the dispatch path knows the name but not the category.
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    $cats = Get-FabriqIosCategories
    foreach ($cat in $cats.Values) {
        $entry = Resolve-ModuleEntry -CategoryId $cat.id -Name $Name
        if ($entry) { return $entry }
    }
    return $null
}

function Get-CategoryModuleCompletion {
    # Returns the sorted list of entry names belonging to the given
    # category, minus any in the global exclusion list.
    param([string]$CategoryId)
    $cat = Get-FabriqIosCategoryById -Id $CategoryId
    if (-not $cat) { return @() }
    $excluded = @($script:FabriqIosExcludedModules)
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($item in $cat.modules) {
        $itemName = Get-FabriqIosEntryName -Item $item
        if (-not $itemName) { continue }
        if ($itemName -in $excluded) { continue }
        $names.Add($itemName)
    }
    return @($names | Sort-Object)
}

function Test-ModuleInCategory {
    param(
        [string]$CategoryId,
        [string]$ModuleName
    )
    return ($null -ne (Resolve-ModuleEntry -CategoryId $CategoryId -Name $ModuleName))
}
