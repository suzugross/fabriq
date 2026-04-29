# Module category definitions loader.
# Reads apps/fabriq_ios/data/module_categories.json (schemaVersion 1)
# and exposes lookup helpers used by the GlobalConfig dispatcher
# (verb -> category), the CategoryConfig sub-mode (modules listed
# under a category), and the completer (filtered candidates).
#
# The exclusion list in lib/commands/module.ps1
# ($script:FabriqIosExcludedModules) takes precedence: a module
# present in JSON but excluded globally is never returned.

$script:FabriqIosCategoryCache = $null

function Get-FabriqIosCategories {
    # Returns a hashtable mapping category ID to the JSON object.
    # Cached for the session.
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
    # Returns the category object whose verb matches, or $null.
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

function Get-CategoryModuleCompletion {
    # Returns the list of module names belonging to the given
    # category, minus any in the global exclusion list. Sorted.
    param([string]$CategoryId)
    $cat = Get-FabriqIosCategoryById -Id $CategoryId
    if (-not $cat) { return @() }
    $excluded = @($script:FabriqIosExcludedModules)
    $names = @($cat.modules | Where-Object { $_ -and ($_ -notin $excluded) })
    return @($names | Sort-Object)
}

function Test-ModuleInCategory {
    # Returns $true if the module is listed under the category and
    # not globally excluded.
    param(
        [string]$CategoryId,
        [string]$ModuleName
    )
    if ([string]::IsNullOrWhiteSpace($ModuleName)) { return $false }
    if ($ModuleName -in @($script:FabriqIosExcludedModules)) { return $false }
    $cat = Get-FabriqIosCategoryById -Id $CategoryId
    if (-not $cat) { return $false }
    return ($ModuleName -in @($cat.modules))
}
