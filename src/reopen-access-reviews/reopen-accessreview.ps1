<#
Clone an existing Access Review schedule definition by ID and create a NEW one-time definition
that starts today (or StartDate) and stays open for 7 days (or InstanceDurationInDays).

Supports both simple accessReviewQueryScope and complex principalResourceMembershipsScope
with multiple principal and resource scopes (including B2B direct connect users and shared channels).

Requires:
- Microsoft.Graph.Identity.Governance
- AccessReview.ReadWrite.All

Last Modified: 2026-01-28 12:50
Fixed: Settings structure, scope handling, description defaults
#>

param(
    [Parameter(Mandatory = $false, Position = 0)]
    [string] $OldDefinitionId,

    [Parameter(Mandatory = $false)]
    [string] $FromFile,

    [Parameter(Mandatory = $false)]
    [datetime] $StartDate = (Get-Date),   # default: today

    [Parameter(Mandatory = $false)]
    [int] $InstanceDurationInDays = 7,    # default: one week open

    [Parameter(Mandatory = $false)]
    [string] $NewDisplayNameSuffix = " - Reopened (One-time)",

    [Parameter(Mandatory = $false)]
    [switch] $WhatIf,

    [Parameter(Mandatory = $false)]
    [switch] $DumpDefinition
)

$ErrorActionPreference = "Stop"

# Build list of IDs to process
$DefinitionIds = @()
if ($FromFile) {
    if (-not (Test-Path $FromFile)) {
        throw "File not found: $FromFile"
    }
    $DefinitionIds = Get-Content $FromFile | Where-Object { $_ -and $_.Trim() } | ForEach-Object { $_.Trim() }
    if ($DefinitionIds.Count -eq 0) {
        throw "No valid definition IDs found in file: $FromFile"
    }
    Write-Host "Loaded $($DefinitionIds.Count) definition ID(s) from file: $FromFile"
} elseif ($OldDefinitionId) {
    $DefinitionIds = @($OldDefinitionId)
} else {
    throw "-OldDefinitionId must be specified when -FromFile is not used."
}

# ---------------- Utilities ----------------
function toIso8601Duration {
    param($duration)
    if ($duration -is [string] -and $duration -match '^[Pp]') { return $duration }

    $ts = $null
    if ($duration -is [TimeSpan]) { $ts = $duration }
    elseif ($duration -and $duration.PSObject.Properties['Ticks']) { $ts = [TimeSpan]::FromTicks([int64]$duration.Ticks) }
    else { return $null }

    $d = [int]$ts.Days
    $h = [int]$ts.Hours
    $m = [int]$ts.Minutes
    $s = [int]$ts.Seconds

    $datePart = if ($d -gt 0) { "${d}D" } else { "" }
    $timeParts = @()
    if ($h -gt 0) { $timeParts += "${h}H" }
    if ($m -gt 0) { $timeParts += "${m}M" }
    if ($s -gt 0) { $timeParts += "${s}S" }

    if ($datePart -eq "" -and $timeParts.Count -eq 0) { return "PT0S" }
    if ($timeParts.Count -gt 0) { return "P$($datePart)T$($timeParts -join '')" }
    return "P$($datePart)"
}

function removeNullsRecursively {
    param([Parameter(Mandatory)]$obj)

    if ($obj -is [System.Collections.IDictionary]) {
        foreach ($k in @($obj.Keys)) {
            $v = $obj[$k]
            if ($null -eq $v) {
                $obj.Remove($k) | Out-Null
                continue
            }
            removeNullsRecursively -obj $v
        }
        return
    }

    if ($obj -is [System.Collections.IList]) {
        for ($i = 0; $i -lt $obj.Count; $i++) {
            $v = $obj[$i]
            if ($null -ne $v) { removeNullsRecursively -obj $v }
        }
        return
    }
}

# Replace only a leading /beta/ with /v1.0/
function normalizeGraphPathVersion {
    param([string]$path)
    if (-not $path) { return $path }
    if ($path -match '^/beta/') { return ($path -replace '^/beta/', '/v1.0/') }
    return $path
}

# ---------------- Deep coercion ----------------
function coerceGraphValue {
    param($v)

    if ($null -eq $v) { return $null }

    if ($v -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($k in $v.Keys) { $h["$k"] = coerceGraphValue $v[$k] }
        return $h
    }

    if (($v -is [System.Collections.IEnumerable]) -and -not ($v -is [string])) {
        if ($v -is [System.Collections.IDictionary]) { return (coerceGraphValue $v) }
        $arr = @()
        foreach ($x in $v) { $arr += (coerceGraphValue $x) }
        return $arr
    }

    return $v
}

function normalizeCommonGraphKeys {
    param($ht)
    if (-not ($ht -is [System.Collections.IDictionary])) { return $ht }

    if ($ht.ContainsKey('Query') -and -not $ht.ContainsKey('query')) {
        $ht['query'] = $ht['Query']
        $ht.Remove('Query') | Out-Null
    }
    if ($ht.ContainsKey('QueryType') -and -not $ht.ContainsKey('queryType')) {
        $ht['queryType'] = $ht['QueryType']
        $ht.Remove('QueryType') | Out-Null
    }
    if ($ht.ContainsKey('QueryRoot') -and -not $ht.ContainsKey('queryRoot')) {
        $ht['queryRoot'] = $ht['QueryRoot']
        $ht.Remove('QueryRoot') | Out-Null
    }
    return $ht
}

function normalizeGraphStructureRecursively {
    param($obj)

    if ($null -eq $obj) { return $null }

    if ($obj -is [System.Collections.IDictionary]) {
        $obj = normalizeCommonGraphKeys $obj
        foreach ($k in @($obj.Keys)) {
            $obj[$k] = normalizeGraphStructureRecursively $obj[$k]

            # Ensure certain properties are arrays
            if ($k -in @('principalScopes', 'resourceScopes', 'fallbackReviewers')) {
                if ($obj[$k] -isnot [System.Collections.IList] -and $null -ne $obj[$k]) {
                    $obj[$k] = @($obj[$k])
                }
            }
        }
        # Normalize query paths
        if ($obj.ContainsKey('query') -and $obj['query']) {
            $obj['query'] = normalizeGraphPathVersion $obj['query']
        }
        return $obj
    }

    if ($obj -is [System.Collections.IList]) {
        $arr = @()
        foreach ($item in $obj) {
            $arr += normalizeGraphStructureRecursively $item
        }
        return $arr
    }

    return $obj
}

# If it looks like a "dictionary dump object" (Keys/Values/SyncRoot/etc), use SyncRoot.
function unwrapDictDumpIfPresent {
    param($obj)

    if ($null -eq $obj) { return $null }

    if ($obj -is [System.Collections.IDictionary]) {
        if ($obj.ContainsKey('SyncRoot') -and $obj.ContainsKey('Keys') -and $obj.ContainsKey('Values')) {
            return $obj['SyncRoot']
        }
        return $obj
    }

    if ($obj.PSObject -and $obj.PSObject.Properties['SyncRoot'] -and $obj.PSObject.Properties['Keys'] -and $obj.PSObject.Properties['Values']) {
        return $obj.SyncRoot
    }

    return $obj
}

function unwrapGraphObject {
    param($sdkObj)

    if (-not $sdkObj) { return $null }

    # If already a dictionary/list, deep-coerce
    if ($sdkObj -is [System.Collections.IDictionary] -or (($sdkObj -is [System.Collections.IEnumerable]) -and -not ($sdkObj -is [string]))) {
        $coerced = coerceGraphValue $sdkObj
        if ($coerced -is [hashtable]) {
            $coerced = unwrapDictDumpIfPresent $coerced
            return normalizeGraphStructureRecursively $coerced
        }
        return normalizeGraphStructureRecursively $coerced
    }

    # Merge properties + AdditionalProperties
    $h = @{}
    $props = $sdkObj.PSObject.Properties | Where-Object { $_.Name -ne 'AdditionalProperties' }
    foreach ($p in $props) {
        $h[$p.Name] = coerceGraphValue $p.Value
    }
    if ($sdkObj.PSObject.Properties['AdditionalProperties']) {
        foreach ($k in $sdkObj.AdditionalProperties.Keys) {
            $h[$k] = coerceGraphValue $sdkObj.AdditionalProperties[$k]
        }
    }
    return normalizeGraphStructureRecursively $h
}

function normalizeQueryScopeItem {
    param($item)
    if (-not $item) { return $null }

    $unwrapped = unwrapGraphObject $item
    if ($unwrapped -isnot [System.Collections.IDictionary]) { return $null }

    # Build a clean scope with normalized keys
    $normalized = @{
        '@odata.type' = if ($unwrapped.'@odata.type') { $unwrapped.'@odata.type' } else { '#microsoft.graph.accessReviewQueryScope' }
    }

    # Add query if present
    if ($unwrapped.ContainsKey('query') -and $unwrapped['query']) {
        $normalized['query'] = $unwrapped['query']
    }

    # Add queryType if present
    if ($unwrapped.ContainsKey('queryType') -and $unwrapped['queryType']) {
        $normalized['queryType'] = $unwrapped['queryType']
    }

    # Add queryRoot if present (used in some reviewer scopes)
    if ($unwrapped.ContainsKey('queryRoot') -and $unwrapped['queryRoot']) {
        $normalized['queryRoot'] = $unwrapped['queryRoot']
    }

    return $normalized
}

function normalizeQueryScopeArray {
    param($items)
    if (-not $items) { return @() }

    $result = @()
    foreach ($item in @($items)) {
        $normalized = normalizeQueryScopeItem $item
        if ($normalized) {
            $result += $normalized
        }
    }
    return $result
}

function convertReviewerScopesToArray {
    param($items)
    if (-not $items) { return @() }

    $out = @()
    foreach ($i in @($items)) {
        $h = unwrapGraphObject $i
        if (-not $h) { continue }
        if ($h -isnot [System.Collections.IDictionary]) { continue }

        # PowerShell hashtables are case-insensitive, so Query and query are the same key
        # Check for query (case-insensitive check will find Query, query, etc.)
        if (-not $h.ContainsKey('query') -or -not $h['query']) { continue }

        # Build a NEW hashtable with lowercase keys to ensure JSON serialization is correct
        $normalized = @{
            '@odata.type' = '#microsoft.graph.accessReviewReviewerScope'
            'query'       = $h['query']
            'queryType'   = if ($h.ContainsKey('queryType') -and $h['queryType']) { $h['queryType'] } else { 'MicrosoftGraph' }
        }

        # Add queryRoot if it exists and is not null/empty
        if ($h.ContainsKey('queryRoot') -and $h['queryRoot']) {
            $normalized['queryRoot'] = $h['queryRoot']
        }

        $out += $normalized
    }
    return $out
}

function convertAdditionalRecipientsToArray {
    param($items)
    if (-not $items) { return @() }

    $result = @()
    foreach ($i in @($items)) {
        if (-not $i) { continue }

        $raw = unwrapGraphObject $i
        if (-not $raw) { continue }
        if ($raw -isnot [System.Collections.IDictionary]) { continue }

        $templateType = $raw['notificationTemplateType']
        if (-not $templateType) { $templateType = $raw['NotificationTemplateType'] }
        if (-not $templateType) { continue }

        $scope = $null
        foreach ($cand in @('notificationRecipientScope', 'NotificationRecipientScope', 'RecipientScope', 'Scope')) {
            if ($raw.ContainsKey($cand)) { $scope = $raw[$cand]; break }
        }

        $scopeH = normalizeQueryScopeItem $scope
        if (-not $scopeH -or ($scopeH -is [System.Collections.IDictionary] -and -not $scopeH.ContainsKey('query'))) { continue }

        $scopeH['@odata.type'] = '#microsoft.graph.accessReviewNotificationRecipientQueryScope'

        $result += @{
            '@odata.type'              = '#microsoft.graph.accessReviewNotificationRecipientItem'
            notificationTemplateType   = $templateType
            notificationRecipientScope = $scopeH
        }
    }

    return $result
}

# One-time recurrence starting StartDate with 1 occurrence
function buildOneTimeRecurrence {
    param([datetime]$start)

    $startStr = $start.ToString('yyyy-MM-dd')

    return @{
        pattern = @{
            type     = "weekly"
            interval = 1
        }
        range   = @{
            type                = "numbered"
            startDate           = $startStr
            numberOfOccurrences = 1
        }
    }
}

# Build settings object with all required fields
function buildSettingsObject {
    param(
        [object]$oldSettings,
        [datetime]$StartDate,
        [int]$InstanceDurationInDays
    )

    # Start with a safe subset of settings from the old definition
    $settingsHt = @{}

    foreach ($k in @(
            'mailNotificationsEnabled',
            'reminderNotificationsEnabled',
            'justificationRequiredOnApproval',
            'defaultDecisionEnabled',
            'defaultDecision',
            'recommendationsEnabled',
            'autoApplyDecisionsEnabled',
            'accessRecommendationsEnabled',
            'decisionHistoriesForReviewersEnabled'
        )) {
        if ($oldSettings -and $oldSettings.PSObject.Properties[$k]) {
            $settingsHt[$k] = $oldSettings.$k
        }
    }

    # Set the instance duration and recurrence at the settings level
    $settingsHt['instanceDurationInDays'] = $InstanceDurationInDays
    $settingsHt['recurrence'] = buildOneTimeRecurrence $StartDate

    return $settingsHt
}

# ---------------- Debug file generation ----------------
function Prompt-GenerateDebugFile {
    param(
        [string]$DefinitionId,
        [object]$Definition,
        [object]$Body,
        [string]$ErrorMessage,
        [string]$ErrorDetails
    )

    Write-Host "`n" -NoNewline
    Write-Host "Would you like to generate a debug file for the developer? [Y/n] (Auto-yes in 5 seconds): " -ForegroundColor Yellow -NoNewline

    $timeout = 5
    $startTime = Get-Date
    $response = $null

    while (((Get-Date) - $startTime).TotalSeconds -lt $timeout) {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            $response = $key.KeyChar.ToString().ToLower()
            break
        }
        Start-Sleep -Milliseconds 100
    }

    if ($null -eq $response -or $response -eq '' -or $response -eq 'y') {
        Write-Host "Yes"

        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $debugFileName = "debug-accessreview-$DefinitionId-$timestamp.json"

        $debugData = @{
            Timestamp    = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
            DefinitionId = $DefinitionId
            Definition   = $Definition
            RequestBody  = $Body
            Error        = @{
                Message = $ErrorMessage
                Details = $ErrorDetails
            }
        }

        try {
            $debugData | ConvertTo-Json -Depth 100 | Out-File -FilePath $debugFileName -Encoding UTF8
            Write-Host "Debug file created: $debugFileName" -ForegroundColor Green
        } catch {
            Write-Host "Failed to create debug file: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "No"
        Write-Host "Debug file generation skipped." -ForegroundColor Yellow
    }
}

# ---------------- Main processing function ----------------
function processAccessReviewDefinition {
    param(
        [string]$DefinitionId,
        [datetime]$StartDate,
        [int]$InstanceDurationInDays,
        [string]$NewDisplayNameSuffix,
        [bool]$IsWhatIf
    )

    Write-Host "`n========================================"
    Write-Host "Processing Definition ID: $DefinitionId"
    Write-Host "========================================`n"

    try {
        $old = Get-MgIdentityGovernanceAccessReviewDefinition -AccessReviewScheduleDefinitionId $DefinitionId

        # ---------------- Build new payload ----------------
        # Unwrap SDK object - actual scope data is in AdditionalProperties
        $oldScope = $old.Scope
        if ($oldScope.PSObject.Properties['AdditionalProperties'] -and $oldScope.AdditionalProperties) {
            $scopeData = $oldScope.AdditionalProperties
        } else {
            $scopeData = $oldScope
        }

        Write-Host "Scope data type: $($scopeData.GetType().Name)" -ForegroundColor Cyan
        Write-Host "Scope @odata.type: $(if ($scopeData.'@odata.type') { $scopeData.'@odata.type' } else { 'N/A' })" -ForegroundColor Cyan

        # Unwrap and properly structure the scope based on its type
        $scopeHt = unwrapGraphObject $scopeData

        # Ensure @odata.type is present (required for principalResourceMembershipsScope)
        if (-not $scopeHt.'@odata.type') {
            # If missing, try to determine from structure
            if ($scopeHt.ContainsKey('principalScopes') -or $scopeHt.ContainsKey('resourceScopes')) {
                $scopeHt.'@odata.type' = '#microsoft.graph.principalResourceMembershipsScope'
            } elseif ($scopeHt.ContainsKey('query')) {
                $scopeHt.'@odata.type' = '#microsoft.graph.accessReviewQueryScope'
            }
        }

        Write-Host "Using scope type: $($scopeHt.'@odata.type')" -ForegroundColor Green

        $reviewersArr = @(convertReviewerScopesToArray $old.Reviewers)

        if ($reviewersArr.Count -eq 0) {
            throw "No valid reviewers found after conversion (Graph requires reviewers with query/queryType)."
        }

        # Build settings with proper structure (includes recurrence and instanceDurationInDays)
        $settingsHt = buildSettingsObject -oldSettings $old.Settings -StartDate $StartDate -InstanceDurationInDays $InstanceDurationInDays

        # Build the body for the new definition
        # Based on MS documentation: recurrence and instanceDurationInDays go INSIDE settings
        $body = @{
            displayName             = $old.DisplayName + $NewDisplayNameSuffix
            descriptionForAdmins    = if ($old.DescriptionForAdmins) { $old.DescriptionForAdmins } else { "Cloned from $DefinitionId on $(Get-Date -Format 'yyyy-MM-dd')" }
            descriptionForReviewers = if ($old.DescriptionForReviewers) { $old.DescriptionForReviewers } else { "Please review your access." }
            scope                   = $scopeHt
            reviewers               = $reviewersArr
            settings                = $settingsHt
        }

        # Add fallbackReviewers if they exist
        $fallbackReviewersArr = @(convertReviewerScopesToArray $old.FallbackReviewers)
        if ($fallbackReviewersArr.Count -gt 0) {
            $body['fallbackReviewers'] = $fallbackReviewersArr
        }

        removeNullsRecursively -obj $body

        Write-Host "`n--- Payload to Graph ---"
        $body | ConvertTo-Json -Depth 100 | Write-Host

        if ($IsWhatIf) {
            Write-Host "`n[WHATIF] Would create Access Review Definition with the payload shown above" -ForegroundColor Yellow
            return $null
        }

        try {
            $newDef = New-MgIdentityGovernanceAccessReviewDefinition -BodyParameter $body -ErrorAction Stop
            Write-Host "`nCreated Access Review Definition Id: $($newDef.Id)" -ForegroundColor Green
            return $newDef
        } catch {
            Write-Host "`n========================================" -ForegroundColor Red
            Write-Host "ERROR: Creation Failed" -ForegroundColor Red
            Write-Host "========================================" -ForegroundColor Red

            $errorMsg = $_.Exception.Message
            $errorDetails = if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $_.ErrorDetails.Message } else { "" }

            Write-Host "`nError Message:" -ForegroundColor Red
            Write-Host $errorMsg

            if ($errorDetails) {
                Write-Host "`nError Details:" -ForegroundColor Red
                Write-Host $errorDetails
            }

            Write-Host "`nDefinition ID: $DefinitionId" -ForegroundColor Cyan

            # Prompt for debug file generation
            Prompt-GenerateDebugFile -DefinitionId $DefinitionId -Definition $old -Body $body -ErrorMessage $errorMsg -ErrorDetails $errorDetails

            Write-Host "`nSkipping this definition and continuing..." -ForegroundColor Yellow
            return $null
        }
    } catch {
        Write-Host "`n========================================" -ForegroundColor Red
        Write-Host "ERROR: Failed to Process Definition" -ForegroundColor Red
        Write-Host "========================================" -ForegroundColor Red
        Write-Host "`nDefinition ID: $DefinitionId" -ForegroundColor Cyan
        Write-Host "`nError: $_" -ForegroundColor Red

        # Prompt for debug file generation
        Prompt-GenerateDebugFile -DefinitionId $DefinitionId -Definition $null -Body $null -ErrorMessage $_.Exception.Message -ErrorDetails ""

        Write-Host "`nSkipping this definition and continuing..." -ForegroundColor Yellow
        return $null
    }
}

# ---------------- Connect & load ----------------
Import-Module Microsoft.Graph.Identity.Governance -ErrorAction Stop

# If DumpDefinition mode, connect and dump, then exit
if ($DumpDefinition.IsPresent) {
    if (-not (Get-MgContext)) {
        Connect-MgGraph -Scopes "AccessReview.Read.All"
    }

    foreach ($defId in $DefinitionIds) {
        Write-Host "`n========================================" -ForegroundColor Cyan
        Write-Host "Definition ID: $defId" -ForegroundColor Cyan
        Write-Host "========================================`n" -ForegroundColor Cyan

        try {
            $definition = Get-MgIdentityGovernanceAccessReviewDefinition -AccessReviewScheduleDefinitionId $defId

            # Fully convert the definition object to a hashtable with all nested objects expanded
            $convertedDef = @{
                Id                               = $definition.Id
                DisplayName                      = $definition.DisplayName
                DescriptionForAdmins             = $definition.DescriptionForAdmins
                DescriptionForReviewers          = $definition.DescriptionForReviewers
                CreatedDateTime                  = if ($definition.CreatedDateTime) { $definition.CreatedDateTime.ToString('o') } else { $null }
                LastModifiedDateTime             = if ($definition.LastModifiedDateTime) { $definition.LastModifiedDateTime.ToString('o') } else { $null }
                Status                           = $definition.Status
                InstanceDurationInDays           = $definition.InstanceDurationInDays
                Scope                            = unwrapGraphObject $definition.Scope
                Reviewers                        = @(convertReviewerScopesToArray $definition.Reviewers)
                FallbackReviewers                = @(convertReviewerScopesToArray $definition.FallbackReviewers)
                Settings                         = unwrapGraphObject $definition.Settings
                InstanceEnumerationScope         = unwrapGraphObject $definition.InstanceEnumerationScope
                Recurrence                       = unwrapGraphObject $definition.Recurrence
                AdditionalNotificationRecipients = @(convertAdditionalRecipientsToArray $definition.AdditionalNotificationRecipients)
            }

            # Remove null values
            removeNullsRecursively -obj $convertedDef

            # Convert to JSON with full depth for complete text representation
            $jsonOutput = $convertedDef | ConvertTo-Json -Depth 100
            Write-Host $jsonOutput
            Write-Host ""

        } catch {
            Write-Host "Failed to dump definition $defId : $_" -ForegroundColor Red
            Write-Host $_.Exception.Message -ForegroundColor Red
        }
    }

    Write-Host "`n========================================" -ForegroundColor Cyan
    Write-Host "Dump complete." -ForegroundColor Cyan
    Write-Host "========================================`n" -ForegroundColor Cyan
    exit 0
}

# Connect for normal processing
if (-not (Get-MgContext)) {
    Connect-MgGraph -Scopes "AccessReview.ReadWrite.All"
}

# Process each definition ID

$results = @()
foreach ($defId in $DefinitionIds) {
    $result = processAccessReviewDefinition `
        -DefinitionId $defId `
        -StartDate $StartDate `
        -InstanceDurationInDays $InstanceDurationInDays `
        -NewDisplayNameSuffix $NewDisplayNameSuffix `
        -IsWhatIf $WhatIf.IsPresent
    if ($result) {
        $results += $result
    }
}

if ($results.Count -gt 0) {
    Write-Host "`n========================================"
    Write-Host "Summary: Created $($results.Count) Access Review Definition(s)"
    Write-Host "========================================"
    $results | ForEach-Object { Write-Host "  - $($_.Id): $($_.DisplayName)" }
} elseif ($WhatIf.IsPresent) {
    Write-Host "`n[WHATIF] No definitions were created (WhatIf mode)" -ForegroundColor Yellow
}
