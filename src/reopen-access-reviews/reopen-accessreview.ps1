<#
Clone an existing Access Review schedule definition by ID and create a NEW one-time definition
that starts today (or StartDate) and stays open for 7 days (or InstanceDurationInDays).

Requires:
- Microsoft.Graph.Identity.Governance
- AccessReview.ReadWrite.All

Last Modified: 2026-01-28 12:28
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

function convertReviewerScopesToArray {
    param($items)
    if (-not $items) { return @() }

    $out = @()
    foreach ($i in @($items)) {
        $h = unwrapGraphObject $i
        if (-not $h) { continue }
        if ($h -isnot [System.Collections.IDictionary]) { continue }
        if (-not $h.ContainsKey('@odata.type')) { $h['@odata.type'] = '#microsoft.graph.accessReviewReviewerScope' }
        if ($h.ContainsKey('query') -and $h['query']) {
            if (-not $h.ContainsKey('queryType') -or -not $h['queryType']) { $h['queryType'] = 'MicrosoftGraph' }
        }
        # MUST have query
        if (-not $h.ContainsKey('query') -or -not $h['query']) { continue }
        $out += $h
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
        # OPTION A: Convert to simple accessReviewQueryScope (non-custom scope)
        # Extract the primary query from the old scope
        $oldScope = $old.Scope
        $primaryQuery = $null

        # Unwrap SDK object - actual scope data is in AdditionalProperties
        if ($oldScope.PSObject.Properties['AdditionalProperties'] -and $oldScope.AdditionalProperties) {
            $scopeData = $oldScope.AdditionalProperties
        } else {
            $scopeData = $oldScope
        }

        Write-Host "Scope data type: $($scopeData.GetType().Name)" -ForegroundColor Cyan
        Write-Host "Scope data: $(($scopeData | ConvertTo-Json -Depth 5))" -ForegroundColor Cyan

        # Helper function to safely get dictionary values
        function Get-DictValue {
            param($dict, $key)
            if ($dict -is [System.Collections.IDictionary]) {
                return $dict[$key]
            } else {
                return $dict.PSObject.Properties[$key].Value
            }
        }

        function Test-DictKey {
            param($dict, $key)
            if ($dict -is [System.Collections.IDictionary]) {
                return $dict.ContainsKey($key)
            } else {
                return $null -ne $dict.PSObject.Properties[$key]
            }
        }

        # Try to extract query from resourceScopes first
        if (Test-DictKey $scopeData 'resourceScopes') {
            $resourceScopes = Get-DictValue $scopeData 'resourceScopes'
            if ($resourceScopes -and $resourceScopes -is [System.Collections.IList] -and $resourceScopes.Count -gt 0) {
                $primaryQuery = Get-DictValue $resourceScopes[0] 'query'
            }
        }

        # If still no query, try direct query property
        if (-not $primaryQuery -and (Test-DictKey $scopeData 'query')) {
            $primaryQuery = Get-DictValue $scopeData 'query'
        }

        # If still no query, try principalScopes
        if (-not $primaryQuery -and (Test-DictKey $scopeData 'principalScopes')) {
            $principalScopes = Get-DictValue $scopeData 'principalScopes'
            if ($principalScopes -and $principalScopes -is [System.Collections.IList] -and $principalScopes.Count -gt 0) {
                $primaryQuery = Get-DictValue $principalScopes[0] 'query'
            }
        }

        $scopeDataKeys = if ($scopeData -is [System.Collections.IDictionary]) { $scopeData.Keys -join ', ' } else { $scopeData.PSObject.Properties.Name -join ', ' }
        if (-not $primaryQuery) {
            throw "Could not extract primary query from old definition scope. Scope keys: $scopeDataKeys"
        }

        # Build simple accessReviewQueryScope (Option A)
        # Note: Do NOT include queryType for basic scopes - it triggers custom scoping validation
        $scopeHt = @{
            '@odata.type' = '#microsoft.graph.accessReviewQueryScope'
            'query'       = normalizeGraphPathVersion $primaryQuery
        }

        $reviewersArr = @(convertReviewerScopesToArray $old.Reviewers)

        if ($reviewersArr.Count -eq 0) {
            throw "No valid reviewers found after conversion (Graph requires reviewers with query/queryType)."
        }

        # Settings: copy a safe subset, then override recurrence + duration
        $settings = $old.Settings
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
            if ($settings -and $settings.PSObject.Properties[$k]) { $settingsHt[$k] = $settings.$k }
        }

        # Build the body for the new definition
        $body = @{
            displayName             = $old.DisplayName + $NewDisplayNameSuffix
            descriptionForAdmins    = $old.DescriptionForAdmins
            descriptionForReviewers = $old.DescriptionForReviewers
            scope                   = $scopeHt
            reviewers               = $reviewersArr
            settings                = $settingsHt
            recurrence              = buildOneTimeRecurrence $StartDate
            instanceDurationInDays  = $InstanceDurationInDays
        }

        # Add fallbackReviewers if they exist
        $fallbackReviewersArr = @(convertReviewerScopesToArray $old.FallbackReviewers)
        if ($fallbackReviewersArr.Count -gt 0) {
            $body['fallbackReviewers'] = $fallbackReviewersArr
        }

        # Ensure required descriptions
        if (-not $body.descriptionForAdmins -or [string]::IsNullOrWhiteSpace($body.descriptionForAdmins)) {
            $body.descriptionForAdmins = "Cloned from $OldDefinitionId on $(Get-Date -Format 'yyyy-MM-dd')"
        }
        if (-not $body.descriptionForReviewers -or [string]::IsNullOrWhiteSpace($body.descriptionForReviewers)) {
            $body.descriptionForReviewers = "Please review your access."
        }

        removeNullsRecursively -obj $body

        Write-Host "`n--- Payload to Graph ---"
        $body | ConvertTo-Json -Depth 100 | Write-Host

        if ($IsWhatIf) {
            Write-Host "`n[WHATIF] Would create Access Review Definition with the payload shown above" -ForegroundColor Yellow
            return $null
        }

        try {
            $newDef = New-MgIdentityGovernanceAccessReviewDefinition -BodyParameter $body
            Write-Host "`nCreated Access Review Definition Id: $($newDef.Id)" -ForegroundColor Green
            return $newDef
        } catch {
            Write-Host "`nCreation failed." -ForegroundColor Red
            if ($_.ErrorDetails -and $_.ErrorDetails.Message) {
                Write-Host "`n--- ErrorDetails.Message ---" -ForegroundColor Red
                Write-Host $_.ErrorDetails.Message
            }
            Write-Host "`n--- Exception ---" -ForegroundColor Red
            Write-Host $_.Exception.Message
            throw
        }
    } catch {
        Write-Host "Failed to process definition $DefinitionId : $_" -ForegroundColor Red
        throw
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
