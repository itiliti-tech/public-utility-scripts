<#
Clone an existing Access Review schedule definition by ID and create a NEW one-time definition
that starts today (or StartDate) and stays open for 30 days (or InstanceDurationInDays).

Automatically simplifies complex principalResourceMembershipsScope to basic accessReviewQueryScope
for maximum tenant compatibility. Supports both simple and complex source definitions.

Requires:
- Microsoft.Graph.Identity.Governance
- AccessReview.ReadWrite.All

Last Modified: 2026-01-28 16:18
Fixed: Settings structure, scope handling, description defaults, tenant compatibility
#>

param(
    [Parameter(Mandatory = $false, Position = 0)]
    [string] $OldDefinitionId,

    [Parameter(Mandatory = $false)]
    [string] $FromFile,

    [Parameter(Mandatory = $false)]
    [datetime] $StartDate = (Get-Date),   # default: today

    [Parameter(Mandatory = $false)]
    [int] $InstanceDurationInDays = 30,    # default: 30 days open

    [Parameter(Mandatory = $false)]
    [string] $NewDisplayNameSuffix = " - Reopened (One-time)",

    [Parameter(Mandatory = $false)]
    [switch] $WhatIf,

    [Parameter(Mandatory = $false)]
    [switch] $DumpDefinition,

    [Parameter(Mandatory = $false)]
    [switch] $SuppressOutputLogs,

    [Parameter(Mandatory = $false)]
    [switch] $DisplayBody
)

$ErrorActionPreference = "Stop"

# ---------------- Script Metadata ----------------
$ScriptName = Split-Path -Leaf $PSCommandPath
$LastModifiedDate = "Unknown"

# Extract Last Modified date from header
try {
    $headerContent = Get-Content $PSCommandPath -First 30 -ErrorAction SilentlyContinue
    $lastModifiedLine = $headerContent | Where-Object { $_ -match 'Last Modified:\s*(.+)' } | Select-Object -First 1
    if ($lastModifiedLine -and $Matches[1]) {
        $LastModifiedDate = $Matches[1].Trim()
    }
} catch {
    # If we can't read the file, just use Unknown
}

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Script: $ScriptName" -ForegroundColor Cyan
Write-Host "Last Modified: $LastModifiedDate" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

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

    # Build settings with only essential, safe properties
    $settingsHt = @{
        # Core notification settings
        'mailNotificationsEnabled'        = if ($oldSettings.PSObject.Properties['MailNotificationsEnabled']) { $oldSettings.MailNotificationsEnabled } else { $true }
        'reminderNotificationsEnabled'    = if ($oldSettings.PSObject.Properties['ReminderNotificationsEnabled']) { $oldSettings.ReminderNotificationsEnabled } else { $true }

        # Approval requirements
        'justificationRequiredOnApproval' = if ($oldSettings.PSObject.Properties['JustificationRequiredOnApproval']) { $oldSettings.JustificationRequiredOnApproval } else { $true }

        # For one-time reviews, disable recommendations to avoid complex insight settings
        'recommendationsEnabled'          = $false

        # Default decision when reviewers don't respond
        'defaultDecisionEnabled'          = if ($oldSettings.PSObject.Properties['DefaultDecisionEnabled']) { $oldSettings.DefaultDecisionEnabled } else { $false }
        'defaultDecision'                 = if ($oldSettings.PSObject.Properties['DefaultDecision']) { $oldSettings.DefaultDecision } else { 'None' }

        # Auto-apply decisions to resource
        'autoApplyDecisionsEnabled'       = if ($oldSettings.PSObject.Properties['AutoApplyDecisionsEnabled']) { $oldSettings.AutoApplyDecisionsEnabled } else { $false }

        # Scheduling - set by us for one-time review
        'instanceDurationInDays'          = $InstanceDurationInDays
        'recurrence'                      = buildOneTimeRecurrence $StartDate
    }

    # Handle recommendationLookBackDuration - extract days and create ISO 8601 duration
    if ($oldSettings.PSObject.Properties['RecommendationLookBackDuration']) {
        $lookback = $oldSettings.RecommendationLookBackDuration
        if ($lookback) {
            # Try to convert to ISO 8601 duration string using just the days
            $durationStr = toIso8601Duration $lookback
            if ($durationStr) {
                $settingsHt['recommendationLookBackDuration'] = $durationStr
                Write-Host "  Including recommendationLookBackDuration: $durationStr" -ForegroundColor Cyan
            } else {
                Write-Host "  Warning: recommendationLookBackDuration exists but could not be converted - excluding from new definition" -ForegroundColor Yellow
            }
        }
    }

    # Copy applyActions if present AND non-empty (actions to apply on denied guest users)
    if ($oldSettings.PSObject.Properties['ApplyActions']) {
        $actions = unwrapGraphObject $oldSettings.ApplyActions
        # Filter out empty objects that cause validation errors
        $validActions = @($actions | Where-Object {
                $_ -and ($_ -is [System.Collections.IDictionary]) -and $_.Keys.Count -gt 0
            })
        if ($validActions.Count -gt 0) {
            $settingsHt['applyActions'] = $validActions
        }
    }

    return $settingsHt
}

# ---------------- Scope Simplification ----------------
function Simplify-Scope {
    param(
        [hashtable]$ScopeHt
    )

    # If scope is empty or null, return null (don't create a default)
    if (-not $ScopeHt -or $ScopeHt.Keys.Count -eq 0) {
        Write-Host "Scope is empty - will not be included in new definition" -ForegroundColor Yellow
        return $null
    }

    # If it's already a simple scope, return as-is
    if ($ScopeHt.'@odata.type' -eq '#microsoft.graph.accessReviewQueryScope') {
        return $ScopeHt
    }

    # If it's a principalResourceMembershipsScope, convert to simple scope
    if ($ScopeHt.'@odata.type' -eq '#microsoft.graph.principalResourceMembershipsScope') {
        Write-Host "Simplifying principalResourceMembershipsScope to basic accessReviewQueryScope" -ForegroundColor Yellow

        # Use the first resourceScope as the base scope
        if ($ScopeHt.ContainsKey('resourceScopes') -and $ScopeHt['resourceScopes'] -and $ScopeHt['resourceScopes'].Count -gt 0) {
            $firstResource = $ScopeHt['resourceScopes'][0]

            $simplifiedScope = @{
                '@odata.type' = '#microsoft.graph.accessReviewQueryScope'
                'query'       = $firstResource['query']
                'queryType'   = $firstResource['queryType']
            }

            if ($ScopeHt['resourceScopes'].Count -gt 1) {
                Write-Host "  Note: Original scope had $($ScopeHt['resourceScopes'].Count) resource scopes. Using first scope only." -ForegroundColor Yellow
            }

            Write-Host "  Simplified query: $($simplifiedScope['query'])" -ForegroundColor Cyan
            return $simplifiedScope
        } else {
            Write-Host "  Warning: principalResourceMembershipsScope has no resourceScopes. Scope will be omitted." -ForegroundColor Yellow
            return $null
        }
    }

    # Return as-is if we don't know how to simplify
    return $ScopeHt
}

# ---------------- Problematic Settings Detection ----------------
function Test-ProblematicSettings {
    param(
        [object]$Settings
    )

    $problematicSettings = @()

    # Check for problematic settings that were removed in simplified implementation
    # Note: recommendationLookBackDuration is now handled properly by extracting days

    if ($Settings.PSObject.Properties['RecommendationInsightSettings']) {
        $value = $Settings.RecommendationInsightSettings
        # Only warn if it has actual data (not empty array or array of empty objects)
        if ($value -and $value.Count -gt 0) {
            $hasNonEmptyItems = $false
            foreach ($item in $value) {
                # Skip null items
                if (-not $item) { continue }

                # Check if item is a hashtable/dictionary with no keys
                if ($item -is [System.Collections.IDictionary] -and $item.Keys.Count -eq 0) {
                    continue
                }

                # For Graph SDK objects, check if they're effectively empty
                $props = $item.PSObject.Properties | Where-Object { $_.Name -ne 'AdditionalProperties' }
                if ($props.Count -eq 0) {
                    # No properties besides AdditionalProperties
                    if ($item.PSObject.Properties['AdditionalProperties']) {
                        $additionalProps = $item.AdditionalProperties
                        # Check if AdditionalProperties is empty or only has OData metadata
                        if (-not $additionalProps -or $additionalProps.Keys.Count -eq 0) {
                            continue
                        }
                        # Check if only has @odata.type property (which is just metadata)
                        $nonODataKeys = @($additionalProps.Keys | Where-Object { -not $_.StartsWith('@odata.') })
                        if ($nonODataKeys.Count -eq 0) {
                            continue
                        }
                    } else {
                        # No properties at all
                        continue
                    }
                }

                # If we get here, the item has actual data
                $hasNonEmptyItems = $true
                break
            }
            if ($hasNonEmptyItems) {
                $problematicSettings += "recommendationInsightSettings (complex objects may cause validation errors)"
            }
        }
    }

    if ($Settings.PSObject.Properties['AccessRecommendationsEnabled']) {
        $problematicSettings += "accessRecommendationsEnabled (not supported in v1.0 API)"
    }

    return $problematicSettings
}

# ---------------- Output log generation ----------------
function Save-OutputLog {
    param(
        [string]$DefinitionId,
        [object]$Definition,
        [object]$Body,
        [string]$ErrorMessage,
        [string]$ErrorDetails,
        [string[]]$ProblematicSettings,
        [bool]$IsSuccess = $false,
        [bool]$PromptUser = $false
    )

    $shouldGenerate = $true

    # If prompting is required (logs were suppressed and error occurred)
    if ($PromptUser) {
        Write-Host "`n" -NoNewline
        Write-Host "Would you like to generate an output log for troubleshooting? [Y/n] (Auto-yes in 5 seconds): " -ForegroundColor Yellow -NoNewline

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

        if ($response -eq 'n') {
            Write-Host "No"
            $shouldGenerate = $false
        } else {
            Write-Host "Yes"
        }
    }

    if ($shouldGenerate) {
        $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
        $outputFileName = "output-accessreview-$DefinitionId-$timestamp.json"

        # Build output in specific order: script, invocation, originalDefinition, body
        $outputData = [ordered]@{
            Script     = [ordered]@{
                ScriptName         = $ScriptName
                ScriptLastModified = $LastModifiedDate
            }
            Invocation = [ordered]@{
                DefinitionId = $DefinitionId
                Timestamp    = $timestamp
                Success      = $IsSuccess
            }
        }

        # Add problematic settings to invocation if detected
        if ($ProblematicSettings -and $ProblematicSettings.Count -gt 0) {
            $outputData['Invocation']['ProblematicSettingsDetected'] = $ProblematicSettings
        }

        # Add error details to invocation if not successful
        if (-not $IsSuccess) {
            $outputData['Invocation']['ErrorMessage'] = $ErrorMessage
            $outputData['Invocation']['ErrorDetails'] = $ErrorDetails
        }

        # Properly unwrap the Graph SDK object to include all AdditionalProperties
        $outputData['OriginalDefinition'] = unwrapGraphObject $Definition
        $outputData['Body'] = $Body

        $outputData | ConvertTo-Json -Depth 100 | Out-File $outputFileName -Encoding UTF8

        if ($PromptUser) {
            Write-Host "`nOutput log saved: $outputFileName" -ForegroundColor Green
        } else {
            Write-Host "Output log saved: $outputFileName" -ForegroundColor Cyan
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
        [bool]$IsWhatIf,
        [bool]$SuppressLogs,
        [bool]$ShowBody
    )

    Write-Host "`n========================================"
    Write-Host "Processing Definition ID: $DefinitionId"
    Write-Host "========================================`n"

    try {
        $old = Get-MgIdentityGovernanceAccessReviewDefinition -AccessReviewScheduleDefinitionId $DefinitionId

        Write-Host "Source Definition: $($old.DisplayName)" -ForegroundColor White

        # Check for problematic settings
        $problematicSettings = @()
        if ($old.Settings) {
            $problematicSettings = Test-ProblematicSettings -Settings $old.Settings
            if ($problematicSettings.Count -gt 0) {
                Write-Host "`n[WARNING] Detected problematic settings in source definition:" -ForegroundColor Yellow
                foreach ($setting in $problematicSettings) {
                    Write-Host "  - $setting" -ForegroundColor Yellow
                }
                Write-Host "These settings will NOT be copied to the new definition." -ForegroundColor Yellow
                Write-Host "See output log for details.`n" -ForegroundColor Yellow
            }
        }

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

        # Check if scope is empty (common in some definitions)
        if (-not $scopeHt -or $scopeHt.Keys.Count -eq 0 -or (-not $scopeHt.'@odata.type' -and -not $scopeHt.ContainsKey('query'))) {
            Write-Host "Warning: Source definition has no valid scope. Scope will be omitted." -ForegroundColor Yellow
            $scopeHt = $null
        } else {
            # Ensure @odata.type is present
            if (-not $scopeHt.'@odata.type') {
                # If missing, try to determine from structure
                if ($scopeHt.ContainsKey('principalScopes') -or $scopeHt.ContainsKey('resourceScopes')) {
                    $scopeHt.'@odata.type' = '#microsoft.graph.principalResourceMembershipsScope'
                } elseif ($scopeHt.ContainsKey('query')) {
                    $scopeHt.'@odata.type' = '#microsoft.graph.accessReviewQueryScope'
                }
            }

            Write-Host "Using scope type: $($scopeHt.'@odata.type')" -ForegroundColor Green

            # Always simplify scope to ensure tenant compatibility
            $scopeHt = Simplify-Scope -ScopeHt $scopeHt
            if ($scopeHt) {
                Write-Host "Final scope type: $($scopeHt.'@odata.type')" -ForegroundColor Green
            }
        }

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
            reviewers               = $reviewersArr
            settings                = $settingsHt
        }

        # Add scope only if it exists
        if ($scopeHt) {
            $body['scope'] = $scopeHt
        }

        # Add fallbackReviewers if they exist
        $fallbackReviewersArr = @(convertReviewerScopesToArray $old.FallbackReviewers)
        if ($fallbackReviewersArr.Count -gt 0) {
            $body['fallbackReviewers'] = $fallbackReviewersArr
        }

        removeNullsRecursively -obj $body

        # Display body if requested
        if ($ShowBody) {
            Write-Host "`n--- Payload to Graph ---"
            $body | ConvertTo-Json -Depth 100 | Write-Host
        }

        # Always show scope details for debugging
        if ($body.ContainsKey('scope')) {
            Write-Host "`nScope being sent to API:" -ForegroundColor Magenta
            Write-Host "  @odata.type: $($body.scope.'@odata.type')" -ForegroundColor Magenta
            Write-Host "  Keys: $($body.scope.Keys -join ', ')" -ForegroundColor Magenta
            if ($body.scope.ContainsKey('query')) {
                Write-Host "  query: $($body.scope.query)" -ForegroundColor Magenta
            }
        }

        if ($IsWhatIf) {
            Write-Host "`n[WHATIF] Would create Access Review Definition with the payload shown above" -ForegroundColor Yellow
            return $null
        }

        try {
            $newDef = New-MgIdentityGovernanceAccessReviewDefinition -BodyParameter $body -ErrorAction Stop
            Write-Host "`nCreated Access Review Definition Id: $($newDef.Id)" -ForegroundColor Green

            # Save output log for successful creation (unless suppressed)
            if (-not $SuppressLogs) {
                Save-OutputLog -DefinitionId $DefinitionId -Definition $old -Body $body `
                    -ErrorMessage "" -ErrorDetails "" -ProblematicSettings $problematicSettings `
                    -IsSuccess $true -PromptUser $false
            }

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
            # Check for Custom Scoping Conditions error (tenant lacks advanced scope feature)
            if ($errorDetails -match "Custom Scoping Conditions" -or $errorMsg -match "Custom Scoping Conditions") {
                Write-Host "`n[INFO]" -ForegroundColor Yellow
                Write-Host "This tenant does not support advanced scope features (principalResourceMembershipsScope)." -ForegroundColor Yellow
                Write-Host "The script automatically simplifies scopes to basic accessReviewQueryScope for compatibility." -ForegroundColor Yellow
                Write-Host "This error should not occur - please check the output log for details." -ForegroundColor Yellow
            }
            Write-Host "`nDefinition ID: $DefinitionId" -ForegroundColor Cyan

            # Save output log (prompt only if suppressed)
            if ($SuppressLogs) {
                Save-OutputLog -DefinitionId $DefinitionId -Definition $old -Body $body `
                    -ErrorMessage $errorMsg -ErrorDetails $errorDetails -ProblematicSettings $problematicSettings `
                    -IsSuccess $false -PromptUser $true
            } else {
                Save-OutputLog -DefinitionId $DefinitionId -Definition $old -Body $body `
                    -ErrorMessage $errorMsg -ErrorDetails $errorDetails -ProblematicSettings $problematicSettings `
                    -IsSuccess $false -PromptUser $false
            }

            Write-Host "`nSkipping this definition and continuing..." -ForegroundColor Yellow
            return $null
        }
    } catch {
        Write-Host "`n========================================" -ForegroundColor Red
        Write-Host "ERROR: Failed to Process Definition" -ForegroundColor Red
        Write-Host "========================================" -ForegroundColor Red
        Write-Host "`nDefinition ID: $DefinitionId" -ForegroundColor Cyan
        Write-Host "`nError: $_" -ForegroundColor Red

        # Save output log (prompt only if suppressed)
        if ($SuppressLogs) {
            Save-OutputLog -DefinitionId $DefinitionId -Definition $null -Body $null `
                -ErrorMessage $_.Exception.Message -ErrorDetails "" -ProblematicSettings @() `
                -IsSuccess $false -PromptUser $true
        } else {
            Save-OutputLog -DefinitionId $DefinitionId -Definition $null -Body $null `
                -ErrorMessage $_.Exception.Message -ErrorDetails "" -ProblematicSettings @() `
                -IsSuccess $false -PromptUser $false
        }

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
        -IsWhatIf $WhatIf.IsPresent `
        -SuppressLogs $SuppressOutputLogs.IsPresent `
        -ShowBody $DisplayBody.IsPresent
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
