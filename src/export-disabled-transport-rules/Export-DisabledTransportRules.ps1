<#
.SYNOPSIS
    Exports disabled Exchange transport rules with full and summary outputs.

.DESCRIPTION
    Retrieves transport rules, filters for disabled rules, and exports two files:
    a full JSON object dump and a flattened CSV summary.

.PARAMETER OutputFolder
    Output directory for exported files. Defaults to a timestamped folder in
    the current working directory.

.PARAMETER ExchangeOnline
    When provided, the script connects to Exchange Online before querying rules.

.PARAMETER TenantUPN
    Optional UPN for Exchange Online sign-in.

.EXAMPLE
    .\Export-DisabledTransportRules.ps1 -ExchangeOnline

.EXAMPLE
    .\Export-DisabledTransportRules.ps1 -ExchangeOnline -TenantUPN admin@contoso.com

.EXAMPLE
    .\Export-DisabledTransportRules.ps1 -OutputFolder .\out\transport-rules

.OUTPUTS
  - DisabledTransportRules_FULL.json  (full fidelity)
  - DisabledTransportRules_SUMMARY.csv (flat summary)

.NOTES
  - Exchange Online: requires ExchangeOnlineManagement module and appropriate permissions.
  - Exchange On-Prem: run in Exchange Management Shell (EMS) or load snap-in.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputFolder = (Join-Path $PWD ("TransportRuleExport_" + (Get-Date -Format "yyyyMMdd_HHmmss"))),

    [Parameter(Mandatory = $false)]
    [switch]$ExchangeOnline,

    [Parameter(Mandatory = $false)]
    [string]$TenantUPN
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function New-OutputFolder {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

function Assert-TransportRuleCommandAvailable {
    if (-not (Get-Command -Name Get-TransportRule -ErrorAction SilentlyContinue)) {
        throw "Get-TransportRule is not available. For Exchange Online use -ExchangeOnline. For on-prem, run in Exchange Management Shell or load Exchange cmdlets."
    }
}

function Connect-ExchangeOnlineIfNeeded {
    param([switch]$DoConnect, [string]$UPN)

    if (-not $DoConnect) { return }

    if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
        throw "ExchangeOnlineManagement module not found. Install with: Install-Module ExchangeOnlineManagement"
    }

    Import-Module ExchangeOnlineManagement -ErrorAction Stop

    # If already connected, this will typically no-op.
    if ([string]::IsNullOrWhiteSpace($UPN)) {
        Connect-ExchangeOnline -ShowBanner:$false
    } else {
        Connect-ExchangeOnline -UserPrincipalName $UPN -ShowBanner:$false
    }
}

function ConvertTo-JoinedString {
    param([object]$Value)
    if ($null -eq $Value) { return "" }
    if ($Value -is [System.Array]) { return ($Value | ForEach-Object { "$_" }) -join "; " }
    return "$Value"
}

function Get-PropertyValue {
    param(
        [Parameter(Mandatory = $true)][object]$InputObject,
        [Parameter(Mandatory = $true)][string]$Name
    )

    $prop = $InputObject.PSObject.Properties[$Name]
    if ($null -ne $prop) {
        return $prop.Value
    }

    return $null
}

function Test-RuleIsDisabled {
    param([Parameter(Mandatory = $true)][object]$Rule)

    $enabled = Get-PropertyValue -InputObject $Rule -Name "Enabled"
    if ($null -ne $enabled) {
        return (-not [bool]$enabled)
    }

    $state = Get-PropertyValue -InputObject $Rule -Name "State"
    if (-not [string]::IsNullOrWhiteSpace("$state")) {
        return ("$state" -ieq "Disabled")
    }

    return $false
}

New-OutputFolder -Path $OutputFolder
Connect-ExchangeOnlineIfNeeded -DoConnect:$ExchangeOnline -UPN $TenantUPN
Assert-TransportRuleCommandAvailable

Write-Host "Getting disabled transport rules..." -ForegroundColor Cyan

# Pull all rules; filter for disabled
$disabledRules = Get-TransportRule -ResultSize Unlimited |
    Where-Object { Test-RuleIsDisabled -Rule $_ }

Write-Host ("Found {0} disabled rule(s)." -f $disabledRules.Count) -ForegroundColor Green

# FULL dump object builder
$full = foreach ($r in $disabledRules) {
    # Re-query each rule to ensure we get the full property set (some environments return partials)
    $rule = Get-TransportRule -Identity (Get-PropertyValue -InputObject $r -Name "Identity")

    $enabledValue = Get-PropertyValue -InputObject $rule -Name "Enabled"
    if ($null -eq $enabledValue) {
        $enabledValue = -not (Test-RuleIsDisabled -Rule $rule)
    }

    [pscustomobject]@{
        Name                        = Get-PropertyValue -InputObject $rule -Name "Name"
        Identity                    = "$(Get-PropertyValue -InputObject $rule -Name "Identity")"
        Guid                        = "$(Get-PropertyValue -InputObject $rule -Name "Guid")"
        Enabled                     = $enabledValue
        Mode                        = "$(Get-PropertyValue -InputObject $rule -Name "Mode")"              # Enforce/Audit/AuditAndNotify (where applicable)
        Priority                    = Get-PropertyValue -InputObject $rule -Name "Priority"
        State                       = "$(Get-PropertyValue -InputObject $rule -Name "State")"             # Often present in some builds
        Comments                    = Get-PropertyValue -InputObject $rule -Name "Comments"
        Description                 = Get-PropertyValue -InputObject $rule -Name "Description"

        # Core logic blocks (these are the “settings/actions/conditions” the user wants)
        Conditions                  = Get-PropertyValue -InputObject $rule -Name "Conditions"
        Exceptions                  = Get-PropertyValue -InputObject $rule -Name "Exceptions"
        Actions                     = Get-PropertyValue -InputObject $rule -Name "Actions"

        # Additional common rule fields people care about
        SentTo                      = Get-PropertyValue -InputObject $rule -Name "SentTo"
        SentToMemberOf              = Get-PropertyValue -InputObject $rule -Name "SentToMemberOf"
        From                        = Get-PropertyValue -InputObject $rule -Name "From"
        FromMemberOf                = Get-PropertyValue -InputObject $rule -Name "FromMemberOf"
        SubjectContainsWords        = Get-PropertyValue -InputObject $rule -Name "SubjectContainsWords"
        BodyContainsWords           = Get-PropertyValue -InputObject $rule -Name "BodyContainsWords"
        HeaderContainsMessageHeader = Get-PropertyValue -InputObject $rule -Name "HeaderContainsMessageHeader"
        HeaderContainsWords         = Get-PropertyValue -InputObject $rule -Name "HeaderContainsWords"
        RecipientDomainIs           = Get-PropertyValue -InputObject $rule -Name "RecipientDomainIs"
        SenderDomainIs              = Get-PropertyValue -InputObject $rule -Name "SenderDomainIs"

        # Execution / behavior
        StopRuleProcessing          = Get-PropertyValue -InputObject $rule -Name "StopRuleProcessing"
        RuleErrorAction             = "$(Get-PropertyValue -InputObject $rule -Name "RuleErrorAction")"
        ExceptIf                    = Get-PropertyValue -InputObject $rule -Name "ExceptIf"               # Some versions expose ExceptIf conditions

        # Timestamps (availability varies by platform/version)
        WhenChanged                 = Get-PropertyValue -InputObject $rule -Name "WhenChanged"
        WhenCreated                 = Get-PropertyValue -InputObject $rule -Name "WhenCreated"

        # Catch-all: include the entire object as a stringified property bag, useful for diffs
        AllProperties               = ($rule | Select-Object * )
    }
}

# Export FULL to JSON (deep to keep nested action objects)
$fullJsonPath = Join-Path $OutputFolder "DisabledTransportRules_FULL.json"
$full | ConvertTo-Json -Depth 12 | Out-File -FilePath $fullJsonPath -Encoding UTF8

# Export SUMMARY to CSV (flattened fields)
$summary = $full | Select-Object `
    Name, Enabled, Priority, Mode, StopRuleProcessing, RuleErrorAction, WhenCreated, WhenChanged, `
@{n = "Conditions_Flat"; e = { ConvertTo-JoinedString $_.Conditions } }, `
@{n = "Exceptions_Flat"; e = { ConvertTo-JoinedString $_.Exceptions } }, `
@{n = "Actions_Flat"; e = { ConvertTo-JoinedString $_.Actions } }

$summaryCsvPath = Join-Path $OutputFolder "DisabledTransportRules_SUMMARY.csv"
$summary | Export-Csv -Path $summaryCsvPath -NoTypeInformation -Encoding UTF8

Write-Host "Export complete:" -ForegroundColor Green
Write-Host (" - {0}" -f $fullJsonPath)
Write-Host (" - {0}" -f $summaryCsvPath)

# Optional: Disconnect Exchange Online cleanly
if ($ExchangeOnline) {
    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
}