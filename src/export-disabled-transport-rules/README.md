# Export Disabled Transport Rules

## Overview

`Export-DisabledTransportRules.ps1` exports disabled Exchange transport rules and writes:

- `DisabledTransportRules_FULL.json` with full rule fidelity, including nested condition/action data.
- `DisabledTransportRules_SUMMARY.csv` with flattened values for quick review and reporting.

It supports both Exchange Online and Exchange on-premises environments.

## Script Location

- `src/export-disabled-transport-rules/Export-DisabledTransportRules.ps1`

## Prerequisites

### Exchange Online

- Exchange admin permissions required to run `Get-TransportRule`.
- `ExchangeOnlineManagement` PowerShell module installed.

Install module if needed:

```powershell
Install-Module ExchangeOnlineManagement -Scope CurrentUser
```

### Exchange On-Premises

- Run from Exchange Management Shell (EMS), or
- Load Exchange snap-ins/modules before running the script.

## Parameters

- `-OutputFolder` (optional): Output directory for exported files.
- `-ExchangeOnline` (optional switch): Connects to Exchange Online before collecting rules.
- `-TenantUPN` (optional): UPN used when connecting to Exchange Online.

If `-OutputFolder` is not provided, a timestamped folder is created in the current location.

## Usage

### Exchange Online (interactive sign-in)

```powershell
.\src\export-disabled-transport-rules\Export-DisabledTransportRules.ps1 -ExchangeOnline
```

### Exchange Online (specific UPN)

```powershell
.\src\export-disabled-transport-rules\Export-DisabledTransportRules.ps1 -ExchangeOnline -TenantUPN admin@contoso.com
```

### Exchange On-Premises

```powershell
.\src\export-disabled-transport-rules\Export-DisabledTransportRules.ps1 -OutputFolder .\out\transport-rules
```

## Output Files

- `DisabledTransportRules_FULL.json`: Full object export with nested conditions, exceptions, actions, and additional metadata.
- `DisabledTransportRules_SUMMARY.csv`: Flattened summary focused on rule name, mode, priority, behavior, and core logic blocks.

## Notes

- The script safely handles environments where some transport rule properties may not be available.
- It reconnects each disabled rule by identity to capture the most complete property set.
- If `-ExchangeOnline` is used, it disconnects from Exchange Online at the end.
