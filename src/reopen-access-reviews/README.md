# Reopen Access Reviews

## Overview

`reopen-accessreview.ps1` clones existing access review schedule definitions and creates new one-time reviews.

It supports:

- Single definition ID input
- Batch input from a file of IDs
- Optional dry-run output (`-WhatIf`)

## Files

- `reopen-accessreview.ps1`: Main script.
- `id-to-reopen.txt`: Example input file for `-FromFile`.

## Prerequisites

- Microsoft Graph PowerShell module: `Microsoft.Graph.Identity.Governance`
- Graph permission: `AccessReview.ReadWrite.All`
- Signed-in context with rights to create access review definitions

## Usage

Single definition ID:

```powershell
.\src\reopen-access-reviews\reopen-accessreview.ps1 -OldDefinitionId "<definition-id>"
```

From file of IDs:

```powershell
.\src\reopen-access-reviews\reopen-accessreview.ps1 -FromFile .\src\reopen-access-reviews\id-to-reopen.txt
```

Dry run:

```powershell
.\src\reopen-access-reviews\reopen-accessreview.ps1 -OldDefinitionId "<definition-id>" -WhatIf
```

Useful optional parameters:

- `-StartDate` (default: today)
- `-InstanceDurationInDays` (default: 30)
- `-NewDisplayNameSuffix`
- `-DumpDefinition`
- `-SuppressOutputLogs`
- `-DisplayBody`

## Notes

- The script normalizes source scoping to improve compatibility across tenant configurations.
- Provide either `-OldDefinitionId` or `-FromFile`.
