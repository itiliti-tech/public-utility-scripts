<#
.SYNOPSIS
    Moves items from an Outlook Online Archive mailbox back to the primary mailbox.

.DESCRIPTION
    Iterates every folder in the Online Archive, mirrors the folder structure into
    the primary mailbox, and moves eligible items.  After migration, optionally
    removes empty folder branches from the archive.

    Key behaviours specific to Online Archive / Cached Exchange:
    - Leaf-folder .Delete() silently fails; we delete from the TOP of each empty branch.
    - Exchange rejects folder deletion when a same-named folder already exists in
      Deleted Items; we rename before deleting to avoid collisions.
    - COM object references can go stale; items are re-acquired from the collection
      on retry rather than reusing a cached reference.

.NOTES
    See Outlook-COM-Notes.md for detailed findings on COM / Exchange / Online Archive quirks.
#>

#requires -Version 5.1

# ═══════════════════════════════════════════════════════════════════════════════
# Configuration
# ═══════════════════════════════════════════════════════════════════════════════

# Date filter — only move items received on or before this date.  $null = move everything.
$NewestDate = $null  # e.g. [datetime]"2025-01-01"

# Target mailbox — set both to $null to use your own mailbox.
# For delegated access, the primary store appears as the user's full name
# and the Online Archive store appears with the user's email address.
# Examples:  $TargetMailboxName = 'First Last'   $TargetMailboxEmail = 'user@domain.com'
$TargetMailboxName = $null
$TargetMailboxEmail = $null

# Dry-run mode — preview without moving anything or creating folders
$DryRun = $true

# Remove empty folder branches from the archive after migration
$CleanupEmptyFolders = $false

# Outlook item classes to move.  Anything not listed here is skipped and reported.
# Reference: 43=MailItem  53=MeetingRequest  54=MeetingCancel  55=MeetingRespNeg
#            56=MeetingRespPos  57=MeetingRespTent  104=SharingItem  46=PostItem
$MoveItemClasses = @(43)

# Folders to skip entirely — neither items nor subfolders are processed
$SkipFolders = @(
    'Deleted Items'
    'Calendar'
    'Contacts'
    'Tasks'
    'Notes'
    'Journal'
    'Outbox'
    'Sync Issues'
    'Conflicts'
    'Local Failures'
    'Server Failures'
    'Files'
)

# Folders where items are still migrated but the folder itself is never deleted
$ProtectFromDeletion = @(
    'Archive'
)

# Adaptive throttle / retry for item moves
$ThrottleConfig = @{
    BaseDelayMs = 0        # delay between moves when no errors
    MaxDelayMs  = 30000    # ceiling for exponential backoff (30 s)
    MaxRetries  = 10       # per-item retry limit
}

# ═══════════════════════════════════════════════════════════════════════════════
# Runtime state
# ═══════════════════════════════════════════════════════════════════════════════

$script:CurrentDelayMs = $ThrottleConfig.BaseDelayMs
$script:Stats = @{
    Moved          = 0
    Folders        = 0
    Errors         = 0
    CleanedFolders = 0
}
$script:SkippedItems = [System.Collections.Generic.List[PSCustomObject]]::new()

# ═══════════════════════════════════════════════════════════════════════════════
# Helpers — COM / Folder navigation
# ═══════════════════════════════════════════════════════════════════════════════

function Write-Ts {
    <# Emit a timestamped, coloured log line. #>
    param(
        [string]$Message,
        [System.ConsoleColor]$Color = 'White',
        [switch]$NoNewline
    )
    $stamp = (Get-Date).ToString('HH:mm:ss')
    Write-Host "[$stamp] $Message" -ForegroundColor $Color -NoNewline:$NoNewline
}

function Release-ComObject {
    <# Safely release a COM reference. #>
    param($Object)
    if ($Object) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object) | Out-Null } catch {}
    }
}

function Resolve-FolderPath {
    <# Walk an array of folder names under a root; return $null if any part is missing. #>
    param(
        [Parameter(Mandatory)] $Root,
        [Parameter(Mandatory)] [string[]]$Parts
    )
    $current = $Root
    foreach ($name in $Parts) {
        try { $current = $current.Folders.Item($name) }
        catch { return $null }
    }
    return $current
}

function Ensure-FolderPath {
    <# Walk folder Parts under Root, creating missing folders.  Returns $null in DryRun. #>
    param(
        [Parameter(Mandatory)] $Root,
        [Parameter(Mandatory)] [string[]]$Parts
    )
    $current = $Root
    foreach ($name in $Parts) {
        $child = $null
        try { $child = $current.Folders.Item($name) } catch {}
        if (-not $child) {
            if ($DryRun) {
                Write-Ts "  [DRY RUN] Would create folder: $name (under $($current.FolderPath))" -Color Magenta
                return $null
            }
            Write-Ts "  Creating folder: $name (under $($current.FolderPath))" -Color Yellow
            $child = $current.Folders.Add($name)
        }
        $current = $child
    }
    return $current
}

function Get-RelativePathParts {
    <# Return the folder-name chain from $Root down to $Folder (exclusive of Root). #>
    param(
        [Parameter(Mandatory)] $Folder,
        [Parameter(Mandatory)] $Root
    )
    $parts = [System.Collections.Generic.List[string]]::new()
    $cursor = $Folder
    while ($cursor.EntryID -ne $Root.EntryID) {
        $parts.Insert(0, $cursor.Name)
        $cursor = $cursor.Parent
    }
    return $parts
}

function Test-IsUnderProtectedFolder {
    <# Walk up the tree to see if any ancestor is in the skip list. #>
    param([Parameter(Mandatory)] $Folder)
    $cursor = $Folder.Parent
    while ($cursor) {
        if ($cursor.Name -in $SkipFolders) { return $true }
        try { $cursor = $cursor.Parent } catch { break }
    }
    return $false
}

function Test-BranchEmpty {
    <# Recursively check that a folder and ALL descendants contain zero items. #>
    param([Parameter(Mandatory)] $Folder)
    if ($Folder.Items.Count -gt 0) { return $false }
    foreach ($sub in $Folder.Folders) {
        if (-not (Test-BranchEmpty -Folder $sub)) { return $false }
    }
    return $true
}

# ═══════════════════════════════════════════════════════════════════════════════
# Core — Item migration
# ═══════════════════════════════════════════════════════════════════════════════

function Move-ArchiveItems {
    <#
    .SYNOPSIS  Move eligible items from a single archive folder to the matching primary folder.
    .DESCRIPTION
        Items are processed via a shrinking-index pattern: on a successful move the
        item vanishes from the COM collection so we re-read .Items.Count; on a skip
        we increment the index to step past it.
    #>
    param(
        [Parameter(Mandatory)] $SourceFolder,
        [Parameter(Mandatory)] $TargetRoot,
        [Parameter(Mandatory)] [string[]]$RelativeParts
    )

    $itemCount = $SourceFolder.Items.Count
    if ($itemCount -eq 0) {
        Write-Ts "  (empty)" -Color DarkGray
        return
    }

    $targetPath = "\\$($TargetRoot.Name)\$($RelativeParts -join '\')"
    $targetFolder = Resolve-FolderPath -Root $TargetRoot -Parts $RelativeParts
    $targetCreatedForDryRun = $false

    Write-Ts "  $itemCount item(s) — processing..." -Color DarkGray

    $moved = $skippedClass = $skippedDate = $errors = 0
    $index = 1
    $remaining = $itemCount

    while ($index -le $remaining) {
        # Acquire item
        $item = $null
        try { $item = $SourceFolder.Items.Item($index) }
        catch {
            Write-Host ''
            Write-Ts "  Could not access item $index : $_" -Color Red
            $index++; continue
        }

        # ── Filter: item class ──
        if ($item.Class -notin $MoveItemClasses) {
            $skippedClass++
            $msgClass = ''; try { $msgClass = $item.MessageClass } catch {}
            $subj = ''; try { $subj = $item.Subject }      catch {}
            $script:SkippedItems.Add([PSCustomObject]@{
                    Folder   = $SourceFolder.FolderPath
                    Class    = $item.Class
                    MsgClass = $msgClass
                    Subject  = $subj
                })
            Release-ComObject $item
            $index++
            continue
        }

        # ── Filter: date ──
        if ($NewestDate) {
            $receivedTime = $null
            try { $receivedTime = $item.ReceivedTime } catch {}
            if ($receivedTime -and $receivedTime -gt $NewestDate) {
                $skippedDate++
                Release-ComObject $item
                $index++
                continue
            }
        }

        # ── Dry run ──
        if ($DryRun) {
            if (-not $targetCreatedForDryRun) {
                Write-Ts "  [DRY RUN] Would create folder: $targetPath" -Color Magenta
                $targetCreatedForDryRun = $true
            }
            $subj = ''; try { $subj = $item.Subject } catch {}
            Write-Ts "  [DRY RUN] Would move: '$subj'" -Color Magenta
            $script:Stats.Moved++
            $moved++
            Release-ComObject $item
            $index++
            continue
        }

        # ── Lazy-create target folder on first eligible item ──
        if (-not $targetFolder) {
            $targetFolder = Ensure-FolderPath -Root $TargetRoot -Parts $RelativeParts
            if (-not $targetFolder) { return }
            $targetPath = $targetFolder.FolderPath
        }

        # ── Move with retry + adaptive backoff ──
        $success = $false
        for ($retry = 0; $retry -le $ThrottleConfig.MaxRetries; $retry++) {
            try {
                if ($script:CurrentDelayMs -gt 0) {
                    Start-Sleep -Milliseconds $script:CurrentDelayMs
                }
                $item.Move($targetFolder) | Out-Null
                $success = $true

                $script:Stats.Moved++
                $moved++

                # Ease off throttle on success
                if ($script:CurrentDelayMs -gt $ThrottleConfig.BaseDelayMs) {
                    $script:CurrentDelayMs = [Math]::Max(
                        $ThrottleConfig.BaseDelayMs,
                        [int]($script:CurrentDelayMs * 0.75))
                }

                # Item vanished from collection — refresh count, keep index
                $remaining = $SourceFolder.Items.Count
                break
            } catch {
                if ($retry -lt $ThrottleConfig.MaxRetries) {
                    # Exponential backoff: 1 s → 2 s → 4 s … capped
                    $script:CurrentDelayMs = if ($script:CurrentDelayMs -lt 1000) { 1000 }
                    else { [Math]::Min($ThrottleConfig.MaxDelayMs, $script:CurrentDelayMs * 2) }

                    Write-Host ''
                    Write-Ts "  Retry $($retry+1)/$($ThrottleConfig.MaxRetries) item $index (${script:CurrentDelayMs} ms): $_" -Color Yellow

                    Release-ComObject $item
                    Start-Sleep -Milliseconds $script:CurrentDelayMs

                    # Re-acquire (COM ref may be stale after Exchange throttle)
                    try { $item = $SourceFolder.Items.Item($index) }
                    catch { Write-Ts "  Could not re-acquire item $index after sleep" -Color Red; break }
                } else {
                    Write-Host ''
                    Write-Ts "  FAILED after $($ThrottleConfig.MaxRetries) retries (item $index): $_" -Color Red
                    $script:Stats.Errors++
                    $errors++
                    $index++
                }
            }
        }

        Release-ComObject $item

        # ── Inline progress ──
        $done = $moved + $skippedClass + $skippedDate + $errors
        $label = if ($DryRun) { 'would-move' } else { 'moved' }
        $ts = (Get-Date).ToString('HH:mm:ss')
        Write-Host "`r$(' ' * 120)" -NoNewline
        Write-Host "`r[$ts]   $done/$itemCount  ($label`: $moved  skip: $($skippedClass + $skippedDate)  err: $errors  delay: $($script:CurrentDelayMs) ms)" -NoNewline
    }

    Write-Host ''
    Write-Ts "  Done — moved: $moved, skipped-class: $skippedClass, skipped-date: $skippedDate, errors: $errors" -Color DarkGray
    if ($moved -gt 0) { $script:Stats.Folders++ }
}

# ═══════════════════════════════════════════════════════════════════════════════
# Core — Recursive folder walk
# ═══════════════════════════════════════════════════════════════════════════════

function Invoke-ArchiveMigration {
    <# Recursively walk the archive tree, migrating items to the primary mailbox. #>
    param(
        [Parameter(Mandatory)] $Folder,
        [Parameter(Mandatory)] $ArchiveRoot,
        [Parameter(Mandatory)] $PrimaryRoot
    )

    if ($Folder.Name -in $SkipFolders) {
        Write-Ts "SKIP: $($Folder.FolderPath)" -Color DarkGray
        return
    }

    $parts = Get-RelativePathParts -Folder $Folder -Root $ArchiveRoot

    # Archive root itself — just recurse children
    if ($parts.Count -eq 0) {
        Write-Ts 'Enumerating archive root...' -Color DarkGray
        foreach ($sub in $Folder.Folders) {
            Invoke-ArchiveMigration -Folder $sub -ArchiveRoot $ArchiveRoot -PrimaryRoot $PrimaryRoot
        }
        return
    }

    Write-Ts "Entering: $($Folder.FolderPath)" -Color Cyan
    Move-ArchiveItems -SourceFolder $Folder -TargetRoot $PrimaryRoot -RelativeParts $parts

    foreach ($sub in $Folder.Folders) {
        Invoke-ArchiveMigration -Folder $sub -ArchiveRoot $ArchiveRoot -PrimaryRoot $PrimaryRoot
    }
}

# ═══════════════════════════════════════════════════════════════════════════════
# Core — Empty-branch cleanup
# ═══════════════════════════════════════════════════════════════════════════════

function Remove-EmptyBranches {
    <#
    .SYNOPSIS  Top-down removal of entirely-empty folder branches.
    .DESCRIPTION
        Exchange Online Archive silently ignores .Delete() on leaf folders.  The
        workaround is to delete from the highest point of a fully-empty branch.

        If a folder with the same name already sits in Deleted Items, Exchange also
        silently fails.  We rename the folder to a unique timestamp before deletion
        so the name never collides.
    #>
    param(
        [Parameter(Mandatory)] $Folder,
        [Parameter(Mandatory)] $DeletedItemsFolder
    )

    $removed = 0
    # Snapshot to avoid mutating the collection mid-iteration
    $children = @($Folder.Folders)

    foreach ($child in $children) {
        # Skip protected folders
        if ($child.Name -in $SkipFolders -or $child.Name -in $ProtectFromDeletion) {
            Write-Ts "  SKIP (protected): $($child.FolderPath)" -Color DarkGray
            continue
        }
        if (Test-IsUnderProtectedFolder -Folder $child) { continue }

        if (Test-BranchEmpty -Folder $child) {
            if ($DryRun) {
                Write-Ts "  [DRY RUN] Would remove branch: $($child.FolderPath)" -Color Magenta
                $removed++
                continue
            }

            $originalName = $child.Name
            $deleteName = $originalName

            # Avoid duplicate-name collision in Deleted Items
            $collision = $false
            foreach ($df in $DeletedItemsFolder.Folders) {
                if ($df.Name -eq $originalName) { $collision = $true; break }
            }
            if ($collision) {
                $deleteName = "${originalName}_cleanup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
                Write-Ts "  Renaming '$originalName' -> '$deleteName' (Deleted Items collision)" -Color DarkYellow
                $child.Name = $deleteName
            }

            Write-Ts "  Removing branch: $($Folder.FolderPath)\$originalName" -Color Yellow

            try {
                $child.Delete()

                # Verify it arrived in Deleted Items
                Start-Sleep -Milliseconds 500
                $found = $false
                foreach ($df in $DeletedItemsFolder.Folders) {
                    if ($df.Name -eq $deleteName) { $found = $true; break }
                }
                if ($found) {
                    Write-Ts "  [OK] '$originalName' confirmed in Deleted Items" -Color Green
                } else {
                    Write-Ts "  [WARN] '$originalName' NOT found in Deleted Items — may have failed" -Color Red
                }
                $removed++
            } catch {
                Write-Ts "  FAILED to remove '$originalName': $_" -Color Red
            }
        } else {
            # Branch still has items — recurse deeper
            $removed += Remove-EmptyBranches -Folder $child -DeletedItemsFolder $DeletedItemsFolder
        }
    }
    return $removed
}

# ═══════════════════════════════════════════════════════════════════════════════
# Execution
# ═══════════════════════════════════════════════════════════════════════════════

# ── Connect to Outlook ──
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace('MAPI')

# ── Locate stores ──
if ($TargetMailboxName -and $TargetMailboxEmail) {
    $escapedName = [regex]::Escape($TargetMailboxName)
    $escapedEmail = [regex]::Escape($TargetMailboxEmail)

    # Primary store displays as the user's full name
    $primaryStore = $namespace.Stores | Where-Object {
        $_.DisplayName -match $escapedName -and $_.DisplayName -notmatch 'Online Archive'
    }
    # Archive store displays as "Online Archive - user@domain.com"
    $archiveStore = $namespace.Stores | Where-Object {
        $_.DisplayName -match 'Online Archive' -and $_.DisplayName -match $escapedEmail
    }

    if (-not $archiveStore -or -not $primaryStore) {
        Write-Host "Could not find stores for '$TargetMailboxName' / '$TargetMailboxEmail'.  Available stores:" -ForegroundColor Yellow
        foreach ($s in $namespace.Stores) {
            Write-Host "  - $($s.DisplayName)  (Type: $($s.ExchangeStoreType))" -ForegroundColor Cyan
        }
        if (-not $primaryStore) { Write-Host '  >> No matching primary store found (looked for name)' -ForegroundColor Red }
        if (-not $archiveStore) { Write-Host '  >> No matching archive store found (looked for email)' -ForegroundColor Red }
        return
    }
} else {
    # Default: own mailbox
    $archiveStore = $namespace.Stores | Where-Object { $_.DisplayName -match 'Online Archive' }
    if (-not $archiveStore) {
        Write-Host 'No archive mailbox found.  Available stores:' -ForegroundColor Yellow
        foreach ($s in $namespace.Stores) {
            Write-Host "  - $($s.DisplayName)  (Type: $($s.ExchangeStoreType))" -ForegroundColor Cyan
        }
        return
    }
    $primaryStore = $namespace.GetDefaultFolder(6).Store   # 6 = olFolderInbox
}

$archiveRoot = $archiveStore.GetRootFolder()
$primaryRoot = $primaryStore.GetRootFolder()

# ── Banner ──
Write-Host ''
Write-Host "Archive : $($archiveStore.DisplayName)" -ForegroundColor Cyan
Write-Host "Primary : $($primaryStore.DisplayName)" -ForegroundColor Green
if ($NewestDate) {
    Write-Host "Filter  : items on or before $($NewestDate.ToString('yyyy-MM-dd'))" -ForegroundColor Yellow
} else {
    Write-Host 'Filter  : none (all items)' -ForegroundColor Yellow
}
if ($DryRun) {
    Write-Host 'Mode    : DRY RUN' -ForegroundColor Magenta
}
Write-Host ([string][char]0x2550 * 60) -ForegroundColor DarkGray

# ── Phase 1: Migrate items ──
Invoke-ArchiveMigration -Folder $archiveRoot -ArchiveRoot $archiveRoot -PrimaryRoot $primaryRoot

# ── Phase 2: Cleanup empty branches ──
if ($CleanupEmptyFolders) {
    Write-Host ''
    Write-Host 'Cleaning up empty folder branches in archive...' -ForegroundColor Cyan
    Write-Host ([string][char]0x2550 * 60) -ForegroundColor DarkGray

    $deletedItemsFolder = $null
    foreach ($f in $archiveRoot.Folders) {
        if ($f.Name -eq 'Deleted Items') { $deletedItemsFolder = $f; break }
    }
    if ($deletedItemsFolder) {
        $script:Stats.CleanedFolders = Remove-EmptyBranches -Folder $archiveRoot -DeletedItemsFolder $deletedItemsFolder
    } else {
        Write-Ts '[WARN] Could not locate Deleted Items in archive — skipping cleanup' -Color Red
    }
    Write-Host ''
}

# ── Summary ──
Write-Host ([string][char]0x2550 * 60) -ForegroundColor DarkGray
if ($DryRun) {
    Write-Host "DRY RUN complete.  Would move $($script:Stats.Moved) item(s) across $($script:Stats.Folders) folder(s)." -ForegroundColor Magenta
    if ($CleanupEmptyFolders) {
        Write-Host "Would remove $($script:Stats.CleanedFolders) empty branch(es)." -ForegroundColor Magenta
    }
} else {
    Write-Host "Done!  Moved $($script:Stats.Moved) item(s) across $($script:Stats.Folders) folder(s)." -ForegroundColor Green
    if ($CleanupEmptyFolders) {
        Write-Host "Removed $($script:Stats.CleanedFolders) empty branch(es)." -ForegroundColor Green
    }
}
if ($script:Stats.Errors -gt 0) {
    Write-Host "Errors: $($script:Stats.Errors)" -ForegroundColor Red
}

# ── Report: skipped item classes ──
if ($script:SkippedItems.Count -gt 0) {
    Write-Host ''
    Write-Host ([string][char]0x2550 * 60) -ForegroundColor DarkGray
    Write-Host 'Skipped items (class not in MoveItemClasses):' -ForegroundColor Yellow
    Write-Host ([string][char]0x2500 * 60) -ForegroundColor DarkGray
    $script:SkippedItems | Group-Object Folder | ForEach-Object {
        Write-Host "  $($_.Name)" -ForegroundColor Cyan
        foreach ($entry in $_.Group) {
            Write-Host "    Class=$($entry.Class) ($($entry.MsgClass))  $($entry.Subject)" -ForegroundColor DarkYellow
        }
    }
    Write-Host ([string][char]0x2500 * 60) -ForegroundColor DarkGray
    Write-Host "Total skipped: $($script:SkippedItems.Count)" -ForegroundColor Yellow
}
