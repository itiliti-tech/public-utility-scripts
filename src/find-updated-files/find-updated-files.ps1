<#
.SYNOPSIS
    Find Updated Files Script - PowerShell Edition

.DESCRIPTION
    Scans a directory tree for the most recently modified file in each subdirectory
    and outputs the results to a CSV file. This is useful for tracking which project
    folders have been updated since a specified date.

    This script uses built-in PowerShell cmdlets only and does not require
    external modules.

.PARAMETER RootDirectory
    The root directory to start scanning from. If not provided, will prompt for input.

.PARAMETER OutputFile
    The output CSV filename. If not provided, will prompt for input.

.PARAMETER DateLimit
    The date threshold for file modification checks (format: M/d/yyyy or MM/dd/yyyy).
    Default: 01/01/2016

.PARAMETER ExcludedFiles
    Array of filenames to exclude from the search.
    Default: Thumbs.db, desktop.ini, .DS_Store, etc.

.EXAMPLE
    .\find-updated-files.ps1 -RootDirectory "C:\Projects" -OutputFile "results.csv"

.EXAMPLE
    .\find-updated-files.ps1  # Prompts for input

.NOTES
    Compatibility: PowerShell 5.0+
    PowerShell 7.x required code: none
    Modified: auto-updated by pre-commit hook
#>

[CmdletBinding()]
param(
    [string]$RootDirectory,
    [string]$OutputFile,
    [string]$DateLimit = "01/01/2016",
    [string[]]$ExcludedFiles = @(
        "Thumbs.db",
        "desktop.ini",
        "Icon`r",
        "System Volume Information",
        ".DS_Store"
    )
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Prompt for input if parameters not provided
if ([string]::IsNullOrWhiteSpace($RootDirectory)) {
    $RootDirectory = Read-Host -Prompt "Input root directory to parse"
}

if ([string]::IsNullOrWhiteSpace($OutputFile)) {
    $OutputFile = Read-Host -Prompt "Input output file name"
}

# Convert date string to datetime object for comparison
try {
    $ModifyDateLimit = [datetime]::ParseExact($DateLimit, "M/d/yyyy", $null)
} catch {
    Write-Error "Invalid date format. Please use M/d/yyyy (e.g., 1/1/2016 or 01/01/2016)"
    exit 1
}

# Validate root directory exists
if (-not (Test-Path -Path $RootDirectory -PathType Container)) {
    Write-Error "Root directory does not exist: $RootDirectory"
    exit 1
}

<#
.FUNCTION
    Search-UpdatedFiles

.DESCRIPTION
    Recursively searches directories for the most recently modified file.

    Scans the root directory and its subdirectories to find files modified
    after the specified date. For each top-level folder, outputs the most
    recently modified file and its modification date.

.PARAMETER RootDirectory
    The root directory to start searching from

.PARAMETER ModifyDateLimit
    The date threshold for modifications (as a datetime object)

.PARAMETER ExcludedFiles
    Array of filenames to skip during the search

.PARAMETER OutputFile
    Path to the output CSV file
#>
function Search-UpdatedFiles {
    param(
        [string]$RootDirectory,
        [datetime]$ModifyDateLimit,
        [string[]]$ExcludedFiles,
        [string]$OutputFile
    )

    # Iterate through all top-level folders separately
    $TopLevelDirs = Get-ChildItem -Path $RootDirectory -Directory -ErrorAction SilentlyContinue

    foreach ($Dir in $TopLevelDirs) {
        $DirPath = $Dir.FullName
        $FileFound = $false
        $FoundFilePath = ""
        $FoundFileModifiedDate = $null

        # Traverse all subdirectories and files using Get-ChildItem recursively
        $AllFiles = Get-ChildItem -Path $DirPath -File -Recurse -ErrorAction SilentlyContinue

        foreach ($File in $AllFiles) {
            # Skip excluded files
            if ($File.Name -in $ExcludedFiles) {
                continue
            }

            # Check if file was modified after the threshold date
            # If so, we found our most recent file for this directory branch
            if ($File.LastWriteTime -gt $ModifyDateLimit) {
                $FileFound = $true
                $FoundFilePath = $File.FullName
                $FoundFileModifiedDate = $File.LastWriteTime
                # Break out of file loop to move to next directory
                break
            }
        }

        # Output the results for this top-level directory
        if ($FileFound) {
            # Write to console and CSV file: folder, file path, and modification date
            $FormattedDate = $FoundFileModifiedDate.ToString("MM/dd/yyyy")
            Write-Host "$DirPath   $FoundFilePath   $FormattedDate"

            # Write to CSV file with proper quoting
            $CsvLine = "`"$DirPath`",`"$FormattedDate`",`"$FoundFilePath`""
            Add-Content -Path $OutputFile -Value $CsvLine -Encoding UTF8
        } else {
            # No files found modified after the threshold date
            $FormattedDateLimit = $ModifyDateLimit.ToString("MM/dd/yyyy")
            Write-Host "$DirPath - no modification found after $FormattedDateLimit"

            # Write to CSV file
            $CsvLine = "`"$DirPath`",`"No modification found`","""""
            Add-Content -Path $OutputFile -Value $CsvLine -Encoding UTF8
        }
    }
}

# =============================================================================
# Main Script Execution
# =============================================================================

try {
    # Clear the output file and write header
    "Folder Name,Last Modified,File" | Set-Content -Path $OutputFile -Encoding UTF8

    # Execute the search
    Search-UpdatedFiles -RootDirectory $RootDirectory `
        -ModifyDateLimit $ModifyDateLimit `
        -ExcludedFiles $ExcludedFiles `
        -OutputFile $OutputFile

    Write-Host "✓ Scan complete. Results written to: $OutputFile"
} catch {
    Write-Error "Error: $_"
    exit 1
}
