# Find Updated Files

## Overview

This folder contains PowerShell and Python implementations for scanning a root directory and reporting recently modified files by top-level folder.

Output is a CSV with columns:

- Folder Name
- Last Modified
- File

## Files

- `find-updated-files.ps1`: Parameterized PowerShell version.
- `find-updated-files.py`: Interactive Python version.
- `requirements.txt`: Notes that no external Python dependencies are required.

## PowerShell Usage

Run with explicit parameters:

```powershell
.\src\find-updated-files\find-updated-files.ps1 -RootDirectory "C:\Projects" -OutputFile ".\results.csv" -DateLimit "01/01/2024"
```

Run interactively (prompts for root and output):

```powershell
.\src\find-updated-files\find-updated-files.ps1
```

Optional parameter:

- `-ExcludedFiles` to override the default excluded filenames.

## Python Usage

```powershell
python .\src\find-updated-files\find-updated-files.py
```

The script prompts for:

- Root directory to scan
- Output CSV file name

## Notes

- Default date threshold is `01/01/2016`.
- The script reports the first qualifying recent file found while traversing each top-level folder.
- Use the PowerShell version when you want automation-friendly parameters.
