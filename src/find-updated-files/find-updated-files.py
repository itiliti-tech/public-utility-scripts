#!/usr/bin/env python3
"""
Find Updated Files Script

Scans a directory tree for the most recently modified file in each subdirectory
and outputs the results to a CSV file. This is useful for tracking which project
folders have been updated since a specified date.

Usage:
    python find-updated-files.py
    
    The script will prompt for:
    - Root directory to scan
    - Output CSV filename
    
Output:
    CSV file with columns: Folder Name, Last Modified, File
    
Modified: auto-updated by pre-commit hook
"""

import os
import time
import unicodedata


def searchFiles(rootDirectory, modifyDateLimit, excludedFiles):
    """
    Recursively search directories for the most recently modified file.
    
    Scans the root directory and its subdirectories to find files modified
    after the specified date. For each top-level folder, outputs the most
    recently modified file and its modification date.
    
    Args:
        rootDirectory (str): The root directory to start searching from
        modifyDateLimit (time.struct_time): The date threshold for modifications
        excludedFiles (list): List of filenames to skip during the search
        
    Returns:
        None (outputs results to file and stdout)
    """
    # iterate through all top level folders seperately

    for dir in os.listdir(rootDirectory):
        # get path for next level and make sure path is a directory (not a file)
        fileFound = False
        foundFilePath = ""
        foundFileModifiedDate = ""

        dirPath = os.path.join(rootDirectory, dir)

        if not os.path.isdir(dirPath):
            continue

        # traverse all branches of the path using os.walk
        for root, dirs, files in os.walk(dirPath, topdown=True, followlinks=False):
            # Search through all files in this directory level
            for name in files:
                filePath = os.path.join(root, name)

                # Verify the path points to a file
                if not os.path.isfile(filePath):
                    continue

                # Skip excluded files (system/temporary files)
                if name in excludedFiles:
                    continue

                # get the time the file was modified in seconds (float)
                modifyTime = os.path.getmtime(filePath)

                # convert modifyTime to time.struct_time object for comparison
                modifyDateTime = time.localtime(modifyTime)

                # Check if this file was modified after the given threshold date.
                # If so, we found our most recent file for this directory branch.
                if modifyDateTime > modifyDateLimit:
                    fileFound = True
                    foundFilePath = filePath
                    foundFileModifiedDate = modifyDateTime
                    # Break out of file loop to move to next directory branch
                    break

            # If we found a recent file, stop searching deeper in this branch
            if fileFound:
                # Since we found a file that was modified after the given
                # modify date, we can break out of the os.walk loop
                break

        # Output the results for this top-level directory
        if fileFound:
            # Write to stdout and CSV file: folder, file path, and modification date
            print(
                unicodedata.normalize("NFC", dirPath),
                " ",
                unicodedata.normalize("NFC", foundFilePath),
                " ",
                time.strftime("%m/%d/%Y", foundFileModifiedDate),
            )

            with open(outputFile, "a") as f:
                f.write(
                    '"'
                    + unicodedata.normalize("NFC", dirPath)
                    + '","'
                    + time.strftime("%m/%d/%Y", foundFileModifiedDate)
                    + '","'
                    + unicodedata.normalize("NFC", foundFilePath)
                    + '"\n'
                )

        else:
            # No files found modified after the threshold date
            print(
                unicodedata.normalize("NFC", dirPath),
                " - no modification found after ",
                time.strftime("%m/%d/%Y", modifyDateLimit),
            )

            with open(outputFile, "a", encoding="utf8") as f:
                f.write(
                    '"'
                    + unicodedata.normalize("NFC", dirPath)
                    + '","No modification found",\n'
                )


# ============================================================================
# Main Script Execution
# ============================================================================

# Set the date threshold for file modification checks
dateLimit = "01/01/2016"
modifyDateLimit = time.strptime(dateLimit, "%m/%d/%Y")

# Files to exclude from the search (system and temporary files)
excludedFiles = [
    "Thumbs.db",
    "desktop.ini",
    "Icon\r",
    "$RECYCLE.BIN",
    "System Volume Information",
    ".DS_Store",
]

# take user input for the root directory and output file name
directory = input("Input root directory to parse:\n")
outputFile = input("Input output file name:\n")


# write header to output file
with open(outputFile, "w", encoding="utf8") as f:
    f.write("Folder Name,Last Modified,File\n")

try:
    searchFiles(directory, modifyDateLimit, excludedFiles)

except OSError as e:
    print(f"Error: {e}")
