#!/bin/bash

# Pre-commit hook to manage .gitkeep files
# - Removes .gitkeep from directories that contain other files
# - Adds .gitkeep to empty directories that should be preserved
# This prevents tracking placeholder files when they're no longer needed,
# while preserving empty directories in the repo structure.

REPO_ROOT="$(git rev-parse --show-toplevel 2>/dev/null)" || {
    echo "Error: Not in a git repository"
    exit 1
}

cd "$REPO_ROOT"

# Directories that should always have .gitkeep if empty
PRESERVED_DIRS=("src" "scripts" "docs" ".config")

# Step 1: Remove unnecessary .gitkeep files from directories with content
while IFS= read -r gitkeep_file; do
    [ -z "$gitkeep_file" ] && continue

    dir="$(dirname "$gitkeep_file")"

    # Skip if directory doesn't exist (was deleted)
    [ ! -d "$dir" ] && continue

    # Count files in directory, excluding .gitkeep
    file_count=$(find "$dir" -maxdepth 1 -type f ! -name '.gitkeep' 2>/dev/null | wc -l)

    # If directory has files besides .gitkeep, remove it
    if [ "$file_count" -gt 0 ]; then
        echo "Removing unnecessary .gitkeep: $gitkeep_file (directory has $file_count files)"
        rm -f "$gitkeep_file"

        # Stage the removal
        git add -u "$gitkeep_file" 2>/dev/null || true
    fi
done < <(find . -name '.gitkeep' -type f)

# Step 2: Add .gitkeep to empty preserved directories
for dir in "${PRESERVED_DIRS[@]}"; do
    # Skip if directory doesn't exist
    [ ! -d "$dir" ] && continue

    # Count files in directory (including hidden files, excluding . and ..)
    file_count=$(find "$dir" -maxdepth 1 -type f 2>/dev/null | wc -l)

    # If directory is empty, add .gitkeep
    if [ "$file_count" -eq 0 ]; then
        gitkeep="$dir/.gitkeep"
        if [ ! -f "$gitkeep" ]; then
            echo "Adding .gitkeep to empty directory: $dir"
            touch "$gitkeep"
            git add "$gitkeep"
        fi
    fi
done

exit 0
