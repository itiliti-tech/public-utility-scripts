#!/bin/bash

# Pre-commit hook to manage .gitkeep files and update Last Modified dates
# - Removes .gitkeep from directories that contain other files
# - Adds .gitkeep to empty directories that should be preserved
# - Updates "Last Modified:" timestamps in staged scripts
# This prevents tracking placeholder files when they're no longer needed,
# while preserving empty directories in the repo structure.

REPO_ROOT="$(git rev-parse --show-toplevel 2>/dev/null)" || {
    echo "Error: Not in a git repository"
    exit 1
}

# Function to update Last Modified date in scripts
# This function updates the timestamp regardless of whether the changes
# are comment-only or include code modifications
update_last_modified() {
    local file="$1"
    local current_date
    current_date=$(date '+%Y-%m-%d %H:%M')

    # Check if file contains "Last Modified:" pattern in the first 30 lines (header area)
    if head -n 30 "$file" | grep -q "Last Modified:"; then
        echo "Updating Last Modified date in: $file"

        # Create temporary file for in-place editing
        local temp_file="${file}.tmp"

        # Update the date (works for both # and <!-- --> style comments)
        sed -E "s/(Last Modified:)[[:space:]]*[0-9]{4}-[0-9]{2}-[0-9]{2}[[:space:]]+[0-9]{2}:[0-9]{2}/\1 ${current_date}/g" "$file" > "$temp_file"

        # Check if file was actually modified
        if ! cmp -s "$file" "$temp_file"; then
            mv "$temp_file" "$file"
            git add "$file"
            return 0
        else
            rm -f "$temp_file"
            return 1
        fi
    fi
    return 1
}

cd "$REPO_ROOT"

# Step 0: Update Last Modified dates in staged script files
while IFS= read -r file; do
    [ -z "$file" ] && continue
    [ ! -f "$file" ] && continue

    # Only process script files (common extensions)
    case "$file" in
        *.ps1|*.sh|*.py|*.js|*.ts|*.bash|*.zsh)
            update_last_modified "$file"
            ;;
    esac
done < <(git diff --cached --name-only --diff-filter=ACM)

# Directories that should always have .gitkeep if empty
PRESERVED_DIRS=("src" "scripts" "docs" ".config", ".github")

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
