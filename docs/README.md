# Documentation Index

## Git Hooks

### Pre-commit Hook

The repository includes a pre-commit hook that automatically manages `.gitkeep` files.

#### Installation

The pre-commit hook is located at `scripts/pre-commit-cleanup-gitkeep.sh`. To install it:

**Bash:**

```bash
# Navigate to repository root
cd /path/to/repository

# Create symbolic link to the hook
ln -s ../../scripts/pre-commit-cleanup-gitkeep.sh .git/hooks/pre-commit

# Make it executable
chmod +x .git/hooks/pre-commit
```

**PowerShell:**

```powershell
# Navigate to repository root
cd C:\path\to\repository

# Create symbolic link to the hook
New-Item -ItemType SymbolicLink -Path .git\hooks\pre-commit -Target ..\..\scripts\pre-commit-cleanup-gitkeep.sh -Force
```

#### What it does

- **Removes** unnecessary `.gitkeep` files from directories that contain actual files
- **Adds** `.gitkeep` to empty directories that should be preserved in the repository structure

This keeps the repository clean and ensures empty directory structure is maintained through version control.

#### How it works

The hook runs automatically before each commit. If it makes changes:

1. It removes `.gitkeep` from directories with files
2. It adds `.gitkeep` to empty preserved directories (src, scripts, docs, .config)
3. It stages these changes so they're included in your commit

You can also run it manually:

```bash
bash scripts/pre-commit-cleanup-gitkeep.sh
```
