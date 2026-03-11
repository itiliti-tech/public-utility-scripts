# Outlook COM & Exchange Online Archive — Operational Notes

> Hard-won knowledge from automating mailbox operations with the Outlook COM
> object model against Exchange Online Archive stores. Keep this file as a
> reference for any future scripts that manipulate folders or items via COM.

---

## 1. Connecting to the Online Archive

```powershell
$outlook   = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace('MAPI')

# The archive shows up as a separate store whose DisplayName contains "Online Archive"
$archiveStore = $namespace.Stores | Where-Object { $_.DisplayName -match 'Online Archive' }
$archiveRoot  = $archiveStore.GetRootFolder()

# The primary mailbox is the store that owns the default Inbox
$primaryStore = $namespace.GetDefaultFolder(6).Store   # 6 = olFolderInbox
```

- The archive store is **only visible** when the Outlook profile has the archive
  mailbox configured (Exchange Online plans that include archiving).
- Cached Exchange Mode must have the archive enabled; otherwise the store may
  not appear in `$namespace.Stores`.

---

## 2. Folder Deletion Quirks (Online Archive)

### 2.1 Leaf-folder `.Delete()` silently fails

Calling `.Delete()` on a **leaf** (bottom-most) folder in an Online Archive
store does **not** throw an error — it simply does nothing. The folder remains
in place and never appears in Deleted Items.

`.MoveTo($deletedItems)` exhibits the same silent-failure behaviour on leaf
folders.

**Workaround — top-down branch deletion:**  
Delete from the _highest_ ancestor whose entire branch is empty. Exchange does
successfully delete a parent folder along with all its (empty) children in a
single `.Delete()` call.

```text
Archive\
  ProjectA\          ← items=0
    2023\            ← items=0   (leaf — Delete() would silently fail)
    2024\            ← items=0   (leaf — Delete() would silently fail)

# Deleting "ProjectA" instead WORKS and takes both child folders with it.
```

### 2.2 Duplicate-name collision in Deleted Items

If a folder with the **same name** already exists inside Deleted Items,
`.Delete()` on the archive folder silently fails — again with no error.

**Workaround — rename before delete:**  
Give the folder a unique name before calling `.Delete()` so there can be no
collision:

```powershell
$folder.Name = "${originalName}_cleanup_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
$folder.Delete()
```

### 2.3 Post-delete verification

Because deletions can silently fail, always verify the outcome by checking
whether the folder appeared in Deleted Items:

```powershell
$folder.Delete()
Start-Sleep -Milliseconds 500   # brief settle time

$found = $false
foreach ($df in $deletedItemsFolder.Folders) {
    if ($df.Name -eq $expectedName) { $found = $true; break }
}
```

---

## 3. COM Object Lifetime & Stale References

### 3.1 References go stale after throttle / sleep

When Exchange Online throttles a request and the script sleeps before retrying,
the original COM reference to an item (or folder) may become invalid.
Re-acquire the object from the collection before retrying:

```powershell
# After sleep, re-acquire the item by index
try {
    $item = $sourceFolder.Items.Item($index)
} catch {
    Write-Host "Could not re-acquire item $index after sleep"
}
```

### 3.2 Always release COM objects

Unreleased COM objects can cause Outlook to hang on exit or leak memory.
Use a helper to keep the code clean:

```powershell
function Release-ComObject {
    param($Object)
    if ($Object) {
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object) | Out-Null } catch {}
    }
}
```

---

## 4. Moving Items — Shrinking-Index Pattern

When you `.Move()` an item out of a folder, the item **disappears** from the
COM `Items` collection and every subsequent item shifts down by one. Reading
items with a simple `for ($i = 1; $i -le $count; $i++)` will skip every other
item.

**Correct pattern:**

```powershell
$index     = 1
$remaining = $folder.Items.Count

while ($index -le $remaining) {
    $item = $folder.Items.Item($index)

    if (<should skip>) {
        $index++          # item stays → advance past it
        continue
    }

    $item.Move($target)
    # item vanished → DON'T increment $index
    $remaining = $folder.Items.Count   # refresh count
}
```

---

## 5. Exchange Online Throttling & Adaptive Backoff

Exchange Online enforces request-rate limits. Symptoms: `Move()` or
`Delete()` throws a transient COM exception.

**Strategy — exponential backoff with decay on success:**

| Event      | Action                                        |
| ---------- | --------------------------------------------- |
| Move fails | `delay = max(1000, delay * 2)` capped at 30 s |
| Move works | `delay = max(baseDelay, delay * 0.75)`        |

This ramps up quickly when Exchange pushes back and eases off as the server
recovers.

---

## 6. Outlook Item Class Reference

| Class | Constant                      | Description          |
| ----: | ----------------------------- | -------------------- |
|    43 | `olMailItem`                  | Standard email       |
|    46 | `olPostItem`                  | Post item            |
|    53 | `olMeetingRequest`            | Meeting request      |
|    54 | `olMeetingCancellation`       | Meeting cancellation |
|    55 | `olMeetingResponseNegative`   | Decline              |
|    56 | `olMeetingResponsePositive`   | Accept               |
|    57 | `olMeetingResponseTentative`  | Tentative            |
|   104 | `olSharingItem` (IPM.Sharing) | Sharing invitation   |

Use `$item.Class` (integer) to filter; `$item.MessageClass` (string) gives the
IPM-based name.

---

## 7. Folder Navigation Helpers

| Task                   | Approach                                                              |
| ---------------------- | --------------------------------------------------------------------- |
| Check if a path exists | Walk `.Folders.Item($name)` in a try/catch; return `$null` on failure |
| Create path lazily     | Only create when the first eligible item needs the target folder      |
| Get relative path      | Walk `.Parent` up to the root, collecting names into a list           |
| Snapshot children      | `@($folder.Folders)` before iterating if you'll modify the collection |

---

## 8. Miscellaneous Tips

- **Folder counts can be stale** in Cached Exchange Mode. After moving or
  deleting, re-read `.Items.Count` or `.Folders.Count` rather than relying on
  a cached value.
- **`$namespace.GetDefaultFolder()`** enums:  
  6 = Inbox, 3 = Deleted Items, 4 = Outbox, 5 = Sent Items, 9 = Calendar,
  10 = Contacts, 13 = Journal, 12 = Tasks.
- **Store type**: `$store.ExchangeStoreType` distinguishes primary mailboxes
  (1 = `olExchangeMailbox`) from archives.
- **Test in Dry-Run first** — always provide a dry-run flag that logs what
  _would_ happen so you can verify folder paths and item counts before
  committing real moves/deletes.
