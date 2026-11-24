# Disk Usage & Duplicate Finder (PowerShell)

User-friendly PowerShell script for normal Windows users.  
Scans disk usage and finds true duplicate files (SHA-256 hash).  
Generates a clean offline HTML report that opens in your browser.

Inspired by [nwlterry/disk_space_analysis-windows-ps](https://github.com/nwlterry/disk_space_analysis-windows-ps).

## Features

- Scan any folder, drive, or multiple paths
- Skips folders and locked files that local Administrators or LOCAL SYSTEM cannot read (`System Volume Information`, `WindowsApps`, `Windows\CSC`, WinRE, live registry hives, `pagefile.sys`). One denied folder no longer stops a `C:\` or `D:\` scan.
- Defaults to common user folders (Documents, Downloads, Desktop, Pictures, Videos, Music) if no path is given
- Largest folders & files
- Breakdown by file type (extension) and by age
- True duplicate detection with SHA-256 (only files above a size threshold are hashed for speed)
- Standalone HTML report (dark theme, tabs, progress bars)
- CSV + JSON exports
- Read-only — never deletes or modifies files
- Works as a normal user (no admin required for most personal folders)

## Quick Start

1. Open PowerShell in the folder where you saved the script.
2. Allow the script for this session:

   ```powershell
   Set-ExecutionPolicy -Scope Process Bypass
   ```

3. Run it:

   ```powershell
   .\Get-DiskUsageAndDuplicates.ps1
   ```

### Specifying paths

**Recommended (always works):**

```powershell
# Single path
.\Get-DiskUsageAndDuplicates.ps1 -Path "C:\Users\terry.ng"

# Multiple paths as a single comma-separated string
.\Get-DiskUsageAndDuplicates.ps1 -Path "C:\Users\terry.ng,C:\Backup"

# Multiple paths as an array (works when run inside PowerShell)
.\Get-DiskUsageAndDuplicates.ps1 -Path "C:\Users\terry.ng","C:\Backup"
```

**Note:** If you launch the script with `powershell.exe -File script.ps1 ...`, prefer the single comma-separated string form. The array form works reliably when you are already inside a PowerShell session.

## Parameters

| Parameter       | Default                  | Description                                      |
|-----------------|--------------------------|--------------------------------------------------|
| `-Path`         | User profile folders     | Folder(s) or drive(s) to scan. Accepts a single path, an array, or a comma/semicolon-separated string |
| `-OutDir`       | `.\DiskUsageReport`      | Output folder for HTML / CSV / JSON              |
| `-MinHashMB`    | `5`                      | Minimum file size (MB) to consider for hashing   |
| `-MaxFilesToHash` | `50000`                | Safety limit on number of files to hash          |

## Output

- `disk-usage-report-YYYYMMDD-HHMMSS.html` — main report (opens automatically)
- `disk-usage-report-YYYYMMDD-HHMMSS.json`
- `largest-files-*.csv`
- `top-folders-*.csv`
- `by-extension-*.csv`
- `by-age-*.csv`
- `duplicates-*.csv` (if any duplicates found)

## Notes

- Hashing is the slow part. Raising `-MinHashMB` makes it faster.
- Access-denied folders are skipped and listed in the report errors section.
- Junctions and reparse points are not followed.
- The HTML report is fully offline and self-contained.

## Example

```powershell
.\Get-DiskUsageAndDuplicates.ps1 -Path "C:\Users\terry.ng,C:\Backup" -MinHashMB 2
```
