<#
.SYNOPSIS
    User-friendly Disk Usage & Duplicate Finder for Windows.

.DESCRIPTION
    Scans a folder or drive, reports disk usage (largest folders/files, by type, by age),
    and finds true duplicate files using SHA-256 hashes.
    Generates CSV + JSON + a standalone HTML report that opens in the browser.
    Designed for normal (non-admin) users. No external tools required.

.PARAMETER Path
    Folder or drive to scan. Accepts a single path, multiple paths, or comma/semicolon-separated string.
    Defaults to the current user's profile folders
    (Documents, Downloads, Desktop, Pictures, Videos, Music) if not specified.

.PARAMETER OutDir
    Where to write reports. Defaults to .\DiskUsageReport in the current directory.

.PARAMETER MinHashMB
    Only consider files this size or larger for duplicate hashing (default 5 MB).
    Smaller files are still counted for usage stats.

.PARAMETER MaxFilesToHash
    Safety limit on how many candidate files to hash (default 50000).

.EXAMPLE
    .\Get-DiskUsageAndDuplicates.ps1
    # Interactive / defaults to user profile folders

.EXAMPLE
    .\Get-DiskUsageAndDuplicates.ps1 -Path "D:\Photos","E:\Videos" -OutDir "C:\Reports" -MinHashMB 1

.NOTES
    Inspired by https://github.com/nwlterry/disk_space_analysis-windows-ps
    Safe: read-only. Does not delete or modify any files.
#>

[CmdletBinding()]
param(
    # Accept any form: single string, array, or comma-separated. Normalized inside.
    [Parameter()]
    [object]$Path = $null,
    [string]$OutDir = ".\DiskUsageReport",
    [int]$MinHashMB = 5,
    [int]$MaxFilesToHash = 50000
)

# ---------- Helpers ----------
function Write-Status {
    param([string]$Message, [string]$Color = "Cyan")
    $ts = Get-Date -Format "HH:mm:ss"
    Write-Host "[$ts] $Message" -ForegroundColor $Color
}

function Format-Bytes {
    param([long]$Bytes)
    if ($Bytes -ge 1TB) { return "{0:N2} TB" -f ($Bytes / 1TB) }
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    return "$Bytes B"
}

function Get-AgeBucket {
    param([datetime]$LastWrite)
    $days = ((Get-Date) - $LastWrite).TotalDays
    if ($days -le 7)   { return "0-7 days" }
    if ($days -le 30)  { return "8-30 days" }
    if ($days -le 90)  { return "31-90 days" }
    if ($days -le 365) { return "91-365 days" }
    return "Older than 1 year"
}

function Normalize-Path([string]$p) {
    if ([string]::IsNullOrEmpty($p)) { return $p }
    return $p.Replace('\','/')
}

# ---------- Resolve scan roots ----------
$roots = @()

# Normalize Path argument into a clean string array
# Handles: null, single string, array, or a string that looks like "a","b" or a,b
$normalized = @()
if ($null -ne $Path) {
    if ($Path -is [string]) {
        $normalized = @($Path)
    } elseif ($Path -is [System.Collections.IEnumerable]) {
        $normalized = @($Path | ForEach-Object { "$_" })
    } else {
        $normalized = @("$Path")
    }
}

# Flatten further: if any element itself contains comma/semicolon or looks like quoted list, split it
$final = @()
foreach ($item in $normalized) {
    if ([string]::IsNullOrWhiteSpace($item)) { continue }
    # Remove surrounding quotes if the whole item is quoted
    $clean = $item.Trim()
    if ($clean.StartsWith('"') -and $clean.EndsWith('"')) { $clean = $clean.Trim('"') }
    if ($clean.StartsWith("'") -and $clean.EndsWith("'")) { $clean = $clean.Trim("'") }
    # Split on comma or semicolon
    $parts = $clean -split '[,;]' | ForEach-Object { $_.Trim().Trim('"').Trim("'") } | Where-Object { $_ }
    $final += $parts
}
$normalized = $final

if ($normalized.Count -eq 0 -or ($normalized.Count -eq 1 -and [string]::IsNullOrWhiteSpace($normalized[0]))) {
    Write-Status "No path specified. Scanning common user folders..." "Yellow"
    $userProfile = $env:USERPROFILE
    $candidates = @(
        "$userProfile\Documents",
        "$userProfile\Downloads",
        "$userProfile\Desktop",
        "$userProfile\Pictures",
        "$userProfile\Videos",
        "$userProfile\Music"
    )
    foreach ($c in $candidates) {
        if (Test-Path -LiteralPath $c) { $roots += (Resolve-Path -LiteralPath $c).Path }
    }
    if ($roots.Count -eq 0) {
        Write-Status "No standard user folders found. Please specify -Path." "Red"
        exit 1
    }
} else {
    $raw = @()
    foreach ($p in $normalized) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        $raw += $p -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    }
    foreach ($r in $raw) {
        if (Test-Path -LiteralPath $r) {
            $roots += (Resolve-Path -LiteralPath $r).Path
        } else {
            Write-Status "Path not found, skipping: $r" "Yellow"
        }
    }
    if ($roots.Count -eq 0) {
        Write-Status "None of the specified paths exist." "Red"
        exit 1
    }
}

Write-Status "Scan roots:" "Green"
$roots | ForEach-Object { Write-Host "  $_" }

# ---------- Prepare output ----------
$OutDir = [System.IO.Path]::GetFullPath($OutDir)
if (-not (Test-Path -LiteralPath $OutDir)) {
    New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
}
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
Write-Status "Reports will be written to: $OutDir" "Green"

# ---------- Data structures ----------
$allFiles = New-Object System.Collections.Generic.List[object]
$folderSizes = @{}          # path -> total size of contents (recursive)
$extensionStats = @{}       # ext -> @{ Size; Count }
$ageStats = @{}             # bucket -> @{ Size; Count }
$errors = New-Object System.Collections.Generic.List[object]

$minHashBytes = [long]$MinHashMB * 1MB
$sizeGroups = @{}           # size -> list of paths (for duplicate candidates)

Write-Status "Starting recursive scan..." "Cyan"
$sw = [System.Diagnostics.Stopwatch]::StartNew()
$fileCount = 0
$dirCount = 0

foreach ($root in $roots) {
    Write-Status "Scanning: $root" "Yellow"
    try {
        $enum = [System.IO.Directory]::EnumerateFiles($root, "*", [System.IO.SearchOption]::AllDirectories)
        foreach ($filePath in $enum) {
            $fileCount++
            if ($fileCount % 2000 -eq 0) {
                Write-Host "`r  Files found: $fileCount" -NoNewline -ForegroundColor DarkGray
            }

            try {
                $fi = [System.IO.FileInfo]::new($filePath)
                # Skip reparse points / junctions
                if ($fi.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { continue }

                $size = $fi.Length
                $ext = if ([string]::IsNullOrEmpty($fi.Extension)) { "(no extension)" } else { $fi.Extension.ToLower() }
                $age = Get-AgeBucket $fi.LastWriteTime

                $allFiles.Add([PSCustomObject]@{
                    Path        = $filePath
                    Name        = $fi.Name
                    Size        = $size
                    Extension   = $ext
                    LastWrite   = $fi.LastWriteTime
                    AgeBucket   = $age
                    Directory   = $fi.DirectoryName
                })

                # Extension stats
                if (-not $extensionStats.ContainsKey($ext)) {
                    $extensionStats[$ext] = @{ Size = 0L; Count = 0 }
                }
                $extensionStats[$ext].Size += $size
                $extensionStats[$ext].Count++

                # Age stats
                if (-not $ageStats.ContainsKey($age)) {
                    $ageStats[$age] = @{ Size = 0L; Count = 0 }
                }
                $ageStats[$age].Size += $size
                $ageStats[$age].Count++

                # Folder size accumulation (walk up)
                $dir = $fi.DirectoryName
                while ($dir -and $dir.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
                    if (-not $folderSizes.ContainsKey($dir)) { $folderSizes[$dir] = 0L }
                    $folderSizes[$dir] += $size
                    $parent = [System.IO.Path]::GetDirectoryName($dir)
                    if ($parent -eq $dir) { break }
                    $dir = $parent
                }

                # Candidate for duplicates
                if ($size -ge $minHashBytes) {
                    if (-not $sizeGroups.ContainsKey($size)) {
                        $sizeGroups[$size] = New-Object System.Collections.Generic.List[string]
                    }
                    $sizeGroups[$size].Add($filePath)
                }
            } catch {
                $errors.Add([PSCustomObject]@{ Path = $filePath; Error = $_.Exception.Message })
            }
        }
    } catch {
        $errors.Add([PSCustomObject]@{ Path = $root; Error = $_.Exception.Message })
        Write-Status "Access issue on $root : $($_.Exception.Message)" "Yellow"
    }
}

Write-Host ""
$sw.Stop()
Write-Status "Scan finished. Files: $fileCount   Time: $($sw.Elapsed.ToString('hh\:mm\:ss'))" "Green"

# ---------- Largest files & folders ----------
$largestFiles = $allFiles | Sort-Object Size -Descending | Select-Object -First 100
$topFolders = $folderSizes.GetEnumerator() |
    Where-Object { $_.Key -notin $roots } |  # skip the roots themselves if you prefer
    Sort-Object Value -Descending |
    Select-Object -First 50 |
    ForEach-Object {
        [PSCustomObject]@{
            Path = (Normalize-Path $_.Key)
            Size = $_.Value
            SizeFormatted = Format-Bytes $_.Value
        }
    }

# ---------- Duplicate detection (hash only size-matched candidates) ----------
Write-Status "Looking for duplicates (SHA-256, files >= $MinHashMB MB)..." "Cyan"
$duplicateGroups = New-Object System.Collections.Generic.List[object]
$hashedCount = 0

foreach ($pair in $sizeGroups.GetEnumerator()) {
    $paths = $pair.Value
    if ($paths.Count -lt 2) { continue }

    # Hash the group
    $byHash = @{}
    foreach ($p in $paths) {
        if ($hashedCount -ge $MaxFilesToHash) {
            Write-Status "Reached MaxFilesToHash limit ($MaxFilesToHash). Stopping further hashing." "Yellow"
            break
        }
        try {
            $hash = (Get-FileHash -LiteralPath $p -Algorithm SHA256 -ErrorAction Stop).Hash
            $hashedCount++
            if (-not $byHash.ContainsKey($hash)) { $byHash[$hash] = New-Object System.Collections.Generic.List[string] }
            $byHash[$hash].Add($p)
        } catch {
            $errors.Add([PSCustomObject]@{ Path = $p; Error = "Hash failed: $($_.Exception.Message)" })
        }
    }
    if ($hashedCount -ge $MaxFilesToHash) { break }

    foreach ($h in $byHash.Keys) {
        $group = $byHash[$h]
        if ($group.Count -ge 2) {
            $size = $pair.Key
            $waste = $size * ($group.Count - 1)
            $duplicateGroups.Add([PSCustomObject]@{
                Hash          = $h
                Size          = $size
                SizeFormatted = Format-Bytes $size
                Count         = $group.Count
                Waste         = $waste
                WasteFormatted= Format-Bytes $waste
                Paths         = $group
            })
        }
    }
}

$duplicateGroups = $duplicateGroups | Sort-Object Waste -Descending
$totalWaste = ($duplicateGroups | Measure-Object Waste -Sum).Sum
if (-not $totalWaste) { $totalWaste = 0 }

Write-Status "Duplicate groups found: $($duplicateGroups.Count)   Potential space savings: $(Format-Bytes $totalWaste)" "Green"

# ---------- Extension & Age tables ----------
$extensionTable = $extensionStats.GetEnumerator() |
    Sort-Object { $_.Value.Size } -Descending |
    Select-Object -First 40 |
    ForEach-Object {
        [PSCustomObject]@{
            Extension = $_.Key
            Size      = $_.Value.Size
            SizeFormatted = Format-Bytes $_.Value.Size
            Count     = $_.Value.Count
        }
    }

$ageOrder = @("0-7 days","8-30 days","31-90 days","91-365 days","Older than 1 year")
$ageTable = foreach ($bucket in $ageOrder) {
    if ($ageStats.ContainsKey($bucket)) {
        [PSCustomObject]@{
            AgeBucket     = $bucket
            Size          = $ageStats[$bucket].Size
            SizeFormatted = Format-Bytes $ageStats[$bucket].Size
            Count         = $ageStats[$bucket].Count
        }
    }
}

# ---------- Totals ----------
$totalSize = ($allFiles | Measure-Object Size -Sum).Sum
if (-not $totalSize) { $totalSize = 0 }

# ---------- Export CSV ----------
$largestFiles | Select-Object Path, Name, @{N='SizeMB';E={[math]::Round($_.Size/1MB,2)}}, Extension, LastWrite |
    Export-Csv -Path "$OutDir\largest-files-$stamp.csv" -NoTypeInformation -Encoding UTF8

$topFolders | Export-Csv -Path "$OutDir\top-folders-$stamp.csv" -NoTypeInformation -Encoding UTF8
$extensionTable | Export-Csv -Path "$OutDir\by-extension-$stamp.csv" -NoTypeInformation -Encoding UTF8
$ageTable | Export-Csv -Path "$OutDir\by-age-$stamp.csv" -NoTypeInformation -Encoding UTF8

$dupExport = foreach ($g in $duplicateGroups) {
    foreach ($p in $g.Paths) {
        [PSCustomObject]@{
            Hash = $g.Hash
            SizeMB = [math]::Round($g.Size/1MB,2)
            Count = $g.Count
            WasteMB = [math]::Round($g.Waste/1MB,2)
            Path = $p
        }
    }
}
if ($dupExport) {
    $dupExport | Export-Csv -Path "$OutDir\duplicates-$stamp.csv" -NoTypeInformation -Encoding UTF8
}

# ---------- Build JSON for HTML ----------
# Force arrays even when empty so JS never sees null
function Ensure-Array($v) {
    if ($null -eq $v) { return @() }
    if ($v -is [array]) { return $v }
    return @($v)
}

$report = [ordered]@{
    generated     = (Get-Date).ToString("o")
    roots         = Ensure-Array ($roots | ForEach-Object { Normalize-Path $_ })
    totalSize     = $totalSize
    totalSizeFormatted = Format-Bytes $totalSize
    fileCount     = $allFiles.Count
    duplicateGroups = $duplicateGroups.Count
    totalWaste    = $totalWaste
    totalWasteFormatted = Format-Bytes $totalWaste
    largestFiles  = Ensure-Array ($largestFiles | Select-Object -First 50 | ForEach-Object {
        [ordered]@{ path=(Normalize-Path $_.Path); name=$_.Name; size=$_.Size; sizeFormatted=Format-Bytes $_.Size; ext=$_.Extension; lastWrite=$_.LastWrite.ToString("yyyy-MM-dd HH:mm") }
    })
    topFolders    = Ensure-Array ($topFolders | Select-Object -First 30)
    extensions    = Ensure-Array $extensionTable
    ages          = Ensure-Array $ageTable
    duplicates    = Ensure-Array ($duplicateGroups | Select-Object -First 100 | ForEach-Object {
        [ordered]@{
            hash = $_.Hash
            size = $_.Size
            sizeFormatted = $_.SizeFormatted
            count = $_.Count
            waste = $_.Waste
            wasteFormatted = $_.WasteFormatted
            paths = Ensure-Array ($_.Paths | ForEach-Object { Normalize-Path $_ })
        }
    })
    errors        = Ensure-Array ($errors | Select-Object -First 50)
}

$jsonPath = "$OutDir\disk-usage-report-$stamp.json"
$report | ConvertTo-Json -Depth 6 | Set-Content -Path $jsonPath -Encoding UTF8

# ---------- Generate HTML report ----------
$htmlPath = "$OutDir\disk-usage-report-$stamp.html"
$jsonRaw = $report | ConvertTo-Json -Depth 6 -Compress
# Base64 encode so the data can never break HTML/JS parsing
$b64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($jsonRaw))

$html = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<title>Disk Usage & Duplicate Report</title>
<style>
  :root { --bg:#0f172a; --card:#1e293b; --text:#e2e8f0; --muted:#94a3b8; --accent:#38bdf8; --good:#4ade80; --warn:#fbbf24; --border:#334155; }
  body { margin:0; font-family:Segoe UI, system-ui, sans-serif; background:var(--bg); color:var(--text); }
  header { padding:1.5rem 2rem; background:#020617; border-bottom:1px solid var(--border); }
  h1 { margin:0 0 .25rem 0; font-size:1.6rem; }
  .subtitle { color:var(--muted); font-size:.95rem; }
  .stats { display:grid; grid-template-columns:repeat(auto-fit,minmax(180px,1fr)); gap:1rem; padding:1.5rem 2rem; }
  .stat { background:var(--card); padding:1rem; border-radius:10px; border:1px solid var(--border); }
  .stat .label { color:var(--muted); font-size:.8rem; text-transform:uppercase; letter-spacing:.05em; }
  .stat .value { font-size:1.4rem; font-weight:600; margin-top:.3rem; }
  .stat .value.good { color:var(--good); }
  .tabs { display:flex; gap:.5rem; padding:0 2rem; border-bottom:1px solid var(--border); }
  .tab { background:none; border:none; color:var(--muted); padding:.75rem 1rem; cursor:pointer; font-size:.95rem; border-bottom:2px solid transparent; }
  .tab.active { color:var(--accent); border-bottom-color:var(--accent); }
  .pane { padding:1.5rem 2rem; display:none; }
  .pane.active { display:block; }
  table { width:100%; border-collapse:collapse; font-size:.9rem; }
  th, td { text-align:left; padding:.55rem .6rem; border-bottom:1px solid var(--border); }
  th { color:var(--muted); font-weight:500; position:sticky; top:0; background:var(--bg); }
  tr:hover { background:#1e293b88; }
  .bar { height:8px; background:#334155; border-radius:4px; overflow:hidden; margin-top:4px; }
  .bar > div { height:100%; background:var(--accent); }
  .path { font-family:Consolas, monospace; font-size:.8rem; word-break:break-all; color:#cbd5e1; }
  .waste { color:var(--warn); font-weight:600; }
  footer { padding:1rem 2rem; color:var(--muted); font-size:.8rem; border-top:1px solid var(--border); }
  .dup-group { background:var(--card); border:1px solid var(--border); border-radius:8px; margin-bottom:1rem; padding:1rem; }
  .dup-group h3 { margin:0 0 .5rem 0; font-size:1rem; }
</style>
</head>
<body>
<header>
  <h1>Disk Usage & Duplicate Report</h1>
  <div class="subtitle">Generated <span id="gen"></span> &nbsp;&middot;&nbsp; Scanned: <span id="roots"></span></div>
</header>
<div class="stats" id="stats"></div>
<div class="tabs">
  <button class="tab active" data-pane="overview">Overview</button>
  <button class="tab" data-pane="folders">Top Folders</button>
  <button class="tab" data-pane="files">Largest Files</button>
  <button class="tab" data-pane="types">By Type</button>
  <button class="tab" data-pane="ages">By Age</button>
  <button class="tab" data-pane="dupes">Duplicates</button>
</div>
<div id="overview" class="pane active"></div>
<div id="folders" class="pane"></div>
<div id="files" class="pane"></div>
<div id="types" class="pane"></div>
<div id="ages" class="pane"></div>
<div id="dupes" class="pane"></div>
<footer>Read-only report. No files were modified. Inspired by nwlterry/disk_space_analysis-windows-ps</footer>
<script>
const data = JSON.parse(atob('__B64_PLACEHOLDER__'));
// Defensive defaults
data.roots = data.roots || [];
data.largestFiles = data.largestFiles || [];
data.topFolders = data.topFolders || [];
data.extensions = data.extensions || [];
data.ages = data.ages || [];
data.duplicates = data.duplicates || [];
data.errors = data.errors || [];

document.getElementById('gen').textContent = data.generated ? new Date(data.generated).toLocaleString() : '';
document.getElementById('roots').textContent = (data.roots || []).join(', ');

function fmt(n){ return n.toLocaleString(); }
function pct(part, total){ return total ? Math.round(part/total*100) : 0; }

document.getElementById('stats').innerHTML =
  '<div class="stat"><div class="label">Total Size</div><div class="value">' + (data.totalSizeFormatted||'') + '</div></div>' +
  '<div class="stat"><div class="label">Files</div><div class="value">' + fmt(data.fileCount||0) + '</div></div>' +
  '<div class="stat"><div class="label">Duplicate Groups</div><div class="value">' + fmt(data.duplicateGroups||0) + '</div></div>' +
  '<div class="stat"><div class="label">Potential Savings</div><div class="value good">' + (data.totalWasteFormatted||'') + '</div></div>';

// Overview
document.getElementById('overview').innerHTML = ``
  <p>This report shows what is using space and where true duplicates (identical SHA-256 content) exist.</p>
  <p>Only files ≥ the configured threshold were hashed. Smaller files are still included in usage totals.</p>
``;

// Folders (tree based on scanned roots)
function buildTree(folders, roots) {
  var nodes = {};
  (folders || []).forEach(function(f) {
    var p = (f.Path || f.path || '').replace(/\\/g, '/');
    if (!p) return;
    var parts = p.split('/').filter(Boolean);
    var cur = '';
    parts.forEach(function(part, i) {
      var next = cur + '/' + part;
      if (!nodes[next]) {
        nodes[next] = { name: part, path: next, children: {}, size: '', isLeaf: false };
      }
      if (i === parts.length - 1) {
        nodes[next].size = f.SizeFormatted || f.sizeFormatted || '';
        nodes[next].isLeaf = true;
      }
      if (cur && nodes[cur]) nodes[cur].children[next] = nodes[next];
      cur = next;
    });
  });
  var rootNodes = [];
  var seen = {};
  (roots || []).forEach(function(r) {
    var rn = r.replace(/\\/g, '/');
    if (nodes[rn]) { rootNodes.push(nodes[rn]); seen[rn] = true; }
  });
  if (rootNodes.length === 0) {
    Object.keys(nodes).forEach(function(k) {
      var depth = k.split('/').filter(Boolean).length;
      if (depth === 1 || depth === 2) rootNodes.push(nodes[k]);
    });
  }
  return rootNodes;
}
function renderTree(nodes, depth) {
  var html = '';
  (nodes || []).forEach(function(n) {
    var indent = '&nbsp;'.repeat(depth * 4);
    var kids = Object.values(n.children || {});
    html += '<div style="padding:2px 0;font-family:Consolas,monospace;font-size:.85rem">' +
            indent + (kids.length ? '📁 ' : '📄 ') + n.name +
            (n.size ? ' <span style="color:#94a3b8">(' + n.size + ')</span>' : '') +
            '</div>';
    if (kids.length) html += renderTree(kids, depth + 1);
  });
  return html;
}
var treeRoots = buildTree(data.topFolders, data.roots);
var fhtml = '<div style="padding:.5rem 0">' + (treeRoots.length ? renderTree(treeRoots, 0) : '<p>No folder data.</p>') + '</div>';
document.getElementById('folders').innerHTML = fhtml;

// Files
let flhtml = '<table><thead><tr><th>File</th><th>Size</th><th>Modified</th></tr></thead><tbody>';
(data.largestFiles || []).forEach(f => {
  var p = f.path || f.Path || '';
  var s = f.sizeFormatted || f.SizeFormatted || '';
  var d = f.lastWrite || f.LastWrite || '';
  flhtml += '<tr><td class="path">' + p + '</td><td>' + s + '</td><td>' + d + '</td></tr>';
});
flhtml += '</tbody></table>';
document.getElementById('files').innerHTML = flhtml;

// Types
let thtml = '<table><thead><tr><th>Extension</th><th>Size</th><th>Count</th><th></th></tr></thead><tbody>';
const maxExt = Math.max(...(data.extensions || []).map(e => e.Size || e.size || 0), 1);
(data.extensions || []).forEach(e => {
  var p = pct(e.Size || e.size || 0, maxExt);
  var ext = e.Extension || e.extension || '';
  var sz = e.SizeFormatted || e.sizeFormatted || '';
  var cnt = fmt(e.Count || e.count || 0);
  thtml += '<tr><td>' + ext + '</td><td>' + sz + '</td><td>' + cnt + '</td><td><div class="bar"><div style="width:' + p + '%"></div></div></td></tr>';
});
thtml += '</tbody></table>';
document.getElementById('types').innerHTML = thtml;

// Ages
let ahtml = '<table><thead><tr><th>Age</th><th>Size</th><th>Files</th><th></th></tr></thead><tbody>';
const maxAge = Math.max(...(data.ages || []).map(a => a.Size || a.size || 0), 1);
(data.ages || []).forEach(a => {
  var p = pct(a.Size || a.size || 0, maxAge);
  var bucket = a.AgeBucket || a.ageBucket || '';
  var sz = a.SizeFormatted || a.sizeFormatted || '';
  var cnt = fmt(a.Count || a.count || 0);
  ahtml += '<tr><td>' + bucket + '</td><td>' + sz + '</td><td>' + cnt + '</td><td><div class="bar"><div style="width:' + p + '%"></div></div></td></tr>';
});
ahtml += '</tbody></table>';
document.getElementById('ages').innerHTML = ahtml;

// Duplicates
let dhtml = '';
if (!data.duplicates || data.duplicates.length === 0) {
  dhtml = '<p>No duplicate groups found (or none above the size threshold).</p>';
} else {
  (data.duplicates || []).forEach(g => {
    dhtml += '<div class="dup-group">' +
      '<h3>' + (g.count||0) + ' copies · ' + (g.sizeFormatted||'') + ' each · <span class="waste">waste ' + (g.wasteFormatted||'') + '</span></h3>' +
      '<ul style="margin:0;padding-left:1.2rem">';
    (g.paths || []).forEach(p => { dhtml += '<li class="path">' + p + '</li>'; });
    dhtml += '</ul></div>';
  });
}
document.getElementById('dupes').innerHTML = dhtml;

// Tabs
document.querySelectorAll('.tab').forEach(btn => {
  btn.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.pane').forEach(p => p.classList.remove('active'));
    btn.classList.add('active');
    document.getElementById(btn.dataset.pane).classList.add('active');
  });
});
</script>
</body>
</html>
"@

$html = $html.Replace('__B64_PLACEHOLDER__', $b64)
$html | Set-Content -Path $htmlPath -Encoding UTF8

# ---------- Done ----------
Write-Status "Done!" "Green"
Write-Host ""
Write-Host "Reports:" -ForegroundColor Cyan
Write-Host "  HTML : $htmlPath"
Write-Host "  JSON : $jsonPath"
Write-Host "  CSVs : $OutDir\*-$stamp.csv"
Write-Host ""
Write-Status "Opening HTML report..." "Cyan"
Start-Process $htmlPath
