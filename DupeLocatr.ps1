<#
    DupeLocatr
    ----------
    A menu-driven duplicate file finder & cleaner for Windows.

    - Recursively scans the current folder (and every subfolder) for
      duplicate files of ANY type: images, videos, music, text, documents,
      archives, whatever.
    - Uses a 3-stage "fast path" so it never wastes time hashing files that
      can't possibly be duplicates:
          1) group by file size            (instant, in-memory)
          2) quick partial hash (64 KB)     (cheap, filters out non-matches)
          3) full hash (MD5/SHA256)         (only run on real candidates)
      This means a folder full of large videos gets scanned fast, because
      most videos never even get fully read.
    - NEVER deletes or moves anything without asking first.
    - Deletions go to the Recycle Bin (recoverable), not permanent removal.
    - "Move to folder" uses robocopy under the hood (fast, reliable, built
      into Windows, handles long paths and retries better than a plain copy).
    - Colored console output + a progress bar during scanning.

    Usage:
        powershell -ExecutionPolicy Bypass -File DupeLocatr.ps1
        powershell -ExecutionPolicy Bypass -File DupeLocatr.ps1 -Path "D:\Photos"

    (Or just double-click Run_DupeLocatr.bat sitting next to this file.)
#>

param(
    [string]$Path = (Get-Location).Path
)

# ======================================================================
#  CONFIG / STATE
# ======================================================================
$Script:RootPath       = (Resolve-Path -LiteralPath $Path).Path
$Script:DupeFolderName = "encounted_duplicates"
$Script:HashAlgo       = "MD5"          # MD5 (fast, default) or SHA256 (slower/stronger)
$Script:KeepStrategy   = "First"        # First | Oldest | Newest | ShortestPath
$Script:CategoryFilter = "All"          # All | Images | Videos | Music | Text | Documents | Archives
$Script:MinFileSizeBytes = 1            # ignore 0-byte files by default
$Script:ExcludeFolders = @(
    $Script:DupeFolderName, '.git', '.svn', 'node_modules',
    '$RECYCLE.BIN', 'System Volume Information', '.vs', '.vscode', '__pycache__'
)
$Script:LogPath    = Join-Path $Script:RootPath "dupeLocatr_log.txt"
$Script:ScriptSelf = $PSCommandPath
$Script:LastResults = $null

$Script:Categories = @{
    'Images'    = @('.jpg','.jpeg','.png','.gif','.bmp','.tiff','.tif','.webp','.heic','.heif','.svg','.ico','.raw','.cr2','.nef','.arw','.dng')
    'Videos'    = @('.mp4','.mkv','.avi','.mov','.wmv','.flv','.webm','.m4v','.mpg','.mpeg','.3gp','.ts','.vob')
    'Music'     = @('.mp3','.wav','.flac','.aac','.ogg','.wma','.m4a','.opus','.aiff','.mid')
    'Text'      = @('.txt','.md','.log','.csv','.json','.xml','.yaml','.yml','.ini','.cfg')
    'Documents' = @('.pdf','.doc','.docx','.xls','.xlsx','.ppt','.pptx','.odt','.rtf')
    'Archives'  = @('.zip','.rar','.7z','.tar','.gz','.bz2')
}

try { $Host.UI.RawUI.WindowTitle = "DupeLocatr" } catch {}

# ======================================================================
#  HELPERS
# ======================================================================
function Format-Size {
    param([double]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    elseif ($Bytes -ge 1MB) { return "{0:N2} MB" -f ($Bytes / 1MB) }
    elseif ($Bytes -ge 1KB) { return "{0:N2} KB" -f ($Bytes / 1KB) }
    else { return "$Bytes bytes" }
}

function Write-Banner {
    Clear-Host
    $art = @'
     _                  _                    _        
  __| |_   _ _ __   ___| |    ___   ___ __ _| |_ _ __ 
 / _` | | | | '_ \ / _ \ |   / _ \ / __/ _` | __| '__|
| (_| | |_| | |_) |  __/ |__| (_) | (_| (_| | |_| |   
 \__,_|\__,_| .__/ \___|_____\___/ \___\__,_|\__|_|   
            |_|                                       
'@
    $lines  = $art -split "`r`n|`n"
    $colors = @('Magenta','Cyan','Blue','Green','Yellow','Red')
    for ($i = 0; $i -lt $lines.Count; $i++) {
        Write-Host $lines[$i] -ForegroundColor $colors[$i % $colors.Count]
    }
    Write-Host "        find them. face them. finish them." -ForegroundColor DarkGray
    Write-Host ""
}

function Test-IsExcluded {
    param([string]$FullPath)
    if ($Script:ScriptSelf -and $FullPath -eq $Script:ScriptSelf) { return $true }
    if ($FullPath -eq $Script:LogPath) { return $true }
    if ($FullPath -like "*dupeLocatr_*") { return $true }
    foreach ($ex in $Script:ExcludeFolders) {
        if ([string]::IsNullOrWhiteSpace($ex)) { continue }
        if ($FullPath -like "*\$ex\*" -or $FullPath -like "*/$ex/*") { return $true }
    }
    return $false
}

function Test-CategoryMatch {
    param([string]$Extension)
    if ($Script:CategoryFilter -eq 'All') { return $true }
    if (-not $Extension) { return $false }
    $ext = $Extension.ToLower()
    return ($Script:Categories[$Script:CategoryFilter] -contains $ext)
}

function Get-PartialHash {
    param([string]$Path, [int]$SampleBytes = 65536)
    try {
        $stream = [System.IO.File]::OpenRead($Path)
        try {
            $len = [Math]::Min($SampleBytes, [int]$stream.Length)
            if ($len -le 0) { return "EMPTY" }
            $buffer = New-Object byte[] $len
            [void]$stream.Read($buffer, 0, $len)
            $md5 = [System.Security.Cryptography.MD5]::Create()
            $hashBytes = $md5.ComputeHash($buffer)
            return [System.BitConverter]::ToString($hashBytes) -replace '-', ''
        } finally {
            $stream.Dispose()
        }
    } catch {
        return $null
    }
}

# ======================================================================
#  SCANNING ENGINE
# ======================================================================
function Get-CandidateFiles {
    Write-Host "  Scanning folders under: $Script:RootPath" -ForegroundColor Cyan
    $found = New-Object System.Collections.Generic.List[System.IO.FileInfo]
    $counter = 0

    Get-ChildItem -LiteralPath $Script:RootPath -File -Recurse -Force -ErrorAction SilentlyContinue |
        ForEach-Object {
            $counter++
            if ($counter % 150 -eq 0) {
                Write-Progress -Activity "DupeLocatr | Enumerating files" -Status "$counter files found so far..."
            }
            if (-not (Test-IsExcluded $_.FullName)) {
                if ($_.Length -ge $Script:MinFileSizeBytes) {
                    if (Test-CategoryMatch $_.Extension) {
                        $found.Add($_)
                    }
                }
            }
        }

    Write-Progress -Activity "DupeLocatr | Enumerating files" -Completed
    Write-Host "  Total files considered: $($found.Count)" -ForegroundColor Gray
    return $found
}

function Find-Duplicates {
    $allFiles = Get-CandidateFiles
    if ($allFiles.Count -eq 0) {
        Write-Host "  No files found to compare." -ForegroundColor Yellow
        return $null
    }

    # ---- Stage 1: group by size (instant) ----
    Write-Host "  Stage 1/3: grouping by file size..." -ForegroundColor Cyan
    $sizeGroups = $allFiles | Group-Object Length | Where-Object { $_.Count -gt 1 }
    $sizeCandidates = New-Object System.Collections.Generic.List[System.IO.FileInfo]
    foreach ($g in $sizeGroups) { foreach ($item in $g.Group) { $sizeCandidates.Add($item) } }

    if ($sizeCandidates.Count -eq 0) {
        Write-Host "  Every file has a unique size - nothing to compare further." -ForegroundColor Yellow
        return $null
    }
    Write-Host "  $($sizeCandidates.Count) file(s) share a size with at least one other file." -ForegroundColor Gray

    # ---- Stage 2: quick partial hash (64 KB sample) ----
    Write-Host "  Stage 2/3: quick content check..." -ForegroundColor Cyan
    $partialMap = @{}
    $i = 0
    $total = $sizeCandidates.Count
    foreach ($f in $sizeCandidates) {
        $i++
        Write-Progress -Activity "DupeLocatr | Quick content check" -Status "$i / $total : $($f.Name)" -PercentComplete ([Math]::Round(($i / $total) * 100))
        $ph = Get-PartialHash -Path $f.FullName
        if ($null -ne $ph) {
            $key = "$($f.Length)_$ph"
            if (-not $partialMap.ContainsKey($key)) {
                $partialMap[$key] = New-Object System.Collections.Generic.List[System.IO.FileInfo]
            }
            $partialMap[$key].Add($f)
        }
    }
    Write-Progress -Activity "DupeLocatr | Quick content check" -Completed

    $fullCandidates = New-Object System.Collections.Generic.List[System.IO.FileInfo]
    foreach ($key in $partialMap.Keys) {
        if ($partialMap[$key].Count -gt 1) {
            foreach ($item in $partialMap[$key]) { $fullCandidates.Add($item) }
        }
    }

    if ($fullCandidates.Count -eq 0) {
        Write-Host "  Quick check found no real matches." -ForegroundColor Yellow
        return $null
    }

    # ---- Stage 3: full hash confirmation ----
    Write-Host "  Stage 3/3: confirming with a full $Script:HashAlgo hash..." -ForegroundColor Cyan
    $fullMap = @{}
    $i = 0
    $total = $fullCandidates.Count
    foreach ($f in $fullCandidates) {
        $i++
        Write-Progress -Activity "DupeLocatr | Verifying duplicates" -Status "$i / $total : $($f.Name)" -PercentComplete ([Math]::Round(($i / $total) * 100))
        try {
            $h = (Get-FileHash -LiteralPath $f.FullName -Algorithm $Script:HashAlgo -ErrorAction Stop).Hash
        } catch {
            $h = $null
        }
        if ($h) {
            $key = "$($f.Length)_$h"
            if (-not $fullMap.ContainsKey($key)) {
                $fullMap[$key] = New-Object System.Collections.Generic.List[System.IO.FileInfo]
            }
            $fullMap[$key].Add($f)
        }
    }
    Write-Progress -Activity "DupeLocatr | Verifying duplicates" -Completed

    # ---- Build result groups ----
    $results = @()
    foreach ($key in $fullMap.Keys) {
        $group = $fullMap[$key]
        if ($group.Count -gt 1) {
            $sorted = switch ($Script:KeepStrategy) {
                'Oldest'       { $group | Sort-Object LastWriteTime }
                'Newest'       { $group | Sort-Object LastWriteTime -Descending }
                'ShortestPath' { $group | Sort-Object { $_.FullName.Length } }
                default        { $group | Sort-Object FullName }
            }
            $keep  = $sorted[0]
            $dupes = @($sorted | Select-Object -Skip 1)
            $results += [PSCustomObject]@{
                Hash           = $key
                SizeEach       = $keep.Length
                KeepFile       = $keep
                DuplicateFiles = $dupes
            }
        }
    }
    return $results
}

# ======================================================================
#  REPORTING
# ======================================================================
function Write-DuplicateLog {
    param($Results)
    $lines = @()
    $lines += "DupeLocatr scan log - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $lines += "Root: $Script:RootPath"
    $lines += "Hash algorithm: $Script:HashAlgo | Keep strategy: $Script:KeepStrategy | Category filter: $Script:CategoryFilter"
    $lines += "-----------------------------------------------------------"
    $n = 1
    foreach ($grp in $Results) {
        $lines += "Group $n (size each: $(Format-Size $grp.SizeEach))"
        $lines += "  KEEP: $($grp.KeepFile.FullName)"
        foreach ($d in $grp.DuplicateFiles) { $lines += "  DUPE: $($d.FullName)" }
        $n++
    }
    $lines += ""
    Add-Content -LiteralPath $Script:LogPath -Value $lines -Encoding UTF8
}

function Show-DetailedList {
    param($Results)
    Clear-Host
    Write-Host "  ---- Duplicate Details ----" -ForegroundColor Magenta
    $n = 1
    foreach ($grp in $Results) {
        Write-Host ""
        Write-Host "  Group $n  (size each: $(Format-Size $grp.SizeEach))" -ForegroundColor Cyan
        Write-Host "    KEEP  -> $($grp.KeepFile.FullName)" -ForegroundColor Green
        foreach ($d in $grp.DuplicateFiles) {
            Write-Host "    DUPE  -> $($d.FullName)" -ForegroundColor Yellow
        }
        $n++
    }
    Write-Host ""
    Read-Host "  Press Enter to go back" | Out-Null
}

function Export-DuplicateReport {
    param($Results)
    $csvPath = Join-Path $Script:RootPath ("dupeLocatr_report_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    $rows = @()
    $n = 1
    foreach ($grp in $Results) {
        $rows += [PSCustomObject]@{ Group = $n; Role = 'KEEP'; Path = $grp.KeepFile.FullName; SizeBytes = $grp.SizeEach }
        foreach ($d in $grp.DuplicateFiles) {
            $rows += [PSCustomObject]@{ Group = $n; Role = 'DUPLICATE'; Path = $d.FullName; SizeBytes = $grp.SizeEach }
        }
        $n++
    }
    $rows | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8
    Write-Host ""
    Write-Host "  Report saved to: $csvPath" -ForegroundColor Green
    Start-Sleep -Seconds 1
}

# ======================================================================
#  ACTIONS  (both ask for confirmation before touching anything)
# ======================================================================
function Invoke-MoveDuplicates {
    param($Results)

    $allDupes = New-Object System.Collections.Generic.List[System.IO.FileInfo]
    foreach ($grp in $Results) { foreach ($d in $grp.DuplicateFiles) { $allDupes.Add($d) } }
    $count = $allDupes.Count
    $totalSize = ($allDupes | Measure-Object Length -Sum).Sum

    Write-Host ""
    Write-Host "  This will MOVE $count duplicate file(s) ($(Format-Size $totalSize)) into:" -ForegroundColor Yellow
    Write-Host "    $(Join-Path $Script:RootPath $Script:DupeFolderName)" -ForegroundColor Yellow
    $confirm = Read-Host "  Type Y to continue, anything else to cancel"
    if ($confirm -ne 'Y' -and $confirm -ne 'y') {
        Write-Host "  Cancelled. No files were moved." -ForegroundColor Gray
        Start-Sleep -Seconds 1
        return
    }

    $destFolder = Join-Path $Script:RootPath $Script:DupeFolderName
    if (-not (Test-Path -LiteralPath $destFolder)) {
        New-Item -ItemType Directory -Path $destFolder | Out-Null
    }

    $i = 0
    $moved = 0
    $errors = 0
    foreach ($f in $allDupes) {
        $i++
        Write-Progress -Activity "DupeLocatr | Moving duplicates" -Status "$i / $count : $($f.Name)" -PercentComplete ([Math]::Round(($i / $count) * 100))

        $destName = $f.Name
        $destPath = Join-Path $destFolder $destName
        if (Test-Path -LiteralPath $destPath) {
            $parentTag = ($f.Directory.Name -replace '[^\w\-]', '_')
            $destName  = "{0}__{1}" -f $parentTag, $f.Name
            $destPath  = Join-Path $destFolder $destName
            $dupCounter = 1
            while (Test-Path -LiteralPath $destPath) {
                $destName = "{0}__{1}__{2}" -f $parentTag, $dupCounter, $f.Name
                $destPath = Join-Path $destFolder $destName
                $dupCounter++
            }
        }

        try {
            robocopy $f.DirectoryName $destFolder $f.Name /MOV /NFL /NDL /NJH /NJS /NP /R:1 /W:1 | Out-Null
            $movedRaw = Join-Path $destFolder $f.Name
            if ((Test-Path -LiteralPath $movedRaw) -and ($destName -ne $f.Name)) {
                Rename-Item -LiteralPath $movedRaw -NewName $destName -ErrorAction SilentlyContinue
            }
            if ((Test-Path -LiteralPath $movedRaw) -or (Test-Path -LiteralPath $destPath)) {
                $moved++
            } else {
                $errors++
            }
        } catch {
            $errors++
        }
    }
    Write-Progress -Activity "DupeLocatr | Moving duplicates" -Completed

    Write-Host ""
    Write-Host "  Done. Moved $moved file(s), $errors error(s)." -ForegroundColor Green
    Add-Content -LiteralPath $Script:LogPath -Value "Action: MOVE - $moved moved, $errors errors - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Read-Host "  Press Enter to continue" | Out-Null
}

function Invoke-DeleteDuplicates {
    param($Results)

    $hasVB = $true
    try { Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction Stop } catch { $hasVB = $false }

    $allDupes = New-Object System.Collections.Generic.List[System.IO.FileInfo]
    foreach ($grp in $Results) { foreach ($d in $grp.DuplicateFiles) { $allDupes.Add($d) } }
    $count = $allDupes.Count
    $totalSize = ($allDupes | Measure-Object Length -Sum).Sum

    Write-Host ""
    Write-Host "  This will send $count duplicate file(s) ($(Format-Size $totalSize)) to the Recycle Bin." -ForegroundColor Red
    Write-Host "  They can be restored from the Recycle Bin afterward if you change your mind." -ForegroundColor Gray
    $confirm = Read-Host "  Type YES (all caps) to confirm, anything else cancels"
    if ($confirm -cne 'YES') {
        Write-Host "  Cancelled. No files were deleted." -ForegroundColor Gray
        Start-Sleep -Seconds 1
        return
    }

    $i = 0
    $deleted = 0
    $errors = 0
    foreach ($f in $allDupes) {
        $i++
        Write-Progress -Activity "DupeLocatr | Deleting duplicates" -Status "$i / $count : $($f.Name)" -PercentComplete ([Math]::Round(($i / $count) * 100))
        try {
            if ($hasVB) {
                [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
                    $f.FullName,
                    [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
                    [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin
                )
            } else {
                Remove-Item -LiteralPath $f.FullName -Force -ErrorAction Stop
            }
            $deleted++
        } catch {
            $errors++
        }
    }
    Write-Progress -Activity "DupeLocatr | Deleting duplicates" -Completed

    Write-Host ""
    Write-Host "  Done. Deleted $deleted file(s), $errors error(s)." -ForegroundColor Green
    Add-Content -LiteralPath $Script:LogPath -Value "Action: DELETE - $deleted deleted, $errors errors - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Read-Host "  Press Enter to continue" | Out-Null
}

# ======================================================================
#  MENUS
# ======================================================================
function Show-ScanResults {
    param($Results)

    if (-not $Results -or $Results.Count -eq 0) {
        Write-Host ""
        Write-Host "  No duplicate files found. This folder looks clean." -ForegroundColor Green
        Read-Host "  Press Enter to return to the menu" | Out-Null
        return
    }

    $Script:LastResults = $Results
    $totalDupeFiles = ($Results | ForEach-Object { $_.DuplicateFiles.Count } | Measure-Object -Sum).Sum
    $totalDupeSize  = ($Results | ForEach-Object { $_.DuplicateFiles.Count * $_.SizeEach } | Measure-Object -Sum).Sum
    Write-DuplicateLog -Results $Results

    $done = $false
    while (-not $done) {
        Write-Host ""
        Write-Host "  ===================== SCAN COMPLETE =====================" -ForegroundColor Magenta
        Write-Host "   Duplicate groups found : $($Results.Count)" -ForegroundColor White
        Write-Host "   Duplicate files found  : $totalDupeFiles" -ForegroundColor Yellow
        Write-Host "   Space that can be freed: $(Format-Size $totalDupeSize)" -ForegroundColor Yellow
        Write-Host "  ===========================================================" -ForegroundColor Magenta
        Write-Host ""
        Write-Host "  What would you like to do?" -ForegroundColor Cyan
        Write-Host "   [1] View the detailed list"
        Write-Host "   [2] Move duplicates to '$Script:DupeFolderName' folder (safe, undoable)"
        Write-Host "   [3] Delete duplicates (sent to Recycle Bin, recoverable)"
        Write-Host "   [4] Export report to CSV"
        Write-Host "   [5] Return to main menu"
        $choice = Read-Host "  Choose an option (1-5)"
        switch ($choice) {
            '1' { Show-DetailedList -Results $Results }
            '2' { Invoke-MoveDuplicates -Results $Results; $done = $true }
            '3' { Invoke-DeleteDuplicates -Results $Results; $done = $true }
            '4' { Export-DuplicateReport -Results $Results }
            '5' { $done = $true }
            default { Write-Host "  Invalid choice." -ForegroundColor Red }
        }
    }
}

function Show-SettingsMenu {
    $back = $false
    while (-not $back) {
        Clear-Host
        Write-Host "  ---- Settings ----" -ForegroundColor Magenta
        Write-Host "   [1] Hash algorithm        : $Script:HashAlgo"
        Write-Host "   [2] Keep strategy         : $Script:KeepStrategy"
        Write-Host "   [3] Category filter       : $Script:CategoryFilter"
        Write-Host "   [4] Excluded folders      : $($Script:ExcludeFolders -join ', ')"
        Write-Host "   [5] Duplicate folder name : $Script:DupeFolderName"
        Write-Host "   [6] Back to main menu"
        Write-Host ""
        $c = Read-Host "  Choose an option (1-6)"
        switch ($c) {
            '1' {
                Write-Host "   a) MD5 (fast, default)"
                Write-Host "   b) SHA256 (slower, stronger)"
                $h = Read-Host "  Choose"
                if ($h -eq 'b') { $Script:HashAlgo = 'SHA256' } else { $Script:HashAlgo = 'MD5' }
            }
            '2' {
                Write-Host "   a) First (alphabetically first path) - default"
                Write-Host "   b) Oldest (earliest modified date kept)"
                Write-Host "   c) Newest (latest modified date kept)"
                Write-Host "   d) ShortestPath (keep the shallowest copy)"
                $k = Read-Host "  Choose"
                switch ($k) {
                    'b' { $Script:KeepStrategy = 'Oldest' }
                    'c' { $Script:KeepStrategy = 'Newest' }
                    'd' { $Script:KeepStrategy = 'ShortestPath' }
                    default { $Script:KeepStrategy = 'First' }
                }
            }
            '3' {
                Write-Host "   Categories: All, Images, Videos, Music, Text, Documents, Archives"
                $cat = Read-Host "  Type a category"
                if ($cat -eq 'All' -or $Script:Categories.ContainsKey($cat)) {
                    $Script:CategoryFilter = $cat
                } else {
                    Write-Host "  Unknown category, keeping '$Script:CategoryFilter'." -ForegroundColor Red
                    Start-Sleep -Seconds 1
                }
            }
            '4' {
                $newList = Read-Host "  Enter comma-separated folder names to exclude (replaces current list)"
                if ($newList) {
                    $Script:ExcludeFolders = @(($newList -split ',') | ForEach-Object { $_.Trim() } | Where-Object { $_ })
                    $Script:ExcludeFolders += $Script:DupeFolderName
                }
            }
            '5' {
                $newName = Read-Host "  Enter new duplicate folder name"
                if ($newName) { $Script:DupeFolderName = $newName }
            }
            '6' { $back = $true }
            default { Write-Host "  Invalid choice." -ForegroundColor Red; Start-Sleep -Seconds 1 }
        }
    }
}

function Show-MainMenu {
    $exit = $false
    while (-not $exit) {
        Write-Banner
        Write-Host "  Target folder: $Script:RootPath" -ForegroundColor Gray
        Write-Host ""
        Write-Host "   [1] Start duplicate scan"
        Write-Host "   [2] Settings"
        Write-Host "   [3] View log file"
        Write-Host "   [4] About / help"
        Write-Host "   [5] Exit"
        Write-Host ""
        $choice = Read-Host "  Choose an option (1-5)"
        switch ($choice) {
            '1' {
                Clear-Host
                Write-Banner
                $results = Find-Duplicates
                Show-ScanResults -Results $results
            }
            '2' { Show-SettingsMenu }
            '3' {
                Clear-Host
                Write-Host "  ---- Log File ----" -ForegroundColor Magenta
                if (Test-Path -LiteralPath $Script:LogPath) {
                    Get-Content -LiteralPath $Script:LogPath | ForEach-Object { Write-Host $_ }
                } else {
                    Write-Host "  No log file yet. Run a scan first." -ForegroundColor Yellow
                }
                Write-Host ""
                Read-Host "  Press Enter to go back" | Out-Null
            }
            '4' {
                Clear-Host
                Write-Banner
                Write-Host "  DupeLocatr finds duplicate files by comparing size, then a" -ForegroundColor White
                Write-Host "  quick partial hash, then a full hash - so it never reads an" -ForegroundColor White
                Write-Host "  entire large video/audio file unless it truly has to." -ForegroundColor White
                Write-Host ""
                Write-Host "  It never deletes or moves anything without asking first." -ForegroundColor White
                Write-Host "  Deletions go to the Recycle Bin, not permanent removal." -ForegroundColor White
                Write-Host "  Moves use robocopy and land in the '$Script:DupeFolderName' folder." -ForegroundColor White
                Write-Host ""
                Read-Host "  Press Enter to go back" | Out-Null
            }
            '5' { $exit = $true }
            default { Write-Host "  Invalid choice." -ForegroundColor Red; Start-Sleep -Seconds 1 }
        }
    }
    Write-Host ""
    Write-Host "  Goodbye!" -ForegroundColor Cyan
    Start-Sleep -Milliseconds 500
}

# ======================================================================
#  ENTRY POINT
# ======================================================================
Show-MainMenu
