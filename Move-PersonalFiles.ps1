# Move-PersonalFiles.ps1
# Advanced file management utility for moving/copying personal media and documents
# Features: Drive selection, duplicate handling, long path support, multiple operation modes

[CmdletBinding()]
param(
    [Parameter(HelpMessage="Target drive letter (e.g., 'G'). If not specified, will prompt for selection.")]
    [string]$Drive,
    
    [Parameter(HelpMessage="Operation mode: Preview, Move, Copy, Cleanup")]
    [ValidateSet("Preview", "Move", "Copy", "Cleanup")]
    [string]$Mode = "Preview",
    
    [Parameter(HelpMessage="Source folders to scan (comma-separated). Default: Documents,Pictures,Downloads,Videos,Desktop")]
    [string[]]$SourceFolders = @('Documents','Pictures','Downloads','Videos','Desktop'),
    
    [Parameter(HelpMessage="File extensions to include (comma-separated). If not specified, uses default media/document extensions.")]
    [string[]]$Extensions,
    
    [Parameter(HelpMessage="How to handle duplicate files: Skip, Overwrite, Rename")]
    [ValidateSet("Skip", "Overwrite", "Rename")]
    [string]$DuplicateHandling = "Skip",
    
    [Parameter(HelpMessage="Enable verbose logging")]
    [switch]$Verbose,
    
    [Parameter(HelpMessage="Skip confirmation prompts (use with caution)")]
    [switch]$Force,
    
    [Parameter(HelpMessage="Maximum file age in days (0 = no limit)")]
    [int]$MaxAgeDays = 0,
    
    [Parameter(HelpMessage="Minimum file size in MB (0 = no limit)")]
    [int]$MinSizeMB = 0,
    
    [Parameter(HelpMessage="Show help and exit")]
    [switch]$Help
)

# ========== CONFIG ==========
$defaultTextExts  = @('.txt','.doc','.docx','.pdf','.rtf','.md','.odt','.xls','.xlsx','.ppt','.pptx')
$defaultImageExts = @('.jpg','.jpeg','.png','.bmp','.gif','.heic','.tif','.tiff','.webp','.raw','.cr2','.nef')
$defaultVideoExts = @('.mp4','.mov','.avi','.mkv','.wmv','.flv','.webm','.m4v','.mpg','.mpeg')
$defaultAudioExts = @('.mp3','.wav','.flac','.aac','.ogg','.wma','.m4a')
$defaultAllExts   = $defaultTextExts + $defaultImageExts + $defaultVideoExts + $defaultAudioExts

# Use provided extensions or defaults
$allExts = if ($Extensions) { $Extensions } else { $defaultAllExts }

# Global variables
$script:logFile = ""
$script:backupFolder = ""
$script:duplicateLog = ""

function Show-Help {
    Write-Host @"
Move-PersonalFiles.ps1 - Advanced Personal File Management Utility

DESCRIPTION:
    Moves, copies, or manages personal media and document files with advanced features
    including duplicate handling, long path support, and multiple operation modes.

SYNTAX:
    .\Move-PersonalFiles.ps1 [-Drive <string>] [-Mode <string>] [-SourceFolders <string[]>] 
                            [-Extensions <string[]>] [-DuplicateHandling <string>] [-Verbose] 
                            [-Force] [-MaxAgeDays <int>] [-MinSizeMB <int>] [-Help]

PARAMETERS:
    -Drive              Target drive letter (e.g., 'G'). Will prompt if not specified.
    -Mode               Operation mode: Preview, Move, Copy, Cleanup (Default: Preview)
    -SourceFolders      Source folders to scan (Default: Documents,Pictures,Downloads,Videos,Desktop)
    -Extensions         File extensions to include (Default: common media/document types)
    -DuplicateHandling  How to handle duplicates: Skip, Overwrite, Rename (Default: Skip)
    -Verbose            Enable detailed logging
    -Force              Skip confirmation prompts
    -MaxAgeDays         Only process files newer than X days (0 = no limit)
    -MinSizeMB          Only process files larger than X MB (0 = no limit)
    -Help               Show this help message

EXAMPLES:
    .\Move-PersonalFiles.ps1 -Help
    .\Move-PersonalFiles.ps1 -Mode Preview
    .\Move-PersonalFiles.ps1 -Drive G -Mode Move -DuplicateHandling Rename
    .\Move-PersonalFiles.ps1 -Mode Copy -SourceFolders Pictures,Downloads -Force
    .\Move-PersonalFiles.ps1 -Mode Cleanup -MaxAgeDays 30 -MinSizeMB 10

"@ -ForegroundColor Cyan
}

function Write-Log { 
    param([string]$msg, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $msg"
    Add-Content -Path $script:logFile -Value $logEntry
    if ($Verbose -or $Level -eq "ERROR") {
        $color = switch ($Level) {
            "ERROR" { "Red" }
            "WARN" { "Yellow" }
            "SUCCESS" { "Green" }
            default { "White" }
        }
        Write-Host $logEntry -ForegroundColor $color
    }
}

function Test-LongPath {
    param([string]$path)
    # Enable long path support for Windows 10 version 1607 and later
    if ($path.Length -gt 260) {
        if ($path -notmatch '^\\\\?\') {
            return "\\?\$path"
        }
    }
    return $path
}

function Get-FileHashSafe {
    param([string]$filePath)
    try {
        $safePath = Test-LongPath $filePath
        return (Get-FileHash -Path $safePath -Algorithm SHA256 -ErrorAction Stop).Hash
    }
    catch {
        Write-Log "Failed to calculate hash for: $filePath - $_" "WARN"
        return $null
    }
}

function Format-Size {
    param([int64]$bytes)
    if ($bytes -gt 1TB) { return "{0:N2} TB" -f ($bytes / 1TB) }
    elseif ($bytes -gt 1GB) { return "{0:N2} GB" -f ($bytes / 1GB) }
    elseif ($bytes -gt 1MB) { return "{0:N2} MB" -f ($bytes / 1MB) }
    elseif ($bytes -gt 1KB) { return "{0:N2} KB" -f ($bytes / 1KB) }
    else { return "$bytes bytes" }
}

function Get-AvailableDrives {
    $drives = Get-WmiObject -Class Win32_LogicalDisk | Where-Object { 
        $_.DriveType -eq 2 -or $_.DriveType -eq 3  # Removable or Fixed drives
    } | Sort-Object DeviceID
    return $drives
}

function Select-TargetDrive {
    if ($Drive) {
        $targetDrive = $Drive.TrimEnd(':') + ':'
        if (Test-Path $targetDrive) {
            return $targetDrive
        } else {
            Write-Host "ERROR: Drive $targetDrive not found." -ForegroundColor Red
            return $null
        }
    }

    Write-Host "`nAvailable drives:" -ForegroundColor Cyan
    $drives = Get-AvailableDrives
    $index = 1
    foreach ($drive in $drives) {
        $freeGB = [math]::Round($drive.FreeSpace / 1GB, 2)
        $totalGB = [math]::Round($drive.Size / 1GB, 2)
        $type = switch ($drive.DriveType) {
            2 { "Removable" }
            3 { "Fixed" }
            default { "Other" }
        }
        Write-Host "  $index. $($drive.DeviceID) - $($drive.VolumeName) [$type] ($freeGB GB free of $totalGB GB)"
        $index++
    }
    
    do {
        $selection = Read-Host "`nSelect drive number (1-$($drives.Count))"
        if ($selection -match '^\d+$' -and [int]$selection -ge 1 -and [int]$selection -le $drives.Count) {
            return $drives[[int]$selection - 1].DeviceID
        }
        Write-Host "Invalid selection. Please try again." -ForegroundColor Yellow
    } while ($true)
}

function Scan-Files {
    param([string[]]$folders, [string[]]$extensions)
    
    Write-Host "`nScanning folders for eligible files..." -ForegroundColor Cyan
    $allFiles = @()
    $skippedFolders = @()
    $processedCount = 0
    
    foreach ($folder in $folders) {
        $src = Join-Path $HOME $folder
        if (-not (Test-Path $src)) {
            Write-Log "Folder not found: $src" "WARN"
            $skippedFolders += $folder
            continue
        }
        
        Write-Host "  Scanning: $src" -ForegroundColor Gray
        
        try {
            foreach ($ext in $extensions) {
                $files = Get-ChildItem -Path $src -Recurse -File -Filter "*$ext" -ErrorAction SilentlyContinue
                foreach ($file in $files) {
                    try {
                        # Apply filters
                        if ($MaxAgeDays -gt 0) {
                            $cutoffDate = (Get-Date).AddDays(-$MaxAgeDays)
                            if ($file.LastWriteTime -lt $cutoffDate) {
                                continue
                            }
                        }
                        
                        if ($MinSizeMB -gt 0) {
                            $minSizeBytes = $MinSizeMB * 1MB
                            if ($file.Length -lt $minSizeBytes) {
                                continue
                            }
                        }
                        
                        $allFiles += $file
                        $processedCount++
                        
                        if ($processedCount % 100 -eq 0) {
                            Write-Host "    Found $processedCount files..." -ForegroundColor Gray
                        }
                    }
                    catch {
                        Write-Log "Error processing file $($file.FullName): $_" "WARN"
                    }
                }
            }
        }
        catch {
            Write-Log "Error scanning folder $src : $($_.Exception.Message)" "ERROR"
            $skippedFolders += $folder
        }
    }
    
    # Remove duplicates based on full path
    $uniqueFiles = $allFiles | Sort-Object FullName | Get-Unique -AsString
    
    Write-Host "  Scan complete: $($uniqueFiles.Count) unique files found" -ForegroundColor Green
    if ($skippedFolders.Count -gt 0) {
        Write-Log "Skipped folders: $($skippedFolders -join ', ')" "WARN"
    }
    
    return $uniqueFiles
}

function Test-DuplicateFile {
    param([string]$sourcePath, [string]$destPath)
    
    if (-not (Test-Path $destPath)) {
        return @{ IsDuplicate = $false; Action = "None" }
    }
    
    try {
        $sourceFile = Get-Item $sourcePath
        $destFile = Get-Item $destPath
        
        # Quick check: if sizes are different, they're different files
        if ($sourceFile.Length -ne $destFile.Length) {
            return @{ IsDuplicate = $false; Action = "SizeDifference" }
        }
        
        # If sizes match, check hashes
        $sourceHash = Get-FileHashSafe $sourcePath
        $destHash = Get-FileHashSafe $destPath
        
        if ($sourceHash -and $destHash -and $sourceHash -eq $destHash) {
            return @{ IsDuplicate = $true; Action = "Identical" }
        }
        else {
            return @{ IsDuplicate = $false; Action = "HashDifference" }
        }
    }
    catch {
        Write-Log "Error comparing files $sourcePath and $destPath : $($_.Exception.Message)" "WARN"
        return @{ IsDuplicate = $false; Action = "Error" }
    }
}

function Get-UniqueDestinationPath {
    param([string]$originalPath)
    
    $directory = Split-Path $originalPath
    $nameWithoutExt = [System.IO.Path]::GetFileNameWithoutExtension($originalPath)
    $extension = [System.IO.Path]::GetExtension($originalPath)
    
    $counter = 1
    do {
        $newName = "${nameWithoutExt}_$counter$extension"
        $newPath = Join-Path $directory $newName
        $counter++
    } while (Test-Path $newPath)
    
    return $newPath
}

function Process-Files {
    param(
        [array]$filesToProcess,
        [string]$operation,
        [string]$targetFolder
    )
    
    $count = $filesToProcess.Count
    $totalBytes = ($filesToProcess | Measure-Object -Property Length -Sum).Sum
    $totalDisplay = Format-Size $totalBytes

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Yellow
    Write-Host "Operation: $operation"
    Write-Host "Files to process: $count"
    Write-Host "Total size: $totalDisplay"
    Write-Host "Target folder: $targetFolder"
    Write-Host "Duplicate handling: $DuplicateHandling"
    Write-Host "============================================================" -ForegroundColor Yellow
    Write-Host ""

    Write-Log "=== $operation SUMMARY at $(Get-Date) ==="
    Write-Log "Files: $count"
    Write-Log "Total Size: $totalDisplay"
    Write-Log "Target: $targetFolder"
    Write-Log "Duplicate Handling: $DuplicateHandling"

    if ($operation -eq "Preview") {
        Write-Host "This is a PREVIEW. No files will be modified." -ForegroundColor Cyan
        Write-Host "Files that would be processed:" -ForegroundColor Gray
        
        $previewCount = 0
        foreach ($file in $filesToProcess) {
            $relPath = $file.FullName.Substring($HOME.Length + 1)
            $dest = Join-Path $targetFolder $relPath
            Write-Log "Would process: $($file.FullName) --> $dest"
            
            $previewCount++
            if ($previewCount -le 10) {
                Write-Host "  $($file.FullName)" -ForegroundColor Gray
            } elseif ($previewCount -eq 11) {
                Write-Host "  ... and $($count - 10) more files (see log for complete list)" -ForegroundColor Gray
            }
        }
        
        Write-Host "`nPreview complete. See $($script:logFile) for full details." -ForegroundColor Green
        return
    }

    # Confirmation for actual operations
    if (-not $Force) {
        $action = if ($operation -eq "Move") { "MOVE" } elseif ($operation -eq "Copy") { "COPY" } else { "CLEAN UP" }
        $confirmation = Read-Host "Proceed to $action these files? Type 'YES' to continue"
        if ($confirmation -ne 'YES') {
            Write-Host "Operation cancelled." -ForegroundColor Yellow
            return
        }
    }

    # Begin processing
    Write-Log "`n=== Begin $operation operation at $(Get-Date) ==="
    $errorCount = 0
    $processedCount = 0
    $skippedCount = 0
    $duplicateCount = 0
    $progressCounter = 0

    foreach ($file in $filesToProcess) {
        $progressCounter++
        if ($progressCounter % 50 -eq 0 -or $progressCounter -eq $count) {
            Write-Progress -Activity "$operation Files" -Status "Processing file $progressCounter of $count" -PercentComplete (($progressCounter / $count) * 100)
        }

        try {
            if ($operation -eq "Cleanup") {
                # For cleanup mode, just delete the file
                Remove-Item -Path $file.FullName -Force -ErrorAction Stop
                Write-Log "Deleted: $($file.FullName)" "SUCCESS"
                $processedCount++
                continue
            }

            $relPath = $file.FullName.Substring($HOME.Length + 1)
            $dest = Join-Path $targetFolder $relPath
            $destDir = Split-Path $dest
            
            # Create destination directory if it doesn't exist
            if (-not (Test-Path $destDir)) {
                New-Item -ItemType Directory -Path $destDir -Force | Out-Null
            }
            
            # Check for duplicates
            $duplicateCheck = Test-DuplicateFile $file.FullName $dest
            if ($duplicateCheck.IsDuplicate) {
                $duplicateCount++
                Add-Content -Path $script:duplicateLog -Value "DUPLICATE: $($file.FullName) <-> $dest"
                
                switch ($DuplicateHandling) {
                    "Skip" {
                        Write-Log "Skipped duplicate: $($file.FullName)" "WARN"
                        $skippedCount++
                        continue
                    }
                    "Rename" {
                        $dest = Get-UniqueDestinationPath $dest
                    }
                    "Overwrite" {
                        # Continue with original destination
                    }
                }
            }
            
            # Perform the operation
            if ($operation -eq "Move") {
                Move-Item -Path $file.FullName -Destination $dest -Force -ErrorAction Stop
                Write-Log "Moved: $($file.FullName) --> $dest" "SUCCESS"
            }
            elseif ($operation -eq "Copy") {
                Copy-Item -Path $file.FullName -Destination $dest -Force -ErrorAction Stop
                Write-Log "Copied: $($file.FullName) --> $dest" "SUCCESS"
            }
            
            $processedCount++
        }
        catch {
            Write-Log "ERROR processing $($file.FullName): $_" "ERROR"
            $errorCount++
        }
    }
    
    Write-Progress -Activity "$operation Files" -Completed
    
    Write-Log "=== $operation completed at $(Get-Date) ==="
    Write-Log "Processed: $processedCount | Skipped: $skippedCount | Duplicates: $duplicateCount | Errors: $errorCount"
    
    Write-Host ""
    Write-Host "Operation complete!" -ForegroundColor Green
    Write-Host "  Processed: $processedCount files" -ForegroundColor Green
    Write-Host "  Skipped: $skippedCount files" -ForegroundColor Yellow
    Write-Host "  Duplicates found: $duplicateCount files" -ForegroundColor Cyan
    Write-Host "  Errors: $errorCount files" -ForegroundColor Red
    Write-Host ""
    Write-Host "Log file: $($script:logFile)" -ForegroundColor Gray
    if ($duplicateCount -gt 0) {
        Write-Host "Duplicate log: $($script:duplicateLog)" -ForegroundColor Gray
    }
}

# ========== MAIN SCRIPT ==========

# Show help if requested
if ($Help) {
    Show-Help
    exit 0
}

# Display banner
Write-Host ""
Write-Host "╭─────────────────────────────────────────────────────────────╮" -ForegroundColor Cyan
Write-Host "│              Personal File Management Utility              │" -ForegroundColor Cyan
Write-Host "│        Move, Copy, or Clean up your personal files         │" -ForegroundColor Cyan
Write-Host "╰─────────────────────────────────────────────────────────────╯" -ForegroundColor Cyan
Write-Host ""

# Select target drive
$targetDrive = Select-TargetDrive
if (-not $targetDrive) {
    Write-Host "No valid drive selected. Exiting." -ForegroundColor Red
    exit 1
}

# Set up paths
$script:backupFolder = "$targetDrive\User_MediaBackup_$env:USERNAME\$(Get-Date -Format 'yyyy-MM-dd_HHmm')"
$script:logFile = "$($script:backupFolder)\operation_log.txt"
$script:duplicateLog = "$($script:backupFolder)\duplicates_log.txt"

# Create backup folder and log file
if (-not (Test-Path $script:backupFolder)) {
    try {
        New-Item -ItemType Directory -Path $script:backupFolder -Force | Out-Null
        Write-Host "Created backup folder: $($script:backupFolder)" -ForegroundColor Green
    }
    catch {
        Write-Host "ERROR: Cannot create backup folder: $_" -ForegroundColor Red
        exit 1
    }
}

if (-not (Test-Path $script:logFile)) {
    try {
        New-Item -ItemType File -Path $script:logFile -Force | Out-Null
    }
    catch {
        Write-Host "ERROR: Cannot create log file: $_" -ForegroundColor Red
        exit 1
    }
}

# Interactive mode selection if not provided
if (-not $Mode -or $Mode -eq "Preview") {
    Write-Host "Select operation mode:" -ForegroundColor Cyan
    Write-Host "  1. Preview - Show what would be processed (safe)" -ForegroundColor Green
    Write-Host "  2. Copy - Copy files to backup location" -ForegroundColor Yellow
    Write-Host "  3. Move - Move files to backup location" -ForegroundColor Red
    Write-Host "  4. Cleanup - Delete files from source (DESTRUCTIVE!)" -ForegroundColor Magenta
    
    do {
        $choice = Read-Host "`nSelect mode (1-4)"
        switch ($choice) {
            "1" { $Mode = "Preview"; break }
            "2" { $Mode = "Copy"; break }
            "3" { $Mode = "Move"; break }
            "4" { 
                $Mode = "Cleanup"
                Write-Host "WARNING: Cleanup mode will DELETE files permanently!" -ForegroundColor Red
                if (-not $Force) {
                    $confirm = Read-Host "Type 'DELETE' to confirm you want to use cleanup mode"
                    if ($confirm -ne "DELETE") {
                        Write-Host "Cleanup mode cancelled. Switching to Preview." -ForegroundColor Yellow
                        $Mode = "Preview"
                    }
                }
                break 
            }
            default { 
                Write-Host "Invalid selection. Please choose 1-4." -ForegroundColor Yellow
                continue 
            }
        }
        break
    } while ($true)
}

# Display current configuration
Write-Host "`nConfiguration:" -ForegroundColor Cyan
Write-Host "  Mode: $Mode" -ForegroundColor White
Write-Host "  Target Drive: $targetDrive" -ForegroundColor White
Write-Host "  Source Folders: $($SourceFolders -join ', ')" -ForegroundColor White
$extPreview = if ($allExts.Count -gt 5) { ($allExts[0..4] -join ', ') + "..." } else { ($allExts -join ', ') }
Write-Host "  File Extensions: $($allExts.Count) types ($extPreview)" -ForegroundColor White
Write-Host "  Duplicate Handling: $DuplicateHandling" -ForegroundColor White
if ($MaxAgeDays -gt 0) { Write-Host "  Max Age: $MaxAgeDays days" -ForegroundColor White }
if ($MinSizeMB -gt 0) { Write-Host "  Min Size: $MinSizeMB MB" -ForegroundColor White }

# Log initial configuration
Write-Log "=== Personal File Management Utility Started at $(Get-Date) ==="
Write-Log "Mode: $Mode"
Write-Log "Target Drive: $targetDrive"
Write-Log "Source Folders: $($SourceFolders -join ', ')"
Write-Log "Extensions: $($allExts -join ', ')"
Write-Log "Duplicate Handling: $DuplicateHandling"
Write-Log "Max Age Days: $MaxAgeDays"
Write-Log "Min Size MB: $MinSizeMB"
Write-Log "Force: $Force"
Write-Log "Verbose: $Verbose"

# Scan for files
$filesToProcess = Scan-Files -folders $SourceFolders -extensions $allExts

if ($filesToProcess.Count -eq 0) {
    Write-Host "`nNo files found matching the criteria." -ForegroundColor Yellow
    Write-Log "No files found matching criteria"
    exit 0
}

# Process files based on mode
if ($Mode -eq "Cleanup") {
    Process-Files -filesToProcess $filesToProcess -operation $Mode -targetFolder ""
} else {
    Process-Files -filesToProcess $filesToProcess -operation $Mode -targetFolder $script:backupFolder
}

Write-Host "`nOperation completed successfully!" -ForegroundColor Green
Write-Host "Check the log files for detailed information:" -ForegroundColor Gray
Write-Host "  Main log: $($script:logFile)" -ForegroundColor Gray
if (Test-Path $script:duplicateLog) {
    $dupLogPath = $script:duplicateLog
    Write-Host "  Duplicate log: $dupLogPath" -ForegroundColor Gray
} 