<#
.SYNOPSIS
    Automated folder backup script for Windows Task Scheduler.

.DESCRIPTION
    Copies a source folder to a timestamped destination under BACKUP_ROOT
    (e.g. D:\backup), optionally compresses it into a .zip, and removes
    backups older than RETENTION_DAYS. Configuration is read from a .env
    file located next to this script.

.NOTES
    Run manually to test:
        powershell -ExecutionPolicy Bypass -File backup_folder.ps1

    Or register it in Task Scheduler (see README.md for setup steps).
#>

$ErrorActionPreference = "Stop"

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$EnvFile   = Join-Path $ScriptDir ".env"

# --- Load config from .env --------------------------------------------------

if (-not (Test-Path $EnvFile)) {
    Write-Error "Missing .env file at $EnvFile (copy .env.example to get started)"
    exit 1
}

$Config = @{}
Get-Content $EnvFile | ForEach-Object {
    $line = $_.Trim()
    if ($line -eq "" -or $line.StartsWith("#")) { return }
    $parts = $line -split "=", 2
    if ($parts.Count -eq 2) {
        $key   = $parts[0].Trim()
        $value = $parts[1].Trim().Trim('"')
        $Config[$key] = $value
    }
}

function Get-Config {
    param($Key, $Default = $null)
    if ($Config.ContainsKey($Key) -and $Config[$Key] -ne "") {
        return $Config[$Key]
    }
    return $Default
}

$SourceDir     = Get-Config "SOURCE_DIR"
$BackupRoot    = Get-Config "BACKUP_ROOT" "D:\backup"
$RetentionDays = [int](Get-Config "RETENTION_DAYS" "7")
$Compress      = (Get-Config "COMPRESS" "true").ToLower() -eq "true"

if (-not $SourceDir) {
    Write-Error "SOURCE_DIR is not set in .env"
    exit 1
}
if (-not (Test-Path $SourceDir)) {
    Write-Error "Source folder not found: $SourceDir"
    exit 1
}
if (-not (Test-Path $BackupRoot)) {
    New-Item -ItemType Directory -Path $BackupRoot -Force | Out-Null
}

$Timestamp  = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$FolderName = Split-Path $SourceDir -Leaf
$LogFile    = Join-Path $BackupRoot "backup.log"

function Write-Log {
    param($Message)
    $line = "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] $Message"
    Write-Output $line
    Add-Content -Path $LogFile -Value $line
}

Write-Log "Starting backup of '$SourceDir'..."

$DestFolder = Join-Path $BackupRoot "${FolderName}_$Timestamp"

# --- Copy ---------------------------------------------------------------
# robocopy is used instead of Copy-Item because it handles long paths,
# locked files, and large trees far more reliably, and gives clear exit codes.

try {
    robocopy $SourceDir $DestFolder /E /R:3 /W:5 /NFL /NDL /NP /NJH /NJS | Out-Null
    $exitCode = $LASTEXITCODE

    # Robocopy exit codes 0-7 are all "success" variants (0 = nothing copied,
    # 1 = files copied, etc). 8 and above indicate a real failure.
    if ($exitCode -ge 8) {
        throw "Robocopy failed with exit code $exitCode"
    }

    Write-Log "Folder copied to $DestFolder"
} catch {
    Write-Log "Backup FAILED: $_"
    exit 1
}

# --- Compress (optional) -------------------------------------------------

if ($Compress) {
    $ZipPath = "$DestFolder.zip"
    try {
        Compress-Archive -Path (Join-Path $DestFolder '*') -DestinationPath $ZipPath -Force
        Remove-Item -Path $DestFolder -Recurse -Force
        Write-Log "Compressed to $ZipPath"
    } catch {
        Write-Log "Compression failed, keeping uncompressed folder instead: $_"
    }
}

# --- Retention cleanup -----------------------------------------------------

Write-Log "Cleaning up backups older than $RetentionDays day(s)..."

$CutoffDate = (Get-Date).AddDays(-$RetentionDays)
Get-ChildItem -Path $BackupRoot | Where-Object {
    $_.Name -like "${FolderName}_*" -and $_.LastWriteTime -lt $CutoffDate
} | ForEach-Object {
    Remove-Item -Path $_.FullName -Recurse -Force
    Write-Log "Deleted old backup: $($_.FullName)"
}

Write-Log "Done."