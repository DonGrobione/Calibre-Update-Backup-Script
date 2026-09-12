<#
.SYNOPSIS
    Backs up Calibre Portable, downloads and installs the latest portable update,
    and removes old backup sets based on retention.

.DESCRIPTION
    The script imports the StratoHiDriveUtils module (https://github.com/DonGrobione/StratoHiDriveUtils)
    and checks for updates via Update-StratoHiDriveUtils,
    uses it to resolve the HiDrive sync root and derive the Calibre installation and backup paths,
    downloads the current Calibre Portable installer to the TEMP folder,
    stops HiDrive to avoid sync/file lock issues during backup and update,
    creates a split 7z backup archive, installs the update, restarts HiDrive,
    and then deletes expired backups.

.EXAMPLE
    .\Calibre Update Backup.ps1

    Runs the full backup-update-cleanup workflow with automatic HiDrive path resolution.

.NOTES
    Version: 2.1.0
    Updated: 2026-09-12
    Mail: dongrobione@proton.me
    Latest version: https://github.com/DonGrobione/Calibre-Update-Backup-Script
    Requires: StratoHiDriveUtils module (https://github.com/DonGrobione/StratoHiDriveUtils)

    Log events should look like this:
    Write-Log -Message "This is an info level message." -LogLevel "Info"
    Write-Log -Message "This is an error level message." -LogLevel "Error"
#>

#Region: Definition of variables, change as needed

# Calibre Update URL
$CalibreUpdateSource = "https://calibre-ebook.com/dist/portable"

# Definition where the the update file will be downloaded to
$CalibreInstaller = "$env:TEMP\calibre-portable-installer.exe"

# 7zip binary
$7zipPath = "$env:ProgramFiles\7-Zip\7z.exe"

# Define Date sting in YYYY-MM-DD format for filename
$Date = (Get-Date).ToString("yyyy-MM-dd_HH-mm")

# Define number of backup datasets to be kept in $CalibreBackup folder and used in Remove-ExpiredBackups. Only the latest n set will be kept.
$CalibreBackupRetention = 3
#EndRegion

#Region: Functions
function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,
        [string]$LogLevel = "Info"
    )
    $LogPath = "$PSScriptRoot\Calibre-Backup-Update.log"
    $TimeStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $LogMessage = "$TimeStamp [$LogLevel] $Message"

    Add-Content -Path $LogPath -Value $LogMessage
}

function Initialize-StratoHiDriveUtils {
    <#
    Imports the StratoHiDriveUtils module (https://github.com/DonGrobione/StratoHiDriveUtils)
    and checks for updates using Update-StratoHiDriveUtils.
    #>
    $ModuleName = "StratoHiDriveUtils"

    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Log -Message "$ModuleName module was not found in PSModulePath. Please install it from https://github.com/DonGrobione/StratoHiDriveUtils" -LogLevel "Error"
        throw "$ModuleName module is not installed."
    }

    try {
        Import-Module -Name $ModuleName -ErrorAction Stop
        Write-Log -Message "$ModuleName module imported successfully." -LogLevel "Info"

        Write-Log -Message "Checking for $ModuleName module updates..." -LogLevel "Info"
        $updateResult = Update-StratoHiDriveUtils -ErrorAction Stop

        if ($updateResult.Status -eq 'Updated' -or $updateResult.ReloadRequired) {
            Import-Module -Name $ModuleName -Force -ErrorAction Stop
            Write-Log -Message "$ModuleName module updated from $($updateResult.LocalVersion) to $($updateResult.RemoteVersion) and reloaded." -LogLevel "Info"
        }
        elseif ($updateResult.Status -eq 'UpToDate') {
            Write-Log -Message "$ModuleName module is up to date (version $($updateResult.LocalVersion))." -LogLevel "Info"
        }
        elseif ($updateResult.Status -eq 'UpdateAvailable') {
            Write-Log -Message "$ModuleName module update available ($($updateResult.LocalVersion) -> $($updateResult.RemoteVersion))." -LogLevel "Info"
        }
    }
    catch {
        Write-Log -Message "Notice during $ModuleName initialization/update: $($_.Exception.Message)" -LogLevel "Warning"
        if (-not (Get-Command -Name Get-HiDriveSyncRoot -ErrorAction SilentlyContinue)) {
            throw
        }
    }
}

function Set-CalibreBackupPath {
    <#
    Determines CalibreBackupPath from the HiDrive sync root (via Get-HiDriveSyncRoot) instead of hostname matching.
    #>
    $HiDriveSyncRoot = Get-HiDriveSyncRoot
    if (-not $HiDriveSyncRoot) {
        Write-Log -Message "Could not determine HiDrive sync root. CalibreBackupPath not set." -LogLevel "Error"
        throw "Could not determine HiDrive sync root."
    }

    $script:CalibreBackupPath = Join-Path -Path $HiDriveSyncRoot -ChildPath "Backup\Calibre"
    Write-Log -Message "Calibre backups path was set to $script:CalibreBackupPath" -LogLevel "Info"

    if (-not (Test-Path -Path $script:CalibreBackupPath -PathType Container)) {
        Write-Log -Message "The path $script:CalibreBackupPath does not exist. Creating directory." -LogLevel "Info"
        New-Item -ItemType Directory -Path $script:CalibreBackupPath -Force | Out-Null
    }

    if (Test-Path -Path $script:CalibreBackupPath -PathType Container) {
        Write-Log -Message "$script:CalibreBackupPath was verified." -LogLevel "Info"
    } else {
        Write-Log -Message "The path $script:CalibreBackupPath does not exist or is not accessible." -LogLevel "Error"
        throw "The path $script:CalibreBackupPath does not exist or is not accessible."
    }
}

function Set-CalibreFolderPath {
    <#
    Determines CalibreFolder from the HiDrive sync root (via Get-HiDriveSyncRoot) instead of hostname matching.
    #>
    $HiDriveSyncRoot = Get-HiDriveSyncRoot
    if (-not $HiDriveSyncRoot) {
        Write-Log -Message "Could not determine HiDrive sync root. CalibreFolder not set." -LogLevel "Error"
        throw "Could not determine HiDrive sync root."
    }

    $script:CalibreFolder = Join-Path -Path $HiDriveSyncRoot -ChildPath "PortableApps\Calibre Portable"
    Write-Log -Message "Calibre portable path was set to $script:CalibreFolder" -LogLevel "Info"

    if (Test-Path -Path $script:CalibreFolder -PathType Container) {
        Write-Log -Message "$script:CalibreFolder was verified." -LogLevel "Info"
    } else {
        Write-Log -Message "The path $script:CalibreFolder does not exist or is not accessible." -LogLevel "Error"
        throw "The path $script:CalibreFolder does not exist or is not accessible."
    }
}

function Get-CalibreUpdate {
    # Verify if the update file from a previous update still exists and delete it if it does
    if (Test-Path -Path $CalibreInstaller -PathType Leaf) {
        Write-Log -Message "Calibre update file from previous update found at $CalibreInstaller. Deleting." -LogLevel "Info"
        Remove-Item -Path $CalibreInstaller -Force
    }
    else {
        Write-Log -Message "No Calibre update file found from a previous download." -LogLevel "Info"
    }    

    # Attempt to download the file
    Write-Log -Message "Downloading $CalibreUpdateSource to $CalibreInstaller" -LogLevel "Info"
    Start-BitsTransfer -Priority Foreground -Source $CalibreUpdateSource -Destination $CalibreInstaller -ErrorAction Stop

    # Verify if the file was downloaded successfully
    if (Test-Path -Path $CalibreInstaller -PathType Leaf) {
        Write-Log -Message "Calibre update downloaded successfully to $CalibreInstaller" -LogLevel "Info"
    }
    else {
        Write-Log -Message "Calibre update file not found at $CalibreInstaller after download attempt." -LogLevel "Error"
        throw "Download failed: Update file is missing."
    }
}

function New-CalibreBackup {
    if (Test-Path -Path $7zipPath -PathType Leaf) {
        Write-Log -Message "7zip found at $7zipPath, starting backup." -LogLevel "Info"
        <#
        https://7ziphelp.com/7zip-command-line
        a - create archive
        mx9 - maximum compression
        v1g - volume / file split after 1 GB
        bsp - verbose activity (progress) stream 
        #>
        $backupProcess = Start-Process -FilePath "$7zipPath" -ArgumentList "a -mx9 -bsp2 -v1g `"$CalibreBackupPath\CalibrePortableBackup_$Date`" `"$CalibreFolder`"" -Wait -NoNewWindow -PassThru
        if ($backupProcess.ExitCode -ne 0 -and $backupProcess.ExitCode -ne 1) {
            throw "7-Zip backup failed with exit code $($backupProcess.ExitCode)."
        }
        Write-Log -Message "7-Zip backup finished." -LogLevel "Info"
    }
    else {
        Write-Log -Message "7zip not found at $7zipPath. Stopping Script." -LogLevel "Error"
        throw "7-Zip binary not found at $7zipPath."
    }
}

function Install-CalibreUpdate {
    # Check if the Calibre process is running and stop it to allow the update to be installed
    if (Get-Process -Name "calibre*", "calibre-parallel*", "ebook-viewer*", "ebook-edit*" -ErrorAction SilentlyContinue) {
        Write-Log -Message "Calibre process is running. Stopping Calibre before update." -LogLevel "Info"
        Stop-Process -Name "calibre*", "calibre-parallel*", "ebook-viewer*", "ebook-edit*" -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    }
    else {
        Write-Log -Message "Calibre process is not running. Proceeding with update." -LogLevel "Info"
    }
    
    Write-Log -Message "Calibre update in $CalibreInstaller will be applied to $CalibreFolder" -LogLevel "Info"
    $installProcess = Start-Process -FilePath "$CalibreInstaller" -ArgumentList "`"$CalibreFolder`"" -Wait -PassThru
    if ($installProcess.ExitCode -ne 0) {
        Write-Log -Message "Calibre installer exited with code $($installProcess.ExitCode)." -LogLevel "Warning"
    }
    
    if (Test-Path -Path $CalibreInstaller -PathType Leaf) {
        Write-Log -Message "Deleting update file $CalibreInstaller" -LogLevel "Info"
        Remove-Item -Path $CalibreInstaller -Force
    }
}

function Remove-ExpiredBackups {
    Write-Log -Message "Cleanup of old backups in $CalibreBackupPath" -LogLevel "Info"

    # List all files in $CalibreBackupPath
    $files = Get-ChildItem -Path $CalibreBackupPath -Filter "CalibrePortableBackup_*.7z.*"

    if (-not $files) {
        Write-Log -Message "No backup files found in $CalibreBackupPath." -LogLevel "Info"
        return
    }

    # Group and sort files by full date and time
    $groupedFiles = $files | Group-Object {
        $_.BaseName -replace "CalibrePortableBackup_([0-9]{4}-[0-9]{2}-[0-9]{2}_[0-9]{2}-[0-9]{2}).*", '$1'
    } | Sort-Object {
        try {
            [datetime]::ParseExact($_.Name, "yyyy-MM-dd_HH-mm", [System.Globalization.CultureInfo]::InvariantCulture)
        }
        catch {
            [datetime]::MinValue
        }
    }

    # Delete all files older than the specified number in $CalibreBackupRetention
    if ($groupedFiles.Count -gt $CalibreBackupRetention) {
        $groupedFiles | Select-Object -First ($groupedFiles.Count - $CalibreBackupRetention) | ForEach-Object {
            $_.Group | ForEach-Object {
                Write-Log -Message "Deleting $($_.FullName)" -LogLevel "Info"
                Remove-Item -Path $_.FullName -Force
            }
        }
    } else {
        Write-Log -Message "No old backups to delete. Only $($groupedFiles.Count) backup set(s) found (retention: $CalibreBackupRetention)." -LogLevel "Info"
    }
}

#EndRegion

#Region: Main script execution
$hiDriveStopped = $false
try {
    Write-Log -Message "=============== Starting script ===============" -LogLevel "Info"
    Initialize-StratoHiDriveUtils
    Set-CalibreBackupPath
    Set-CalibreFolderPath
    Get-CalibreUpdate
    Stop-HiDrive
    $hiDriveStopped = $true
    New-CalibreBackup
    Install-CalibreUpdate
    Start-HiDrive
    $hiDriveStopped = $false
    Remove-ExpiredBackups
    Write-Log -Message "=============== Script completed ===============" -LogLevel "Info"
}
catch {
    Write-Log -Message "Error encountered: $($_.Exception.Message)" -LogLevel "Error"
    Write-Log -Message "StackTrace: $($_.Exception.StackTrace)" -LogLevel "Error"
}
finally {
    if ($hiDriveStopped) {
        Write-Log -Message "Restarting HiDrive after interruption/failure..." -LogLevel "Info"
        try {
            Start-HiDrive
            Write-Log -Message "HiDrive restarted successfully." -LogLevel "Info"
        }
        catch {
            Write-Log -Message "Failed to restart HiDrive: $($_.Exception.Message)" -LogLevel "Error"
        }
    }
}
#EndRegion