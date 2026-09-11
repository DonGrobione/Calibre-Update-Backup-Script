<#
.SYNOPSIS
    Backs up Calibre Portable, downloads and installs the latest portable update,
    and removes old backup sets based on retention.

.DESCRIPTION
    The script installs or updates the StratoHiDriveUtils module (https://github.com/DonGrobione/StratoHiDriveUtils) via git,
    uses it to resolve the HiDrive sync root and derive the Calibre installation and backup paths,
    downloads the current Calibre Portable installer to the TEMP folder,
    stops HiDrive to avoid sync/file lock issues during backup and update,
    creates a split 7z backup archive, installs the update, restarts HiDrive,
    and then deletes expired backups.

.EXAMPLE
    .\Calibre Update Backup.ps1

    Runs the full backup-update-cleanup workflow with automatic HiDrive path resolution.

.NOTES
    Version: 2.0.0
    Updated: 2026-09-11
    Mail: dongrobione@proton.me
    Latest version: https://github.com/DonGrobione/Calibre-Update-Backup-Script
    Requires: git in PATH, StratoHiDriveUtils module (https://github.com/DonGrobione/StratoHiDriveUtils)

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
function Install-StratoHiDriveUtilsModule {
    <#
    Ensures the StratoHiDriveUtils module (https://github.com/DonGrobione/StratoHiDriveUtils) is installed via git clone.
    If already installed, checks GitHub for a newer version and pulls it if available. Imports the module afterwards.
    #>
    $ModuleName = "StratoHiDriveUtils"
    $ModuleRepo = "https://github.com/DonGrobione/StratoHiDriveUtils.git"
    $ModulePath = Join-Path -Path ($env:PSModulePath -split ';')[0] -ChildPath $ModuleName

    if (-not (Get-Command -Name git -ErrorAction SilentlyContinue)) {
        Write-Log -Message "git was not found. Cannot install or update $ModuleName module." -LogLevel "Error"
        exit 1
    }

    if (-not (Test-Path -Path (Join-Path -Path $ModulePath -ChildPath "$ModuleName.psd1"))) {
        Write-Log -Message "$ModuleName module not found at $ModulePath. Cloning from $ModuleRepo." -LogLevel "Info"
        New-Item -ItemType Directory -Path $ModulePath -Force | Out-Null
        git clone --quiet $ModuleRepo $ModulePath 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Write-Log -Message "Failed to clone $ModuleName module from $ModuleRepo." -LogLevel "Error"
            exit 1
        }
        Write-Log -Message "$ModuleName module installed successfully." -LogLevel "Info"
    }
    else {
        Write-Log -Message "$ModuleName module found at $ModulePath. Checking GitHub for a newer version." -LogLevel "Info"
        git -C $ModulePath fetch --quiet 2>&1 | Out-Null
        $LocalCommit = git -C $ModulePath rev-parse HEAD
        $RemoteCommit = git -C $ModulePath rev-parse '@{u}'
        if ($LocalCommit -ne $RemoteCommit) {
            Write-Log -Message "A newer version of $ModuleName is available on GitHub. Pulling update." -LogLevel "Info"
            git -C $ModulePath pull --quiet 2>&1 | Out-Null
            Remove-Module -Name $ModuleName -Force -ErrorAction SilentlyContinue
            Write-Log -Message "$ModuleName module updated successfully." -LogLevel "Info"
        }
        else {
            Write-Log -Message "$ModuleName module is already up to date." -LogLevel "Info"
        }
    }

    Import-Module -Name $ModuleName -Force -ErrorAction Stop
    Write-Log -Message "$ModuleName module imported." -LogLevel "Info"
}

function Set-CalibreBackupPath {
    <#
    Determines CalibreBackupPath from the HiDrive sync root (via Get-HiDriveSyncRoot) instead of hostname matching.
    #>
    $HiDriveSyncRoot = Get-HiDriveSyncRoot
    if (-not $HiDriveSyncRoot) {
        Write-Log -Message "Could not determine HiDrive sync root. CalibreBackupPath not set." -LogLevel "Error"
        exit 1
    }

    $script:CalibreBackupPath = Join-Path -Path $HiDriveSyncRoot -ChildPath "Backup\Calibre"
    Write-Log -Message "Calibre backups path was set to $script:CalibreBackupPath" -LogLevel "Info"

    if (Test-Path -Path $script:CalibreBackupPath -PathType Container) {
        Write-Log -Message "$script:CalibreBackupPath was verified." -LogLevel "Info"
    } else {
        Write-Log -Message "The path $script:CalibreBackupPath does not exist or is not accessible." -LogLevel "Error"
        exit 1
    }
}

function Set-CalibreFolderPath {
    <#
    Determines CalibreFolder from the HiDrive sync root (via Get-HiDriveSyncRoot) instead of hostname matching.
    #>
    $HiDriveSyncRoot = Get-HiDriveSyncRoot
    if (-not $HiDriveSyncRoot) {
        Write-Log -Message "Could not determine HiDrive sync root. CalibreFolder not set." -LogLevel "Error"
        exit 1
    }

    $script:CalibreFolder = Join-Path -Path $HiDriveSyncRoot -ChildPath "PortableApps\Calibre Portable"
    Write-Log -Message "Calibre portable path was set to $script:CalibreFolder" -LogLevel "Info"

    if (Test-Path -Path $script:CalibreFolder -PathType Container) {
        Write-Log -Message "$script:CalibreFolder was verified." -LogLevel "Info"
    } else {
        Write-Log -Message "The path $script:CalibreFolder does not exist or is not accessible." -LogLevel "Error"
        exit 1
    }
}

function Get-CalibreUpdate {
    # Verify if the update file froma previus update still exists and delete it if it does
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
        bsp - verboste activity (progress) stream 
        bse - error stream ?
        #>
        Start-Process -FilePath "$7zipPath" -ArgumentList "a -mx9 -bsp2 -v1g `"$CalibreBackupPath\CalibrePortableBackup_$Date`" `"$CalibreFolder`"" -Wait -NoNewWindow
    }
    else {
        Write-Log -Message "7zip not found at $7zipPath" -LogLevel "Info"
        Write-Log -Message "Stopping Script." -LogLevel "Info"
        exit 1
    }
}

function Install-CalibreUpdate {
    # Check if the Calibre process is running and stop it to allow the update to be installed
    if (Get-Process -Name "Calibre" -ErrorAction SilentlyContinue) {
        Write-Log -Message "Calibre process is running. Stopping Calibre before update." -LogLevel "Error"
        Stop-Process -Name "Calibre"
        Start-Sleep -Seconds 2
    }
    else {
        Write-Log -Message "Calibre process is not running. Proceeding with update." -LogLevel "Info"
    }
    
    # Install the update and reset exit code for Calibre Update
    $global:LASTEXITCODE = $null
    Write-Log -Message "Calibre update in $CalibreInstaller will be applied to $CalibreFolder" -LogLevel "Info"
    Start-Process -FilePath "$CalibreInstaller" -ArgumentList `"$CalibreFolder`" -Wait
    Write-Log -Message "Deleting update file $CalibreInstaller" -LogLevel "Info"
    Remove-Item -Path $CalibreInstaller
}

function Remove-ExpiredBackups {
    Write-Log -Message "Cleanup of old backups in $CalibreBackupPath" -LogLevel "Info"

    # List all files in $CalibreBackupPath
    $files = Get-ChildItem -Path $CalibreBackupPath -Filter "CalibrePortableBackup_*.7z.*"

    # Sort files by full date and time
    $sortedFiles = $files | Sort-Object {
        # Extract the full date and time from the filename (yyyy-MM-dd_HH-mm)
        $dateString = $_.BaseName -replace "CalibrePortableBackup_([0-9]{4}-[0-9]{2}-[0-9]{2}_[0-9]{2}-[0-9]{2}).*", '$1'
        [datetime]::ParseExact($dateString, "yyyy-MM-dd_HH-mm", $null)
    }

    # Group and sort files by full date and time
    $groupedFiles = $sortedFiles | Group-Object {
        $_.BaseName -replace "CalibrePortableBackup_([0-9]{4}-[0-9]{2}-[0-9]{2}_[0-9]{2}-[0-9]{2}).*", '$1'
    } | Sort-Object { [datetime]::ParseExact($_.Name, "yyyy-MM-dd_HH-mm", $null) }

    # Delete all files older than the specified number in $CalibreBackupRetention
    if ($groupedFiles.Count -gt $CalibreBackupRetention) {
        $groupedFiles | Select-Object -First ($groupedFiles.Count - $CalibreBackupRetention) | ForEach-Object {
            $_.Group | ForEach-Object {
                Write-Log -Message "Deleting $($_.FullName)" -LogLevel "Info"
                Remove-Item -Path $_.FullName -Force
            }
        }
    } else {
        Write-Log -Message "No old backups to delete. Only $($groupedFiles.Count) backups found." -LogLevel "Info"
    }
}

#EndRegion

#Region: Main script execution
try {
    Write-Log -Message "=============== Starting script ===============" -LogLevel "Info"
    Install-StratoHiDriveUtilsModule
    Set-CalibreBackupPath
    Set-CalibreFolderPath
    Get-CalibreUpdate
    Stop-HiDrive
    New-CalibreBackup
    Install-CalibreUpdate
    Start-HiDrive
    Remove-ExpiredBackups
    Write-Log -Message "=============== Script completed ===============" -LogLevel "Info"
}
catch {
    <#Do this if a terminating exception happens#>
    Write-Log -Message "Error encountered: $($_.Exception.Message)" -LogLevel "Error"
    Write-Log -Message "StackTrace: $($_.Exception.StackTrace)" -LogLevel "Error"
}
#EndRegion