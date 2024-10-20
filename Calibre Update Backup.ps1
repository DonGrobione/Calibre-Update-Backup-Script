<#
.SYNOPSIS
    This script will check the hostname and depening on it, will change the path where the backups will be saved.
    It will assume your Library is a subfolder in Calibre Portable and compress everything using 7zip in 1 GB archives.
    Update will be downloaded in tmp and applied to Calibre Portable.
    Finally the update file will be deleted and then the script will check past backups and only keeps the latest 3.
    To prevent errors during update, OneDrive will be temporarily stopped.

.DESCRIPTION
This file is the script I use myself, hence you will need to change a few things around. Especially the function DefineBackupPath and the variable CalibreFolder.

.NOTES
    Latest version can be found at https://github.com/DonGrobione/Calibre-Update-Backup-Script
    
    Log events should look like this:
    Write-Log -Message "This is an info level message." -LogLevel "Info"
    Write-Log -Message "This is an error level message." -LogLevel "Error"
#>

##  Definition of variables, change as needed

# Calibre Update URL
$CalibreUpdateSource = "https://calibre-ebook.com/dist/portable"

# Definition where the the update file will be downloaded to
$CalibreInstaller = "$env:TEMP\calibre-portable-installer.exe"

# 7zip binariy
$7zipPath = "$env:ProgramFiles\7-Zip\7z.exe"

# Define Date sting in YYYY-MM-DD format for filename
$Date = (Get-Date).ToString("yyyy-MM-dd_MM-mm")

# Define number of backup datasets to be kept in $CalibreBackup folder and used in Remove-ExpiredBackups. Only the latest n set will be kept.
$CalibreBackupRetention = 3

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
function Set-CalibreBackupPath {
    <#
    Function that will change CalibreBackupPath depending on the hostname.
    Change env:COMPUTERNAME to the hostname of your host and CalibreBackup to the path where the backup will be saved.
    #>
    $script:CalibreBackupPath = $null
    if ($env:COMPUTERNAME -match "DONGROBIONE-PC") {
        $script:CalibreBackupPath = "D:\HiDrive\Backup\Calibre\"
        Write-Log -Message "Calibre backups path was set to $script:CalibreBackupPath" -LogLevel "Info"
    }
    elseif ($env:COMPUTERNAME -match "DESKTOP-GS7HB29") {
        $script:CalibreBackupPath = "E:\HiDrive\Backup\Calibre\"
        Write-Log -Message "Calibre backups path was set to $script:CalibreBackupPath" -LogLevel "Info"
    }
    else {
        Write-Log -Message "Hostname $env:COMPUTERNAME not configured. CalibreBackupPath not set." -LogLevel "Error"
        Start-Sleep -Seconds 5
        exit 1
    }

    # Check if $script:CalibreBackupPath was set correctly
    if (Test-Path -Path $script:CalibreBackupPath -PathType Container) {
        Write-Log -Message "$script:CalibreBackupPath was verified." -LogLevel "Info"
    } else {
        Write-Log -Message "The path $script:CalibreBackupPath does not exist or is not accessible." -LogLevel "Error"
        exit 1
    }
}

function Set-CalibreFolderPath {
    <#
    Function that will change CalibreFolder depending on the hostname.
    Change env:COMPUTERNAME to the hostname of your host and CalibreFolder to the path where the backup will be saved.
    #>
    $script:CalibreFolder = $null
    if ($env:COMPUTERNAME -match "DONGROBIONE-PC") {
        $script:CalibreFolder = "D:\HiDrive\PortableApps\Calibre Portable"
        Write-Log -Message "Calibre portable path was set to $script:CalibreFolder" -LogLevel "Info"
    }
    elseif ($env:COMPUTERNAME -match "DESKTOP-GS7HB29") {
        $script:CalibreFolder = "E:\HiDrive\PortableApps\Calibre Portable"
        Write-Log -Message "Calibre portable path was set to $script:CalibreFolder" -LogLevel "Info"
    }
    else {
        Write-Log -Message "Hostname $env:COMPUTERNAME not configured. CalibreBackupPath not set." -LogLevel "Error"
        Start-Sleep -Seconds 5
        exit 1
    }

    # Check if $script:CalibreFolder was set correctly
    if (Test-Path -Path $script:CalibreFolder -PathType Container) {
        Write-Log -Message "$script:CalibreFolder was verified." -LogLevel "Info"
    } else {
        Write-Log -Message "The path $script:CalibreFolder does not exist or is not accessible." -LogLevel "Error"
        exit 1
    }
}

function Get-CalibreUpdate {
    Write-Log -Message "Starting download from $CalibreUpdateSource to $CalibreInstaller" -LogLevel "Info"
    Start-BitsTransfer -Source $CalibreUpdateSource -Destination $CalibreInstaller -Priority Foreground
}

function New-CalibreBackup {
    # Create a new directory for the backup, using the date as the folder name
    $backupFolder = "$script:CalibreBackupPath\$Date"
    
    # Check if the backup folder already exists, if not, create it
    if (-not (Test-Path -Path $backupFolder)) {
        Write-Log -Message "Creating backup folder: $backupFolder" -LogLevel "Info"
        New-Item -Path $backupFolder -ItemType Directory
    } else {
        Write-Log -Message "Backup folder $backupFolder already exists." -LogLevel "Warning"
    }

    # Now check if 7zip exists and proceed with the backup
    if (Test-Path -Path $7zipPath -PathType Leaf) {
        Write-Log -Message "7zip found at $7zipPath, starting backup." -LogLevel "Info"
        <#
        a - create archive
        mx9 - maximum compression
        v1g - volume / file split after 1 GB
        bsp - verbose activity stream 
        #>
        Start-Process -FilePath "$7zipPath" -ArgumentList "a -mx9 -bsp2 -v1g `"$backupFolder\CalibrePortableBackup_$Date`" `"$CalibreFolder`"" -Wait -NoNewWindow
    }
    else {
        Write-Log -Message "7zip not found at $7zipPath" -LogLevel "Info"
        Write-Log -Message "Stopping Script." -LogLevel "Info"
        exit 1
    }
}

function Install-CalibreUpdate {
    # Install the update and reset exit code for Calibre Update
    $global:LASTEXITCODE = $null
    Write-Log -Message "Calibre update in $CalibreInstaller will be applied to $CalibreFolder" -LogLevel "Info"
    Start-Process -FilePath "$CalibreInstaller" -ArgumentList `"$CalibreFolder`" -Wait
    # Check the exit code for successful installation
    if ($LASTEXITCODE -eq 0) {
        Write-Log -Message "Calibre has been successfully updated." -LogLevel "Info"
        Write-Log -Message "Deleting update file $CalibreInstaller" -LogLevel "Info"
        Remove-Item -Path $CalibreInstaller
    } else {
        Write-Log -Message "Calibre installation failed with exit code: $LASTEXITCODE" -LogLevel "Error"
        exit 1
    }
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
                Write-Log -Message "Deleting old backup files:" -LogLevel "Info"
                Write-Log -Message "$($_.FullName)" -LogLevel "Info"
                Remove-Item -Path $_.FullName -Force
            }
        }
    } else {
        Write-Log -Message "No old backups to delete. Only $($groupedFiles.Count) backups found." -LogLevel "Info"
    }
}

function Start-HiDrive {
    # Define potential HiDrive installation paths
    $HiDrivePotentialPaths = @(
        "${env:ProgramFiles}\STRATO\HiDrive\HiDrive.App.exe",
        "${env:ProgramFiles(x86)}\STRATO\HiDrive\HiDrive.App.exe",
        "${env:LocalAppData}\STRATO\HiDrive\HiDrive.App.exe"
    )

    # Initialize $HiDrivePath variable
    $HiDrivePath = $null

    # Check each potential HiDrive path and define HiDrivePath
    Write-Log -Message "Checking for HiDrive installation." -LogLevel "Info"
    foreach ($path in $HiDrivePotentialPaths) {
        if (Test-Path $path) {
            $HiDrivePath = $path
            Write-Log -Message "HiDrive found in $HiDrivePath" -LogLevel "Info"
            break
        }
    }

    # Check if HiDrivePath was found and start HiDrive
    if ($HiDrivePath) {
        Write-Log -Message "Starting HiDrive in $HiDrivePath" -LogLevel "Info"
        Start-Process -FilePath $HiDrivePath
    } else {
        Write-Log -Message "HiDrive not found on this system. Skipping start." -LogLevel "Error"  # Handling case where no HiDrive path is found
    }
}

function Stop-HiDrive {
    # Check if the HiDrive process is running and stop it to prevent sync errors
    $process = Get-Process -Name "HiDrive.App" -ErrorAction SilentlyContinue
    if ($process) {
        # If the process is running, stop it
        Write-Log -Message "HiDrive process is running. Stopping." -LogLevel "Info"
        Stop-Process -Name $process.Name -Force

        # Check every second if the process has stopped
        while ($true) {
            Start-Sleep -Seconds 1
            $process = Get-Process -Name "HiDrive.App" -ErrorAction SilentlyContinue

            if (-not $process) {
                Write-Log -Message "HiDrive process has been successfully stopped. Proceeding with the script." -LogLevel "Info"
                break
            } else {
                Write-Log -Message "Waiting for HiDrive to stop." -LogLevel "Info"
            }
        }
    } else {
        # If the process is not running, proceed with the rest of the script
        Write-Log -Message "HiDrive process is not running. Proceeding with the script." -LogLevel "Info"
    }
}

## Execution
try {
    Write-Log -Message "Starting script." -LogLevel "Info"
    Set-CalibreBackupPath
    Set-CalibreFolderPath
    Get-CalibreUpdate
    Stop-HiDrive
    New-CalibreBackup
    Install-CalibreUpdate
    Start-HiDrive
    Remove-ExpiredBackups
    Write-Log -Message "Script completed." -LogLevel "Info"
}
catch {
    <#Do this if a terminating exception happens#>
    Write-Log -Message "Error encountered: $($_.Exception.Message)" -LogLevel "Error"
    Write-Log -Message "StackTrace: $($_.Exception.StackTrace)" -LogLevel "Error"
}