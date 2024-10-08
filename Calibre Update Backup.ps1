<#
.SYNOPSIS
    This scrtipt will check the hostname and depening on it, will chnage the path where the backups will be saved.
    It will asume your Library is a subfolder in Calibre Portable and compress everything using 7zip in 1 GB archives.
    Update will be downloaded in tmp and applied to Calibre Portable.
    Finally the update file will be deleted and then the script will check past backups and only keps the latest 3.
    To prevent errors during update, OneDrive will be temporarly stopped.

.DESCRIPTION
This file is the script I use myself, hence you will need to change a few things around. Especially the function DefineBackupPath and the variable CalibreFolder.

.NOTES
    Latest version can be found at https://github.com/DonGrobione/Calibre-Update-Backup-Script
    
    Log events should look like this:
    Write-Log -Message "This is an info level mesage." -LogLevel "Info"
    Write-Log -Message "This is an error level mesage." -LogLevel "Error"
#>

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

##  Definition of variables, change as needed
# Path to Calibre Portable in my OneDrive
$CalibreFolder = "$env:OneDrive\PortableApps\Calibre Portable"

# Calibre Update URL
$CalibreUpdateSource = "https://calibre-ebook.com/dist/portable"

# Definition where the the update file will be downloaded to
$CalibreInstaller = "$env:TEMP\calibre-portable-installer.exe"

# 7zip binariy
$7zipPath = "$env:ProgramFiles\7-Zip\7z.exe"

# Define Date sting in YYYY-MM-DD format for filename
$Date = (Get-Date).ToString("yyyy-MM-dd")

# Define number of backup datasets to be kept in $CalibreBackup. Only the latest n set will be kept.
$CalibreBackupRetention = "3"


function Set-CalibreBackupPath {
    <#
    Function that will change CalibreBackupPath depending on the hostname.
    Change env:COMPUTERNAME to the hostname of your host and CalibreBackup to the path where the backup will be saved.
    #>
    $script:CalibreBackupPath = $null
    if ($env:COMPUTERNAME -match "DONGROBIONE-PC") {
        $script:CalibreBackupPath = "D:\HiDrive\Backup\Calibre\"
        Write-Log -Message "Calibe backups found in $CalibreBackupPath" -LogLevel "Info"
    }
    elseif ($env:COMPUTERNAME -match "DESKTOP-GS7HB29") {
        $script:CalibreBackupPath = "E:\HiDrive\Backup\Calibre\"
        Write-Log -Message "Calibe backups found in $CalibreBackupPath" -LogLevel "Info"
    }
    else {
        Write-Log -Message "Hostname $env:COMPUTERNAME not configured. CalibreBackupPath not set." -LogLevel "Error"
        Start-Sleep -Seconds 5
        Exit-PSSession
    }
}


function Get-CalibreUpdate {
    Write-Log -Message "Starting download from $CalibreUpdateSource to $CalibreInstaller" -LogLevel "Info"
    Start-BitsTransfer -Source $CalibreUpdateSource -Destination $CalibreInstaller -Priority Foreground
}

function CalibreBackup {
    if (Test-Path -Path $7zipPath -PathType Leaf) {
        Write-Log -Message "7zip found in $7zipPath, starting backup"
        <#
        a - create archive
        mx9 - maximum compression
        v1g - volume / file split after 1 GB
        bsp - verboste activity stream 
        #>
        Write-Log -Message "Creating Backups $CalibreBackupPath\CalibrePortableBackup_$Date." -LogLevel "Info"
        Start-SevenZip a -mx9 -v1g -bsp2 "$CalibreBackupPath\CalibrePortableBackup_$Date" $CalibreFolder
    }
    else {
        Write-Log -Message "7zip installation path not found" -LogLevel "Error"
        Start-Sleep -Seconds 5
        #Exit-PSSession
        break
    }    
}

function CalibreUpdate {
    # Check if the OneDrive process is running
    $process = Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue

    if ($process) {
        # If the process is running, stop it
        Write-Log -Message "OneDrive process is running. Stopping." -LogLevel "Info"
        Stop-Process -Name $process.Name -Force
        Start-Sleep -Seconds 5
    } else {
        # If the process is not running, proceed with the rest of the script
        Write-Log -Message "OneDrive process is not running. Proceeding with the script." -LogLevel "Info"
    }
    # Install the update
    Write-Log -Message "Calibre update in $CalibreInstaller will be applied to $CalibreFolder" -LogLevel "Info"
    Start-Process -FilePath "$CalibreInstaller" -ArgumentList `"$CalibreFolder`" -Wait
    # Check the exit code for successful installation
    if ($LASTEXITCODE -eq 0) {
        Write-Log -Message "Calibre has been successfully updated." -LogLevel "Info"
    } else {
        Write-Log -Message "Calibre installation failed with exit code: $LASTEXITCODE" -LogLevel "Error"
    }
}

function UpdateCleanup {
    # Deleteing update file
    Write-Log -Message "Deleting update file $CalibreInstaller" -LogLevel "Info"
    Remove-Item -Path $CalibreInstaller
}

function BackupCleanup {
    Write-Log -Message "Cleanup of old backups in $CalibreBackupPath" -LogLevel "Info"
    # List all files in $CalibreBackupPath
    $files = Get-ChildItem -Path $CalibreBackupPath -Filter "CalibrePortableBackup_*.7z.*"

    # Sort files by date
    $sortedFiles = $files | Sort-Object {
        # Extrahieren Sie das Datum aus dem Dateinamen
        $dateString = $_.BaseName -replace "CalibrePortableBackup_([0-9]{4}-[0-9]{2}-[0-9]{2}).*", '$1'
        [datetime]::ParseExact($dateString, "yyyy-MM-dd", $null)
    }

    # Group and sort files by date
    $groupedFiles = $sortedFiles | Group-Object {
        $_.BaseName -replace "CalibrePortableBackup_([0-9]{4}-[0-9]{2}-[0-9]{2}).*", '$1'
    } | Sort-Object { [datetime]::ParseExact($_.Name, "yyyy-MM-dd", $null) }

    # Delete all files older that $CalibreBackupRetention
    if ($groupedFiles.Count -gt $CalibreBackupRetention) {
        $groupedFiles | Select-Object -First ($groupedFiles.Count - 3) | ForEach-Object {
            $_.Group | ForEach-Object {
                Write-Log -Message "Deleting old backup files:"
                Write-Log -Message "$_.FullName"
                Remove-Item -Path $_.FullName -Force
            }
        }
    }
}

function OneDriveStart {
    # Define potential OneDrive installation paths
    $OneDrivePotentialPaths = @(
        "${env:ProgramFiles}\Microsoft OneDrive\OneDrive.exe",
        "${env:ProgramFiles(x86)}\Microsoft OneDrive\OneDrive.exe",
        "${env:LocalAppData}\Microsoft\OneDrive\OneDrive.exe",
        "${env:WinDir}\SysWOW64\OneDriveSetup.exe"
    )

    # Initialize OneDrivePath variable
    $OneDrivePath = $null

    # Check each potential Onedrive path and define OneDrivePath
    Write-Log -Message "Checking for OneDrive installation"
    foreach ($path in $OneDrivePotentialPaths) {
        if (Test-Path $path) {
            Set-Variable -Name "OneDrivePath" -Value "$path"
            Write-Log -Message "OneDrive found in $OneDrivePath" -LogLevel "Info"
            break
        }
    }

    # Start OneDrive
    Write-Log -Message "Starting OneDrive in $OneDrivePath" -LogLevel "Info"
    Start-Process -FilePath $OneDrivePath
}

## Execution
try {
    Write-Log -Message "Starting script." -LogLevel "Info"
    Set-CalibreBackupPath
    Get-CalibreUpdate
    CalibreBackup
    CalibreUpdate
    OneDriveStart
    UpdateCleanup
    BackupCleanup
}
catch {
    <#Do this if a terminating exception happens#>
    Write-Log -Message "Error encountered:" -LogLevel "Error"
    Write-Log -Message "$_.ScriptStackTrace" -LogLevel "Error"
}