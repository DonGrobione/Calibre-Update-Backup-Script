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
Set-Alias Start-SevenZip $7zipPath -Scope script

# Define Date sting in YYYY-MM-DD format for filename
$Date = (Get-Date).ToString("yyyy-MM-dd")

# Define number of backup datasets to be kept in $CalibreBackup. Only the latest n set will be kept.
$CalibreBackupRetention = "3"

<#
Function that will change CalibreBackupPath depending on the hostname.
Change env:COMPUTERNAME to the hostname of your host and CalibreBackup to the path where the backup will be saved.
$CalibreBackupPath = "$null"
#>
$CalibreBackupPath = "$null"
if ($env:COMPUTERNAME -match "DONGROBIONE-PC") {
    Set-Variable CalibreBackupPath -Value "D:\HiDrive\Backup\Calibre\"
    Write-Log -Message "Calibe backups found in $CalibreBackupPath" -LogLevel "Info"
}
elseif ($env:COMPUTERNAME -match "DESKTOP-GS7HB29") {
    Set-Variable -Name CalibreBackupPath -Value "E:\HiDrive\Backup\Calibre\"
    Write-Log -Message "Calibe backups found in $CalibreBackupPath" -LogLevel "Info"
}
else {
    Write-Log -Message "Hostname $env:COMPUTERNAME not configured. CalibreBackupPath not set." -LogLevel "Error"
    Start-Sleep -Seconds 5
    Exit-PSSession 
}

function CalibreUpdateDownload {
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
    # Check if OneDrive process is running and stop it. This will prevent errors during Calibre Update.
    $OneDriveProcess = Get-Process -Name "OneDrive" -ErrorAction SilentlyContinue
    if ($OneDriveProcess) {
        Write-Log -Message "Stopping OneDrive process." -LogLevel "Info"
        Stop-Process -Name "OneDrive" -Force -Verbose
    }
    else {
        Write-Log -Message "OneDrive not running." -LogLevel "Error"
    }

    Write-Log -Message "Starting Calibre Update $CalibreInstaller" -LogLevel "Info"
    Set-Alias Start-CalibreUpdateExe $CalibreInstaller
    Start-CalibreUpdateExe $CalibreFolder
    Start-Sleep -Seconds 40
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
            Write-Log -Message "Onedrive found in $OneDrivePath" -LogLevel "Info"
            break
        }
    }

    # Start OneDrive
    Write-Log -Message "Starting OneDrive in $OneDrivePath" -LogLevel "Info"
    Start-Process -FilePath $OneDrivePath
}

## Execution
try {
    Clear-Host
    Write-Log -Message "Starting script." -LogLevel "Info"
    CalibreUpdateDownload
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