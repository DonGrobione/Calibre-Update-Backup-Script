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
    Created by DonGrobione
    Latest version can be found at https://github.com/DonGrobione/Calibre-Update-Backup-Script
#>

# Start PS logging
Start-Transcript -Path "$env:TEMP\Calibre-Backup-Update.log" -IncludeInvocationHeader

##  Definition of variables, change as needed
# Path to Calibre Portable in my OneDrive
New-Variable -Name CalibreFolder -Value "$env:OneDrive\PortableApps\Calibre Portable" -Scope script

# Calibre Update URL
New-Variable -Name CalibreUpdateSource -Value "https://calibre-ebook.com/dist/portable" -Scope script

# Definition where the the update file will be downloaded to
New-Variable -Name CalibreInstaller -Value "$env:TEMP\calibre-portable-installer.exe" -Scope script

# 7zip binariy
New-Variable -Name 7zipPath -Value "$env:ProgramFiles\7-Zip\7z.exe" -Scope script
Set-Alias Start-SevenZip $7zipPath -Scope script

# Define Date sting in YYYY-MM-DD format for filename
New-Variable -Name Date -Value (Get-Date).ToString("yyyy-MM-dd") -Scope script

# Define number of backup datasets to be kept in $CalibreBackup. Only the latest n set will be kept.
New-Variable -Name CalibreBackupRetention -Value "3" -Scope script

# Define potential OneDrive installation paths
New-Variable -Name "OneDrivePotentialPaths" -Value @(
    "${env:ProgramFiles}\Microsoft OneDrive\OneDrive.exe",
    "${env:ProgramFiles(x86)}\Microsoft OneDrive\OneDrive.exe",
    "${env:LocalAppData}\Microsoft\OneDrive\OneDrive.exe",
    "${env:WinDir}\SysWOW64\OneDriveSetup.exe"
) -Scope script

# Initialize OneDrivePath variable
New-Variable -Name "OneDrivePath" -Value $null -Scope script

# Check each potential Onedrive path and define OneDrivePath
Write-Output "Checking for OneDrive installation"
foreach ($path in $OneDrivePotentialPaths) {
    if (Test-Path $path) {
        Set-Variable -Name "OneDrivePath" -Value "$path"
        Write-Output "Onedrive found in $OneDrivePath"
        break
    }
}

## Functions
function DefineBackupPath {
    <#
    Function that will change CalibreBackupPath depending on the hostname.
    Change env:COMPUTERNAME to the hostnam of your host and CalibreBackup to the path where the backup will be saved.
    #>
    New-Variable -Name CalibreBackupPath
    if ($env:COMPUTERNAME -match "DONGROBIONE-PC") {
        Set-Variable CalibreBackupPath -Value "D:\HiDrive\HiDrive\Backup\Calibre\"
        Write-Output "Calibe found in $CalibreBackupPath"
    }
    elseif ($env:COMPUTERNAME -match "DESKTOP-GS7HB29") {
        Set-Variable -Name CalibreBackupPath -Value "E:\HiDrive\Backup\Calibre\"
        Write-Output "Calibe found in $CalibreBackupPath"
    }
    else {
        Write-Output "Hostname $env:COMPUTERNAME not configured."
        Write-Output "CalibreBackupPath not set."
        Start-Sleep -Seconds 5
        Exit-PSSession 
    }    
}

function CalibreUpdateDownload {
    Write-Output "Starting download from $CalibreUpdateSource to $CalibreInstaller"
    Start-BitsTransfer -Source $CalibreUpdateSource -Destination $CalibreInstaller -Priority Foreground
}

function CalibreBackup {
    if (Test-Path -Path $7zipPath -PathType Leaf) {
        Write-Output "7zip found in $7zipPath, starting backup"
        <#
        a - create archive
        mx9 - maximum compression
        v1g - volume / file split after 1 GB
        bsp - verboste activity stream 
        #>
        Write-Output "Creating Backups $CalibreBackupPath\CalibrePortableBackup_$Date."
        Start-SevenZip a -mx9 -v1g -bsp2 "$CalibreBackupPath\CalibrePortableBackup_$Date" $CalibreFolder
    }
    else {
        Write-Output "7zip installation path not found"
        Start-Sleep -Seconds 5
        #Exit-PSSession
        break
    }    
}

function CalibreUpdate {
    Write-Output "Starting Calibre Update $CalibreInstaller"
    Set-Alias Start-CalibreUpdateExe $CalibreInstaller
    Start-CalibreUpdateExe $CalibreFolder
}

function UpdateCleanup {
    # Deleteing update file
    Write-Output "Deleting update file $CalibreInstaller"
    Remove-Item -Path $CalibreInstaller
}

function BackupCleanup {
    Write-Output "Cleanup of old backups in $CalibreBackupPath"
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
                Write-Output "Deleting old backup files:"
                Write-Output "$_.FullName"
                Remove-Item -Path $_.FullName -Force
            }
        }
    }
}

function VarDebug {
    Write-Output "Date: $Date"
    Write-Output "CalibreFolder: $CalibreFolder"
    Write-Output "CalibreInstaller: $CalibreInstaller"
    Write-Output "CalibreBackupPath: $CalibreBackupPath"
    Write-Output "CalibreUpdateSource: $CalibreUpdateSource"
    Write-Output "7zipPath: $7zipPath"
    Write-Output "COMPUTERNAME: $env:COMPUTERNAME"
    Write-Output "CalibreBackupRetention: $CalibreBackupRetention"
    Write-Output "OneDrivePotentialPaths: $OneDrivePotentialPaths"
    Write-Output "OneDrivePath: $OneDrivePath"
    Start-Sleep -Seconds 5

}function OneDriveStop {
    # If OneDrivePath is not null, stop OneDrive
    if ($null -ne $OneDrivePath) {
        Write-Output "OneDrive found at $OneDrivePath. Stopping OneDrive Service."
        Stop-Process -Name "OneDrive"
    }
    else {
        Write-Output "OneDrive not found in any of the potential paths."
    }
}

function OneDriveStart {
    # Start OneDrive
    Write-Output "Starting OneDrive in $OneDrivePath"
    Start-Process -FilePath $OneDrivePath
}

## Execution
# VarDebug is commented out by default, as it is used to check if variables are set correctly
Clear-Host
DefineBackupPath
CalibreUpdateDownload
CalibreBackup
OneDriveStop
CalibreUpdate
OneDriveStart
UpdateCleanup
BackupCleanup

#Stop PS Logging
Stop-Transcript