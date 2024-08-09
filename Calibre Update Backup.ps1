<#
.SYNOPSIS
    This scrtipt will check the hostname and depening on it, will chnage the path where the backups will be saved.
    It will asume your Library is a subfolder in Calibre Portable and compress everything using 7zip in 1 GB archives.
    Update will be downloaded in tmp and applied to Calibre Portable.
    Finally the update file will be deleted and then the script will check past backups and only keps the latest 3.
    I use this script on multiple hosts with differing folder structurs, hence the check for hostname to set variables.

.DESCRIPTION
This file is the script I use myself, hence you will need to change a few things around. Especially the function DefineBackupPath and the variable CalibreFolder.

.NOTES
    Created by DonGrobione
    Latest version can be found at https://github.com/DonGrobione/Calibre-Update-Backup-Script
#>

# Start PS logging
Start-Transcript -Path "$PSScriptRoot\Calibre-Backup-Update.log" -IncludeInvocationHeader

##  Definition of variables, change as needed
# Path to Calibre Portable, will be set depending on hostname
New-Variable -Name CalibreFolder -Value $null -Scope script

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

# Variable for the path the backup files,  will be set depending on hostname
New-Variable -Name CalibreBackupPath -Value $null -Scope script

<#
Function that will change CalibreBackupPath depending on the hostname.
Change env:COMPUTERNAME to the hostnam of your host and CalibreBackup to the path where the backup will be saved.
#>
if ($env:COMPUTERNAME -match "DONGROBIONE-PC") {
    Set-Variable CalibreBackupPath -Value "D:\HiDrive\Backup\Calibre\"
    Write-Output "Calibe backups found in $CalibreBackupPath"
    Set-Variable CalibreFolder -Value "D:\HiDrive\PortableApps\Calibre Portable"
    Write-Output "Calibe portable found in $CalibreFolder"
}
elseif ($env:COMPUTERNAME -match "DESKTOP-GS7HB29") {
    Set-Variable -Name CalibreBackupPath -Value "E:\HiDrive\Backup\Calibre\"
    Write-Output "Calibe backups found in $CalibreBackupPath"
    Set-Variable CalibreFolder -Value "E:\HiDrive\PortableApps\Calibre Portable"
    Write-Output "Calibe portable found in $CalibreFolder"
}
else {
    Write-Output "Hostname $env:COMPUTERNAME not configured."
    Write-Output "Host specific variables could not be set."
    Start-Sleep -Seconds 5
    Exit-PSSession 
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
    Start-Sleep -Seconds 40
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

## Execution
Clear-Host
CalibreUpdateDownload
CalibreBackup
CalibreUpdate
UpdateCleanup
BackupCleanup

#Stop PS Logging
Stop-Transcript