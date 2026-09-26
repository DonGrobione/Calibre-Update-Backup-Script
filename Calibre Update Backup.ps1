<#PSScriptInfo

.VERSION 2.2.0

.GUID 9860f6ef-700a-42b6-a78b-d0e88895d44b

.AUTHOR DonGrobione

.COMPANYNAME DonGrobione

.COPYRIGHT (c) DonGrobione. Licensed under the GNU AGPL v3.

.TAGS Calibre HiDrive Backup Update 7-Zip Windows PSEdition_Desktop

.LICENSEURI https://github.com/DonGrobione/Calibre-Update-Backup-Script/blob/main/LICENSE.md

.PROJECTURI https://github.com/DonGrobione/Calibre-Update-Backup-Script

.ICONURI

.EXTERNALMODULEDEPENDENCIES DonGrobione.StratoHiDriveUtils, DonGrobione.Logging

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES
2.2.0: Adds parameters, -WhatIf support, exit codes, logging via DonGrobione.Logging, and automatic installation and update of the required modules. Switches the module self-update to Update-HiDriveUtility. See CHANGELOG.md for details.

.PRIVATEDATA

#>

#Requires -Version 5.1

<#
.SYNOPSIS
    Backs up Calibre Portable, downloads and installs the latest portable update, and removes old backup sets based on retention.

.DESCRIPTION
    The script first makes the required modules DonGrobione.Logging (https://github.com/DonGrobione/Logging) and DonGrobione.StratoHiDriveUtils (https://github.com/DonGrobione/StratoHiDriveUtils) available.
    A missing module is installed for the current user with its official installer, and an installed module is updated with its own update command.
    It resolves the HiDrive sync root, derives the Calibre installation and backup paths from it, and downloads the current Calibre Portable installer to the TEMP folder.
    It then stops HiDrive to avoid sync and file lock issues, creates a split 7z backup archive, installs the update, restarts HiDrive, and deletes expired backup sets.
    If a step fails while HiDrive is stopped, HiDrive is restarted before the script exits.
    Each run writes its own log file to <Documents>\Logs\Calibre-Update-Backup, and the five newest log files are kept.

.PARAMETER CalibreUpdateSource
    HTTPS URL of the Calibre Portable installer.

.PARAMETER SevenZipPath
    Full path to 7z.exe.

.PARAMETER CalibreBackupRetention
    Number of backup sets to keep. Older sets are deleted.

.EXAMPLE
    & '.\Calibre Update Backup.ps1'

    Runs the full backup, update, and cleanup workflow with automatic HiDrive path resolution.

.EXAMPLE
    & '.\Calibre Update Backup.ps1' -WhatIf

    Shows which state-changing steps would run without downloading, stopping HiDrive, backing up, installing, or deleting anything.

.EXAMPLE
    & '.\Calibre Update Backup.ps1' -CalibreBackupRetention 5

    Runs the workflow and keeps the five newest backup sets.

.NOTES
    Latest version: https://github.com/DonGrobione/Calibre-Update-Backup-Script
    Requires: DonGrobione.StratoHiDriveUtils module (https://github.com/DonGrobione/StratoHiDriveUtils), DonGrobione.Logging module (https://github.com/DonGrobione/Logging), 7-Zip, and the STRATO HiDrive client.
    Exit codes: 0 on success, 1 on failure.
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter()]
    [ValidatePattern('^https://')]
    [string]$CalibreUpdateSource = 'https://calibre-ebook.com/dist/portable',

    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$SevenZipPath = (Join-Path -Path $env:ProgramFiles -ChildPath '7-Zip\7z.exe'),

    [Parameter()]
    [ValidateRange(1, 100)]
    [int]$CalibreBackupRetention = 3
)

#Region: Functions
# Builds an ErrorRecord for $PSCmdlet.ThrowTerminatingError() without changing any system state.
function New-ErrorRecord {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseShouldProcessForStateChangingFunctions', '', Justification = 'Only creates an in-memory object.')]
    [CmdletBinding()]
    [OutputType([System.Management.Automation.ErrorRecord])]
    param(
        [Parameter(Mandatory = $true)]
        [System.Exception]$Exception,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ErrorId,

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.ErrorCategory]$Category,

        [Parameter()]
        [object]$TargetObject
    )

    return [System.Management.Automation.ErrorRecord]::new($Exception, $ErrorId, $Category, $TargetObject)
}

# Downloads a module's official installer script to the TEMP folder, runs it for the current user, and deletes it afterwards.
function Install-RequiredModule {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ModuleName,

        [Parameter(Mandatory = $true)]
        [ValidatePattern('^https://')]
        [string]$InstallerUri
    )

    if (-not $PSCmdlet.ShouldProcess($ModuleName, "Install for the current user with $InstallerUri")) {
        $exception = [System.InvalidOperationException]::new("$ModuleName module is not installed, and its installation was skipped.")
        $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -Exception $exception -ErrorId 'ModuleInstallSkipped' -Category NotInstalled -TargetObject $ModuleName))
    }

    $installerPath = Join-Path -Path $env:TEMP -ChildPath ('{0}-Install-{1}.ps1' -f $ModuleName, [guid]::NewGuid().ToString('N'))
    try {
        Write-Verbose -Message "Downloading $InstallerUri to $installerPath"
        Start-BitsTransfer -Source $InstallerUri -Destination $installerPath -ErrorAction Stop
        $null = & $installerPath
    }
    catch {
        $exception = [System.InvalidOperationException]::new("$ModuleName module could not be installed with $InstallerUri. $($_.Exception.Message)", $_.Exception)
        $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -Exception $exception -ErrorId 'ModuleInstallFailed' -Category NotInstalled -TargetObject $ModuleName))
    }
    finally {
        if (Test-Path -LiteralPath $installerPath -PathType Leaf) {
            Remove-Item -LiteralPath $installerPath -Force -ErrorAction SilentlyContinue
        }
    }
}

# Makes a required module available: installs it when it cannot be imported, otherwise runs its update command and reloads it. Returns a status object for the caller to log, because this runs before logging is available. A failed update is non-fatal as long as the installed version can still be imported.
function Initialize-RequiredModule {
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([pscustomobject])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ModuleName,

        [Parameter(Mandatory = $true)]
        [ValidatePattern('^https://')]
        [string]$InstallerUri,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$UpdateCommand
    )

    $module = $null
    try {
        $module = Import-Module -Name $ModuleName -PassThru -ErrorAction Stop | Select-Object -First 1
    }
    catch {
        Write-Verbose -Message "$ModuleName module could not be imported: $($_.Exception.Message)"
    }

    if (-not $module) {
        Install-RequiredModule -ModuleName $ModuleName -InstallerUri $InstallerUri
        $module = Import-Module -Name $ModuleName -PassThru -ErrorAction Stop | Select-Object -First 1
        return [pscustomobject]@{
            ModuleName  = $ModuleName
            Status      = 'Installed'
            Version     = $module.Version
            Message     = "$ModuleName module $($module.Version) was not installed and has been installed."
            ErrorRecord = $null
        }
    }

    if (-not $PSCmdlet.ShouldProcess($ModuleName, "Update with $UpdateCommand")) {
        return [pscustomobject]@{
            ModuleName  = $ModuleName
            Status      = 'UpdateSkipped'
            Version     = $module.Version
            Message     = "$ModuleName module $($module.Version) is installed. The update check was skipped."
            ErrorRecord = $null
        }
    }

    $installedVersion = $module.Version
    try {
        $null = & $UpdateCommand -Confirm:$false -ErrorAction Stop
    }
    catch {
        return [pscustomobject]@{
            ModuleName  = $ModuleName
            Status      = 'UpdateFailed'
            Version     = $installedVersion
            Message     = "$ModuleName module update failed. Continuing with version $installedVersion."
            ErrorRecord = $_
        }
    }

    # Reload so this session uses the newest installed version.
    $module = Import-Module -Name $ModuleName -Force -PassThru -ErrorAction Stop | Select-Object -First 1
    if ($module.Version -gt $installedVersion) {
        return [pscustomobject]@{
            ModuleName  = $ModuleName
            Status      = 'Updated'
            Version     = $module.Version
            Message     = "$ModuleName module updated from $installedVersion to $($module.Version)."
            ErrorRecord = $null
        }
    }
    return [pscustomobject]@{
        ModuleName  = $ModuleName
        Status      = 'UpToDate'
        Version     = $module.Version
        Message     = "$ModuleName module is up to date (version $($module.Version))."
        ErrorRecord = $null
    }
}

# Returns the Calibre backup directory below the HiDrive sync root and creates it when it does not exist.
function Initialize-CalibreBackupFolder {
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
        [string]$HiDriveSyncRoot
    )

    $backupPath = Join-Path -Path $HiDriveSyncRoot -ChildPath 'Backup\Calibre'
    Write-Log -Message "Calibre backup path was set to $backupPath"

    if (Test-Path -LiteralPath $backupPath -PathType Container) {
        Write-Log -Message "$backupPath was verified."
        return $backupPath
    }

    if ($PSCmdlet.ShouldProcess($backupPath, 'Create directory')) {
        New-Item -ItemType Directory -Path $backupPath -Force -ErrorAction Stop | Out-Null
        Write-Log -Message "Created backup directory $backupPath"
    }
    return $backupPath
}

# Returns the Calibre Portable directory below the HiDrive sync root and fails when it does not exist.
function Get-CalibreFolderPath {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$HiDriveSyncRoot
    )

    $calibreFolder = Join-Path -Path $HiDriveSyncRoot -ChildPath 'PortableApps\Calibre Portable'
    if (-not (Test-Path -LiteralPath $calibreFolder -PathType Container)) {
        $exception = [System.IO.DirectoryNotFoundException]::new("The path $calibreFolder does not exist or is not accessible.")
        $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -Exception $exception -ErrorId 'CalibreFolderNotFound' -Category ObjectNotFound -TargetObject $calibreFolder))
    }

    Write-Log -Message "Calibre Portable path $calibreFolder was verified."
    return $calibreFolder
}

# Downloads the Calibre Portable installer after removing a stale installer from a previous run.
function Save-CalibreUpdate {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidatePattern('^https://')]
        [string]$SourceUri,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$InstallerPath
    )

    if (-not $PSCmdlet.ShouldProcess($InstallerPath, "Download $SourceUri")) {
        return
    }

    if (Test-Path -LiteralPath $InstallerPath -PathType Leaf) {
        Write-Log -Message "Calibre update file from a previous update found at $InstallerPath. Deleting."
        Remove-Item -LiteralPath $InstallerPath -Force -ErrorAction Stop
    }

    Write-Log -Message "Downloading $SourceUri to $InstallerPath"
    Start-BitsTransfer -Priority Foreground -Source $SourceUri -Destination $InstallerPath -ErrorAction Stop

    if (-not (Test-Path -LiteralPath $InstallerPath -PathType Leaf)) {
        $exception = [System.IO.FileNotFoundException]::new("Download failed: the update file $InstallerPath is missing.")
        $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -Exception $exception -ErrorId 'CalibreInstallerMissing' -Category ObjectNotFound -TargetObject $InstallerPath))
    }
    Write-Log -Message "Calibre update downloaded successfully to $InstallerPath"
}

# Creates a split, maximum-compression 7-Zip backup of the Calibre Portable directory and fails on 7-Zip errors.
function New-CalibreBackup {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path -LiteralPath $_ -PathType Leaf })]
        [string]$SevenZipPath,

        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$BackupPath,

        [Parameter(Mandatory = $true)]
        [ValidatePattern('^\d{4}-\d{2}-\d{2}_\d{2}-\d{2}$')]
        [string]$Timestamp
    )

    $archivePath = Join-Path -Path $BackupPath -ChildPath "CalibrePortableBackup_$Timestamp"
    if (-not $PSCmdlet.ShouldProcess($archivePath, "Create 7-Zip backup of $SourcePath")) {
        return
    }

    Write-Log -Message "Starting 7-Zip backup of $SourcePath to $archivePath"
    # 7-Zip options: a creates an archive, mx9 selects maximum compression, v1g creates 1 GB volumes, and bsp2 sends progress to standard error.
    $backupProcess = Start-Process -FilePath $SevenZipPath -ArgumentList "a -mx9 -bsp2 -v1g `"$archivePath`" `"$SourcePath`"" -Wait -NoNewWindow -PassThru -ErrorAction Stop

    # Exit code 1 only reports warnings, for example files that could not be read.
    if ($backupProcess.ExitCode -ne 0 -and $backupProcess.ExitCode -ne 1) {
        $exception = [System.InvalidOperationException]::new("7-Zip backup failed with exit code $($backupProcess.ExitCode).")
        $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -Exception $exception -ErrorId 'SevenZipBackupFailed' -Category InvalidResult -TargetObject $archivePath))
    }
    Write-Log -Message "7-Zip backup finished with exit code $($backupProcess.ExitCode)."
}

# Stops running Calibre processes, runs the downloaded installer against the Calibre Portable directory, and removes the installer afterwards.
function Install-CalibreUpdate {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$InstallerPath,

        [Parameter(Mandatory = $true)]
        [ValidateScript({ Test-Path -LiteralPath $_ -PathType Container })]
        [string]$CalibreFolder
    )

    if (-not $PSCmdlet.ShouldProcess($CalibreFolder, "Install Calibre update from $InstallerPath")) {
        return
    }

    if (-not (Test-Path -LiteralPath $InstallerPath -PathType Leaf)) {
        $exception = [System.IO.FileNotFoundException]::new("Calibre installer $InstallerPath was not found.")
        $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -Exception $exception -ErrorId 'CalibreInstallerNotFound' -Category ObjectNotFound -TargetObject $InstallerPath))
    }

    $calibreProcesses = Get-Process -Name 'calibre*', 'calibre-parallel*', 'ebook-viewer*', 'ebook-edit*' -ErrorAction SilentlyContinue
    if ($calibreProcesses) {
        Write-Log -Message 'Calibre process is running. Stopping Calibre before update.'
        $calibreProcesses | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
    }
    else {
        Write-Log -Message 'Calibre process is not running. Proceeding with update.'
    }

    Write-Log -Message "Calibre update in $InstallerPath will be applied to $CalibreFolder"
    $installProcess = Start-Process -FilePath $InstallerPath -ArgumentList "`"$CalibreFolder`"" -Wait -PassThru -ErrorAction Stop
    if ($installProcess.ExitCode -ne 0) {
        Write-Log -Message "Calibre installer exited with code $($installProcess.ExitCode)." -Level WARN
    }

    if (Test-Path -LiteralPath $InstallerPath -PathType Leaf) {
        Write-Log -Message "Deleting update file $InstallerPath"
        Remove-Item -LiteralPath $InstallerPath -Force -ErrorAction Stop
    }
}

# Deletes complete backup sets that exceed the retention count and keeps the newest sets.
function Remove-ExpiredBackup {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$BackupPath,

        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 100)]
        [int]$Retention
    )

    Write-Log -Message "Cleanup of old backups in $BackupPath"
    if (-not (Test-Path -LiteralPath $BackupPath -PathType Container)) {
        Write-Log -Message "Backup path $BackupPath does not exist. Nothing to clean up."
        return
    }

    $files = @(Get-ChildItem -LiteralPath $BackupPath -Filter 'CalibrePortableBackup_*.7z.*' -File -ErrorAction Stop)
    if ($files.Count -eq 0) {
        Write-Log -Message "No backup files found in $BackupPath."
        return
    }

    # Group the split volumes by their timestamp and sort the sets from oldest to newest.
    $backupSets = @($files | Group-Object -Property {
            $_.BaseName -replace '^CalibrePortableBackup_(\d{4}-\d{2}-\d{2}_\d{2}-\d{2}).*$', '$1'
        } | Sort-Object -Property {
            try {
                [datetime]::ParseExact($_.Name, 'yyyy-MM-dd_HH-mm', [System.Globalization.CultureInfo]::InvariantCulture)
            }
            catch {
                [datetime]::MinValue
            }
        })

    if ($backupSets.Count -le $Retention) {
        Write-Log -Message "No old backups to delete. Only $($backupSets.Count) backup set(s) found (retention: $Retention)."
        return
    }

    foreach ($backupSet in ($backupSets | Select-Object -First ($backupSets.Count - $Retention))) {
        foreach ($file in $backupSet.Group) {
            if ($PSCmdlet.ShouldProcess($file.FullName, 'Delete expired backup volume')) {
                Write-Log -Message "Deleting $($file.FullName)"
                Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
            }
        }
    }
}

# Runs the complete workflow and restarts HiDrive when a step fails while HiDrive is stopped. The installer is downloaded before HiDrive is stopped to keep the offline time short.
function Invoke-CalibreUpdateBackup {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidatePattern('^https://')]
        [string]$CalibreUpdateSource,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$SevenZipPath,

        [Parameter(Mandatory = $true)]
        [ValidateRange(1, 100)]
        [int]$CalibreBackupRetention
    )

    if (-not (Test-Path -LiteralPath $SevenZipPath -PathType Leaf)) {
        $exception = [System.IO.FileNotFoundException]::new("7-Zip binary not found at $SevenZipPath.")
        $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -Exception $exception -ErrorId 'SevenZipNotFound' -Category ObjectNotFound -TargetObject $SevenZipPath))
    }

    $hiDriveSyncRoot = Get-HiDriveSyncRoot
    if (-not $hiDriveSyncRoot) {
        $exception = [System.InvalidOperationException]::new('Could not determine the HiDrive sync root.')
        $PSCmdlet.ThrowTerminatingError((New-ErrorRecord -Exception $exception -ErrorId 'HiDriveSyncRootNotFound' -Category ObjectNotFound))
    }

    $calibreBackupPath = Initialize-CalibreBackupFolder -HiDriveSyncRoot $hiDriveSyncRoot
    $calibreFolder = Get-CalibreFolderPath -HiDriveSyncRoot $hiDriveSyncRoot
    $calibreInstaller = Join-Path -Path $env:TEMP -ChildPath 'calibre-portable-installer.exe'
    $timestamp = (Get-Date).ToString('yyyy-MM-dd_HH-mm')

    Save-CalibreUpdate -SourceUri $CalibreUpdateSource -InstallerPath $calibreInstaller

    $hiDriveStopped = $false
    try {
        if ($PSCmdlet.ShouldProcess('STRATO HiDrive client', 'Stop')) {
            Stop-HiDrive -Confirm:$false -ErrorAction Stop
            $hiDriveStopped = $true
        }
        New-CalibreBackup -SevenZipPath $SevenZipPath -SourcePath $calibreFolder -BackupPath $calibreBackupPath -Timestamp $timestamp
        Install-CalibreUpdate -InstallerPath $calibreInstaller -CalibreFolder $calibreFolder
        if ($hiDriveStopped) {
            Start-HiDrive -Confirm:$false -ErrorAction Stop
            $hiDriveStopped = $false
        }
    }
    finally {
        if ($hiDriveStopped) {
            Write-Log -Message 'Restarting HiDrive after interruption or failure...'
            try {
                Start-HiDrive -Confirm:$false -ErrorAction Stop
                Write-Log -Message 'HiDrive restarted successfully.'
            }
            catch {
                Write-Log -Message 'Failed to restart HiDrive.' -Level ERROR -ErrorRecord $_
            }
        }
    }

    Remove-ExpiredBackup -BackupPath $calibreBackupPath -Retention $CalibreBackupRetention
}
#EndRegion

#Region: Main script execution
# Stop here when the script is dot-sourced, for example by the Pester tests, so that only the functions are loaded.
if ($MyInvocation.InvocationName -eq '.') {
    return
}

# Required modules with their official installers and update commands. DonGrobione.Logging comes first because an update reloads it and must happen before Start-Log.
$requiredModules = @(
    @{
        ModuleName    = 'DonGrobione.Logging'
        InstallerUri  = 'https://raw.githubusercontent.com/DonGrobione/Logging/main/Install.ps1'
        UpdateCommand = 'Update-DonGrobioneLogging'
    }
    @{
        ModuleName    = 'DonGrobione.StratoHiDriveUtils'
        InstallerUri  = 'https://raw.githubusercontent.com/DonGrobione/StratoHiDriveUtils/main/Install-StratoHiDriveUtils.ps1'
        UpdateCommand = 'Update-HiDriveUtility'
    }
)

# Logging is not available until the modules are ready, so a failure here is only printed to the terminal.
try {
    $moduleStatuses = @(foreach ($requiredModule in $requiredModules) {
            Initialize-RequiredModule @requiredModule
        })
}
catch {
    Write-Error -ErrorRecord $_ -ErrorAction Continue
    exit 1
}

Start-Log -LogDirectory 'Calibre-Update-Backup'
$exitCode = 0
try {
    Write-Log -Message '=============== Starting script ==============='
    foreach ($moduleStatus in $moduleStatuses) {
        if ($moduleStatus.ErrorRecord) {
            Write-Log -Message $moduleStatus.Message -Level WARN -ErrorRecord $moduleStatus.ErrorRecord
        }
        else {
            Write-Log -Message $moduleStatus.Message
        }
    }
    Invoke-CalibreUpdateBackup -CalibreUpdateSource $CalibreUpdateSource -SevenZipPath $SevenZipPath -CalibreBackupRetention $CalibreBackupRetention
    Write-Log -Message '=============== Script completed ==============='
}
catch {
    Write-Log -Message 'Script aborted.' -Level FATAL -ErrorRecord $_
    Write-Error -ErrorRecord $_ -ErrorAction Continue
    $exitCode = 1
}
finally {
    Stop-Log
}
exit $exitCode
#EndRegion
