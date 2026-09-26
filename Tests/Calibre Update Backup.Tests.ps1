<#
.SYNOPSIS
    Pester tests for Calibre Update Backup.ps1.

.DESCRIPTION
    Dot-sources the script, which loads only its functions, and tests each function in isolation.
    Every destructive or external call (HiDrive, 7-Zip, downloads, installers, module installation and updates, processes, logging) is mocked, and all files are created in TestDrive.
    Uses Pester v4-compatible syntax so the tests run in Windows PowerShell 5.1 with Pester 3.4 or 4.x.
#>

# Stub for the DonGrobione.StratoHiDriveUtils command so it can be mocked without the module installed.
function Get-HiDriveSyncRoot {
    [CmdletBinding()]
    param()
}

# Stub for the DonGrobione.StratoHiDriveUtils command so it can be mocked without the module installed.
function Stop-HiDrive {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    if ($PSCmdlet.ShouldProcess('HiDrive stub', 'Stop')) {
        throw 'The Stop-HiDrive stub must be mocked.'
    }
}

# Stub for the DonGrobione.StratoHiDriveUtils command so it can be mocked without the module installed.
function Start-HiDrive {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    if ($PSCmdlet.ShouldProcess('HiDrive stub', 'Start')) {
        throw 'The Start-HiDrive stub must be mocked.'
    }
}

# Stub for the update command of a fictitious required module used by the Initialize-RequiredModule tests.
function Update-TestModule {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param()
    if ($PSCmdlet.ShouldProcess('Test.Module', 'Update')) {
        throw 'The Update-TestModule stub must be mocked.'
    }
}

# Stub for the DonGrobione.Logging command so it can be mocked without the module installed.
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter()]
        [ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR', 'FATAL')]
        [string]$Level = 'INFO',

        [Parameter()]
        [System.Management.Automation.ErrorRecord]$ErrorRecord
    )
    throw "The Write-Log stub must be mocked. Level: $Level, message: $Message, error: $ErrorRecord"
}

# Stub that shadows the BitsTransfer cmdlet, which Pester 3.4 cannot mock in Windows PowerShell 5.1.
function Start-BitsTransfer {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'Test-only stub that deliberately shadows the cmdlet so it can be mocked.')]
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter()]
        [string]$Source,

        [Parameter()]
        [string]$Destination,

        [Parameter()]
        [string]$Priority
    )
    if ($PSCmdlet.ShouldProcess($Destination, "Download $Source with priority $Priority")) {
        throw 'The Start-BitsTransfer stub must be mocked.'
    }
}

. (Join-Path -Path $PSScriptRoot -ChildPath '..\Calibre Update Backup.ps1')

Describe 'New-ErrorRecord' {
    It 'builds an ErrorRecord with the given id, category, and target' {
        $record = New-ErrorRecord -Exception ([System.IO.FileNotFoundException]::new('missing')) -ErrorId 'TestId' -Category ObjectNotFound -TargetObject 'target'

        ($record -is [System.Management.Automation.ErrorRecord]) | Should Be $true
        $record.FullyQualifiedErrorId | Should Be 'TestId'
        $record.CategoryInfo.Category | Should Be 'ObjectNotFound'
        $record.TargetObject | Should Be 'target'
        $record.Exception.Message | Should Be 'missing'
    }
}

Describe 'Install-RequiredModule' {
    # Pester 3.x keeps mocks defined in an It block until the end of the enclosing Context, so each case gets its own Context.
    $markerPath = Join-Path -Path $TestDrive -ChildPath 'installer-ran.txt'
    $installerParameters = @{
        ModuleName   = 'Test.Module'
        InstallerUri = 'https://example.invalid/Install.ps1'
    }

    Context 'Installer succeeds' {
        It 'downloads the installer to TEMP, runs it, and deletes it' {
            $installerScript = "Set-Content -LiteralPath '$markerPath' -Value 'ran'"
            Mock Start-BitsTransfer { Set-Content -LiteralPath $Destination -Value $installerScript }

            Install-RequiredModule @installerParameters
            Assert-MockCalled Start-BitsTransfer -Times 1 -Exactly -Scope It -ParameterFilter {
                $Source -eq 'https://example.invalid/Install.ps1' -and $Destination.StartsWith($env:TEMP)
            }
            Get-Content -LiteralPath $markerPath | Should Be 'ran'
            @(Get-ChildItem -Path $env:TEMP -Filter 'Test.Module-Install-*.ps1').Count | Should Be 0
        }
    }

    Context 'Installer fails' {
        It 'fails and deletes the installer' {
            Mock Start-BitsTransfer { Set-Content -LiteralPath $Destination -Value "throw 'installer broken'" }

            { Install-RequiredModule @installerParameters } | Should Throw 'could not be installed'
            @(Get-ChildItem -Path $env:TEMP -Filter 'Test.Module-Install-*.ps1').Count | Should Be 0
        }
    }

    Context 'WhatIf' {
        It 'fails without downloading the installer' {
            Mock Start-BitsTransfer {}

            { Install-RequiredModule @installerParameters -WhatIf } | Should Throw 'installation was skipped'
            Assert-MockCalled Start-BitsTransfer -Times 0 -Exactly -Scope It
        }
    }
}

Describe 'Initialize-RequiredModule' {
    # Pester 3.x keeps mocks defined in an It block until the end of the enclosing Context, so each case gets its own Context.
    $moduleParameters = @{
        ModuleName    = 'Test.Module'
        InstallerUri  = 'https://example.invalid/Install.ps1'
        UpdateCommand = 'Update-TestModule'
    }
    Mock Install-RequiredModule {}
    Mock Update-TestModule {}

    Context 'Installed and up to date' {
        It 'runs the update command and reports UpToDate' {
            Mock Import-Module { [pscustomobject]@{ Version = [version]'1.0.0' } } -ParameterFilter { $Name -eq 'Test.Module' }

            $status = Initialize-RequiredModule @moduleParameters
            $status.Status | Should Be 'UpToDate'
            $status.Version | Should Be ([version]'1.0.0')
            $status.ErrorRecord | Should BeNullOrEmpty
            Assert-MockCalled Update-TestModule -Times 1 -Exactly -Scope It
            Assert-MockCalled Install-RequiredModule -Times 0 -Exactly -Scope It
        }
    }

    Context 'Installed and updated' {
        It 'reloads the module and reports the new version' {
            $versions = @{ Queue = [System.Collections.Queue]::new(@([version]'1.0.0', [version]'1.1.0')) }
            Mock Import-Module { [pscustomobject]@{ Version = $versions.Queue.Dequeue() } } -ParameterFilter { $Name -eq 'Test.Module' }

            $status = Initialize-RequiredModule @moduleParameters
            $status.Status | Should Be 'Updated'
            $status.Version | Should Be ([version]'1.1.0')
            $status.Message | Should Match '1\.0\.0 to 1\.1\.0'
            Assert-MockCalled Import-Module -Times 1 -Exactly -Scope It -ParameterFilter { $Name -eq 'Test.Module' -and $Force }
        }
    }

    Context 'Installed and the update fails' {
        It 'keeps the installed version and returns the error for logging' {
            Mock Import-Module { [pscustomobject]@{ Version = [version]'1.0.0' } } -ParameterFilter { $Name -eq 'Test.Module' }
            Mock Update-TestModule { throw 'GitHub unreachable' }

            $status = Initialize-RequiredModule @moduleParameters
            $status.Status | Should Be 'UpdateFailed'
            $status.Version | Should Be ([version]'1.0.0')
            $status.ErrorRecord.Exception.Message | Should Be 'GitHub unreachable'
        }
    }

    Context 'Not installed' {
        It 'installs the module with its installer, imports it, and skips the update' {
            $attempts = @{ Count = 0 }
            Mock Import-Module {
                $attempts.Count++
                if ($attempts.Count -eq 1) {
                    throw 'module not found'
                }
                [pscustomobject]@{ Version = [version]'2.0.0' }
            } -ParameterFilter { $Name -eq 'Test.Module' }

            $status = Initialize-RequiredModule @moduleParameters
            $status.Status | Should Be 'Installed'
            $status.Version | Should Be ([version]'2.0.0')
            Assert-MockCalled Install-RequiredModule -Times 1 -Exactly -Scope It -ParameterFilter {
                $ModuleName -eq 'Test.Module' -and $InstallerUri -eq 'https://example.invalid/Install.ps1'
            }
            Assert-MockCalled Update-TestModule -Times 0 -Exactly -Scope It
        }
    }

    Context 'Not installed and the installation fails' {
        It 'fails' {
            Mock Import-Module { throw 'module not found' } -ParameterFilter { $Name -eq 'Test.Module' }
            Mock Install-RequiredModule { throw 'installation failed' }

            { Initialize-RequiredModule @moduleParameters } | Should Throw 'installation failed'
        }
    }

    Context 'WhatIf' {
        It 'imports the installed module but skips the update' {
            Mock Import-Module { [pscustomobject]@{ Version = [version]'1.0.0' } } -ParameterFilter { $Name -eq 'Test.Module' }

            $status = Initialize-RequiredModule @moduleParameters -WhatIf
            $status.Status | Should Be 'UpdateSkipped'
            Assert-MockCalled Update-TestModule -Times 0 -Exactly -Scope It
        }
    }
}

Describe 'Initialize-CalibreBackupFolder' {
    Mock Write-Log {}

    It 'returns the existing backup folder' {
        $syncRoot = Join-Path -Path $TestDrive -ChildPath 'existing'
        New-Item -ItemType Directory -Path (Join-Path -Path $syncRoot -ChildPath 'Backup\Calibre') -Force | Out-Null

        Initialize-CalibreBackupFolder -HiDriveSyncRoot $syncRoot | Should Be (Join-Path -Path $syncRoot -ChildPath 'Backup\Calibre')
    }

    It 'creates a missing backup folder' {
        $syncRoot = Join-Path -Path $TestDrive -ChildPath 'create'
        New-Item -ItemType Directory -Path $syncRoot -Force | Out-Null

        $backupPath = Initialize-CalibreBackupFolder -HiDriveSyncRoot $syncRoot
        $backupPath | Should Exist
    }

    It 'does not create the folder with -WhatIf' {
        $syncRoot = Join-Path -Path $TestDrive -ChildPath 'whatif'
        New-Item -ItemType Directory -Path $syncRoot -Force | Out-Null

        $backupPath = Initialize-CalibreBackupFolder -HiDriveSyncRoot $syncRoot -WhatIf
        $backupPath | Should Be (Join-Path -Path $syncRoot -ChildPath 'Backup\Calibre')
        $backupPath | Should Not Exist
    }

    It 'rejects a sync root that does not exist' {
        { Initialize-CalibreBackupFolder -HiDriveSyncRoot (Join-Path -Path $TestDrive -ChildPath 'missing') } | Should Throw
    }
}

Describe 'Get-CalibreFolderPath' {
    Mock Write-Log {}

    It 'returns the Calibre Portable folder when it exists' {
        $syncRoot = Join-Path -Path $TestDrive -ChildPath 'present'
        $calibreFolder = Join-Path -Path $syncRoot -ChildPath 'PortableApps\Calibre Portable'
        New-Item -ItemType Directory -Path $calibreFolder -Force | Out-Null

        Get-CalibreFolderPath -HiDriveSyncRoot $syncRoot | Should Be $calibreFolder
    }

    It 'fails when the Calibre Portable folder is missing' {
        { Get-CalibreFolderPath -HiDriveSyncRoot (Join-Path -Path $TestDrive -ChildPath 'absent') } | Should Throw 'does not exist'
    }
}

Describe 'Save-CalibreUpdate' {
    Mock Write-Log {}

    It 'replaces a stale installer with the downloaded file' {
        $installerPath = Join-Path -Path $TestDrive -ChildPath 'stale-installer.exe'
        Set-Content -LiteralPath $installerPath -Value 'old'
        Mock Start-BitsTransfer { Add-Content -LiteralPath $Destination -Value 'new' }

        Save-CalibreUpdate -SourceUri 'https://example.invalid/portable' -InstallerPath $installerPath
        Get-Content -LiteralPath $installerPath | Should Be 'new'
        Assert-MockCalled Start-BitsTransfer -Times 1 -Exactly -Scope It -ParameterFilter { $Source -eq 'https://example.invalid/portable' }
    }

    It 'fails when the installer is missing after the download' {
        $installerPath = Join-Path -Path $TestDrive -ChildPath 'missing-installer.exe'
        Mock Start-BitsTransfer {}

        { Save-CalibreUpdate -SourceUri 'https://example.invalid/portable' -InstallerPath $installerPath } | Should Throw 'Download failed'
    }

    It 'does not download with -WhatIf' {
        Mock Start-BitsTransfer {}

        Save-CalibreUpdate -SourceUri 'https://example.invalid/portable' -InstallerPath (Join-Path -Path $TestDrive -ChildPath 'whatif.exe') -WhatIf
        Assert-MockCalled Start-BitsTransfer -Times 0 -Exactly -Scope It
    }

    It 'rejects a non-HTTPS source' {
        { Save-CalibreUpdate -SourceUri 'http://example.invalid/portable' -InstallerPath (Join-Path -Path $TestDrive -ChildPath 'http.exe') } | Should Throw
    }
}

Describe 'New-CalibreBackup' {
    Mock Write-Log {}
    $sevenZipPath = Join-Path -Path $TestDrive -ChildPath '7z.exe'
    Set-Content -LiteralPath $sevenZipPath -Value 'stub'
    $sourcePath = Join-Path -Path $TestDrive -ChildPath 'Calibre Portable'
    New-Item -ItemType Directory -Path $sourcePath -Force | Out-Null
    $backupPath = Join-Path -Path $TestDrive -ChildPath 'Backup'

    It 'runs 7-Zip with split volumes and the timestamped archive name' {
        Mock Start-Process { [pscustomobject]@{ ExitCode = 0 } }

        New-CalibreBackup -SevenZipPath $sevenZipPath -SourcePath $sourcePath -BackupPath $backupPath -Timestamp '2026-09-26_10-00'
        Assert-MockCalled Start-Process -Times 1 -Exactly -Scope It -ParameterFilter {
            $FilePath -eq $sevenZipPath -and
            ($ArgumentList -join ' ') -match '-v1g' -and
            ($ArgumentList -join ' ') -match [regex]::Escape('CalibrePortableBackup_2026-09-26_10-00')
        }
    }

    It 'accepts 7-Zip exit code 1 (warnings)' {
        Mock Start-Process { [pscustomobject]@{ ExitCode = 1 } }

        { New-CalibreBackup -SevenZipPath $sevenZipPath -SourcePath $sourcePath -BackupPath $backupPath -Timestamp '2026-09-26_10-00' } | Should Not Throw
    }

    It 'fails on 7-Zip exit code 2' {
        Mock Start-Process { [pscustomobject]@{ ExitCode = 2 } }

        { New-CalibreBackup -SevenZipPath $sevenZipPath -SourcePath $sourcePath -BackupPath $backupPath -Timestamp '2026-09-26_10-00' } | Should Throw 'exit code 2'
    }

    It 'does not run 7-Zip with -WhatIf' {
        Mock Start-Process {}

        New-CalibreBackup -SevenZipPath $sevenZipPath -SourcePath $sourcePath -BackupPath $backupPath -Timestamp '2026-09-26_10-00' -WhatIf
        Assert-MockCalled Start-Process -Times 0 -Exactly -Scope It
    }

    It 'rejects a timestamp in the wrong format' {
        Mock Start-Process {}

        { New-CalibreBackup -SevenZipPath $sevenZipPath -SourcePath $sourcePath -BackupPath $backupPath -Timestamp '26.09.2026' } | Should Throw
    }

    It 'rejects a missing 7-Zip binary' {
        Mock Start-Process {}

        { New-CalibreBackup -SevenZipPath (Join-Path -Path $TestDrive -ChildPath 'none.exe') -SourcePath $sourcePath -BackupPath $backupPath -Timestamp '2026-09-26_10-00' } | Should Throw
    }
}

Describe 'Install-CalibreUpdate' {
    Mock Write-Log {}
    Mock Start-Sleep {}
    Mock Stop-Process {}
    $calibreFolder = Join-Path -Path $TestDrive -ChildPath 'Calibre Portable'
    New-Item -ItemType Directory -Path $calibreFolder -Force | Out-Null
    $installerPath = Join-Path -Path $TestDrive -ChildPath 'installer.exe'

    It 'runs the installer against the Calibre folder and deletes the installer' {
        Set-Content -LiteralPath $installerPath -Value 'stub'
        Mock Get-Process {}
        Mock Start-Process { [pscustomobject]@{ ExitCode = 0 } }

        Install-CalibreUpdate -InstallerPath $installerPath -CalibreFolder $calibreFolder
        Assert-MockCalled Start-Process -Times 1 -Exactly -Scope It -ParameterFilter { $FilePath -eq $installerPath -and ($ArgumentList -join ' ') -match [regex]::Escape($calibreFolder) }
        Assert-MockCalled Stop-Process -Times 0 -Exactly -Scope It
        $installerPath | Should Not Exist
    }

    It 'stops running Calibre processes before installing' {
        Set-Content -LiteralPath $installerPath -Value 'stub'
        Mock Get-Process { [System.Diagnostics.Process]::new() }
        Mock Start-Process { [pscustomobject]@{ ExitCode = 0 } }

        Install-CalibreUpdate -InstallerPath $installerPath -CalibreFolder $calibreFolder
        Assert-MockCalled Stop-Process -Times 1 -Exactly -Scope It
    }

    It 'logs a warning when the installer returns a non-zero exit code' {
        Set-Content -LiteralPath $installerPath -Value 'stub'
        Mock Get-Process {}
        Mock Start-Process { [pscustomobject]@{ ExitCode = 5 } }

        Install-CalibreUpdate -InstallerPath $installerPath -CalibreFolder $calibreFolder
        Assert-MockCalled Write-Log -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'WARN' }
    }

    It 'fails when the installer is missing' {
        Mock Get-Process {}
        Mock Start-Process {}

        { Install-CalibreUpdate -InstallerPath (Join-Path -Path $TestDrive -ChildPath 'absent.exe') -CalibreFolder $calibreFolder } | Should Throw 'was not found'
        Assert-MockCalled Start-Process -Times 0 -Exactly -Scope It
    }

    It 'does not stop processes or install with -WhatIf' {
        Set-Content -LiteralPath $installerPath -Value 'stub'
        Mock Get-Process { [System.Diagnostics.Process]::new() }
        Mock Start-Process {}

        Install-CalibreUpdate -InstallerPath $installerPath -CalibreFolder $calibreFolder -WhatIf
        Assert-MockCalled Stop-Process -Times 0 -Exactly -Scope It
        Assert-MockCalled Start-Process -Times 0 -Exactly -Scope It
        $installerPath | Should Exist
    }
}

Describe 'Remove-ExpiredBackup' {
    Mock Write-Log {}

    # Creates two split volumes for each timestamp and returns the directory that contains them.
    function New-TestBackupSet {
        [CmdletBinding(SupportsShouldProcess = $true)]
        param(
            [Parameter(Mandatory = $true)]
            [string]$Name,

            [Parameter(Mandatory = $true)]
            [string[]]$Timestamp
        )

        $directory = Join-Path -Path $TestDrive -ChildPath $Name
        if ($PSCmdlet.ShouldProcess($directory, 'Create test backup sets')) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
            foreach ($stamp in $Timestamp) {
                foreach ($volume in '001', '002') {
                    Set-Content -LiteralPath (Join-Path -Path $directory -ChildPath "CalibrePortableBackup_$stamp.7z.$volume") -Value 'x'
                }
            }
        }
        return $directory
    }

    $timestamps = '2026-09-01_10-00', '2026-09-02_10-00', '2026-09-03_10-00', '2026-09-04_10-00', '2026-09-05_10-00'

    It 'deletes the oldest complete sets beyond the retention count' {
        $directory = New-TestBackupSet -Name 'expire' -Timestamp $timestamps

        Remove-ExpiredBackup -BackupPath $directory -Retention 3
        $remaining = @(Get-ChildItem -LiteralPath $directory -File | Select-Object -ExpandProperty Name)
        $remaining.Count | Should Be 6
        ($remaining -match '2026-09-0[12]_').Count | Should Be 0
    }

    It 'keeps everything when the set count does not exceed the retention' {
        $directory = New-TestBackupSet -Name 'keep' -Timestamp ($timestamps | Select-Object -First 3)

        Remove-ExpiredBackup -BackupPath $directory -Retention 3
        @(Get-ChildItem -LiteralPath $directory -File).Count | Should Be 6
    }

    It 'keeps unrelated files' {
        $directory = New-TestBackupSet -Name 'unrelated' -Timestamp $timestamps
        Set-Content -LiteralPath (Join-Path -Path $directory -ChildPath 'notes.txt') -Value 'x'

        Remove-ExpiredBackup -BackupPath $directory -Retention 1
        (Join-Path -Path $directory -ChildPath 'notes.txt') | Should Exist
        @(Get-ChildItem -LiteralPath $directory -Filter '*.7z.*' -File).Count | Should Be 2
    }

    It 'does nothing when the backup path does not exist' {
        { Remove-ExpiredBackup -BackupPath (Join-Path -Path $TestDrive -ChildPath 'nothing') -Retention 3 } | Should Not Throw
    }

    It 'does not delete anything with -WhatIf' {
        $directory = New-TestBackupSet -Name 'whatif' -Timestamp $timestamps

        Remove-ExpiredBackup -BackupPath $directory -Retention 1 -WhatIf
        @(Get-ChildItem -LiteralPath $directory -File).Count | Should Be 10
    }
}

Describe 'Invoke-CalibreUpdateBackup' {
    # Pester 3.x keeps mocks defined in an It block until the end of the enclosing Context, so cases that override a mock get their own Context.
    Mock Write-Log {}
    Mock Get-HiDriveSyncRoot { 'X:\HiDrive' }
    Mock Initialize-CalibreBackupFolder { 'X:\HiDrive\Backup\Calibre' }
    Mock Get-CalibreFolderPath { 'X:\HiDrive\PortableApps\Calibre Portable' }
    Mock Save-CalibreUpdate {}
    Mock Stop-HiDrive {}
    Mock Start-HiDrive {}
    Mock New-CalibreBackup {}
    Mock Install-CalibreUpdate {}
    Mock Remove-ExpiredBackup {}

    $sevenZipPath = Join-Path -Path $TestDrive -ChildPath '7z.exe'
    Set-Content -LiteralPath $sevenZipPath -Value 'stub'
    $workflowParameters = @{
        CalibreUpdateSource    = 'https://example.invalid/portable'
        SevenZipPath           = $sevenZipPath
        CalibreBackupRetention = 3
    }

    It 'runs every step and restarts HiDrive once' {
        Invoke-CalibreUpdateBackup @workflowParameters

        Assert-MockCalled Save-CalibreUpdate -Times 1 -Exactly -Scope It
        Assert-MockCalled Stop-HiDrive -Times 1 -Exactly -Scope It
        Assert-MockCalled New-CalibreBackup -Times 1 -Exactly -Scope It
        Assert-MockCalled Install-CalibreUpdate -Times 1 -Exactly -Scope It
        Assert-MockCalled Start-HiDrive -Times 1 -Exactly -Scope It
        Assert-MockCalled Remove-ExpiredBackup -Times 1 -Exactly -Scope It -ParameterFilter { $Retention -eq 3 }
    }

    Context 'Backup fails' {
        It 'restarts HiDrive and skips cleanup when the backup fails' {
            Mock New-CalibreBackup { throw 'backup failed' }

            { Invoke-CalibreUpdateBackup @workflowParameters } | Should Throw 'backup failed'
            Assert-MockCalled Start-HiDrive -Times 1 -Exactly -Scope It
            Assert-MockCalled Install-CalibreUpdate -Times 0 -Exactly -Scope It
            Assert-MockCalled Remove-ExpiredBackup -Times 0 -Exactly -Scope It
        }
    }

    Context 'Installation fails' {
        It 'restarts HiDrive when the installation fails' {
            Mock Install-CalibreUpdate { throw 'install failed' }

            { Invoke-CalibreUpdateBackup @workflowParameters } | Should Throw 'install failed'
            Assert-MockCalled Start-HiDrive -Times 1 -Exactly -Scope It
        }
    }

    Context 'Backup and HiDrive restart fail' {
        It 'logs an error but keeps the original failure when the HiDrive restart also fails' {
            Mock New-CalibreBackup { throw 'backup failed' }
            Mock Start-HiDrive { throw 'restart failed' }

            { Invoke-CalibreUpdateBackup @workflowParameters } | Should Throw 'backup failed'
            Assert-MockCalled Write-Log -Times 1 -Exactly -Scope It -ParameterFilter { $Level -eq 'ERROR' -and $ErrorRecord.Exception.Message -eq 'restart failed' }
        }
    }

    Context 'Download fails' {
        It 'does not stop HiDrive when the download fails' {
            Mock Save-CalibreUpdate { throw 'download failed' }

            { Invoke-CalibreUpdateBackup @workflowParameters } | Should Throw 'download failed'
            Assert-MockCalled Stop-HiDrive -Times 0 -Exactly -Scope It
            Assert-MockCalled Start-HiDrive -Times 0 -Exactly -Scope It
        }
    }

    It 'fails before any other step when 7-Zip is missing' {
        $missingSevenZip = $workflowParameters.Clone()
        $missingSevenZip.SevenZipPath = Join-Path -Path $TestDrive -ChildPath 'missing-7z.exe'

        { Invoke-CalibreUpdateBackup @missingSevenZip } | Should Throw '7-Zip binary not found'
        Assert-MockCalled Get-HiDriveSyncRoot -Times 0 -Exactly -Scope It
        Assert-MockCalled Save-CalibreUpdate -Times 0 -Exactly -Scope It
    }

    Context 'Sync root missing' {
        It 'fails when the HiDrive sync root cannot be determined' {
            Mock Get-HiDriveSyncRoot {}

            { Invoke-CalibreUpdateBackup @workflowParameters } | Should Throw 'sync root'
            Assert-MockCalled Save-CalibreUpdate -Times 0 -Exactly -Scope It
        }
    }

    It 'neither stops nor starts HiDrive with -WhatIf' {
        Invoke-CalibreUpdateBackup @workflowParameters -WhatIf

        Assert-MockCalled Stop-HiDrive -Times 0 -Exactly -Scope It
        Assert-MockCalled Start-HiDrive -Times 0 -Exactly -Scope It
    }
}
