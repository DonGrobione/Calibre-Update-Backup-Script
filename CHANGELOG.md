# Changelog

All notable changes to this project are documented in this file.
The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and the project uses [Semantic Versioning](https://semver.org/).

## Calibre Update Backup.ps1

### [2.2.0] - 2026-09-26

#### Added
- Script parameters `-CalibreUpdateSource`, `-SevenZipPath`, and `-CalibreBackupRetention` with validation replace the editable configuration variables.
- `-WhatIf` and `-Confirm` support for the whole workflow and for every state-changing function.
- Exit code 0 on success and 1 on failure, and the error is also printed to the terminal.
- The required modules DonGrobione.Logging and DonGrobione.StratoHiDriveUtils are installed for the current user with their official installers when they are missing, and updated with their own update commands when they are installed (`Initialize-RequiredModule`).
- `PSScriptInfo` metadata block for PowerShell Gallery compliance.
- Pester tests in `Tests/Calibre Update Backup.Tests.ps1`.

#### Changed
- The DonGrobione.StratoHiDriveUtils self-update now calls `Update-HiDriveUtility`, the command name used by version 3.x. The previous call to `Update-StratoHiDriveUtils` failed on every run and was only logged as a warning.
- Logging now uses the DonGrobione.Logging module instead of the inline `Write-Log` function. Each run writes its own log file to `<Documents>\Logs\Calibre-Update-Backup`, the five newest log files are kept, and failures are logged with their stack trace. `Calibre-Backup-Update.log` next to the script is no longer written.
- Errors are logged once by the main block instead of once in the failing function and again in the main block.
- 7-Zip is checked before the installer is downloaded, so a missing 7-Zip no longer stops HiDrive first.
- Functions take parameters instead of reading script-scoped variables. `Set-CalibreBackupPath` became `Initialize-CalibreBackupFolder`, `Set-CalibreFolderPath` became `Get-CalibreFolderPath`, `Get-CalibreUpdate` became `Save-CalibreUpdate`, `Remove-ExpiredBackups` became `Remove-ExpiredBackup`, and `Initialize-StratoHiDriveUtils` was replaced by `Initialize-RequiredModule`.
- The workflow moved into `Invoke-CalibreUpdateBackup`, which restarts HiDrive in its `finally` block when a step fails while HiDrive is stopped.

#### Removed
- The email address in the comment-based help.
- The `*.log` entry in `.gitignore`, because the script no longer writes a log file into the repository folder.

### [2.1.2] and earlier

See the Git history.
