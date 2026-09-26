# Calibre Update Backup Script

This repository contains a PowerShell workflow that:

1. Makes the required modules available: a missing module is installed for the current user with its official installer, and an installed module is updated with its own update command (`Update-DonGrobioneLogging`, `Update-HiDriveUtility`).
2. Resolves Calibre and backup paths from the HiDrive sync root (via `Get-HiDriveSyncRoot`).
3. Downloads the latest Calibre Portable installer.
4. Stops STRATO HiDrive to avoid sync/file-lock issues.
5. Creates a split 7z backup of the Calibre Portable folder.
6. Installs the update.
7. Restarts HiDrive.
8. Deletes old backup sets and keeps only the newest configured amount.

If a step fails while HiDrive is stopped, the script restarts HiDrive before it exits.

## Main Script

- `Calibre Update Backup.ps1`
	Main script for backup, update, HiDrive stop/start, and retention cleanup.

## Requirements

- Windows PowerShell 5.1
- Internet access to GitHub, used to install or update the required PowerShell modules on every run:
  - [DonGrobione.Logging](https://github.com/DonGrobione/Logging)
  - [DonGrobione.StratoHiDriveUtils](https://github.com/DonGrobione/StratoHiDriveUtils) 3.x

  Both are installed automatically for the current user when they are missing. A failed update is only logged as a warning, and the script continues with the installed version.
- 7-Zip, by default at `%ProgramFiles%\7-Zip\7z.exe`
- STRATO HiDrive client, installed and previously synced at least once (so `Get-HiDriveSyncRoot` can resolve the sync root from HiDrive logs)
- Calibre Portable located at `<HiDriveSyncRoot>\PortableApps\Calibre Portable`, backups written to `<HiDriveSyncRoot>\Backup\Calibre`
- For running the tests: Pester 3.4 or 4.x (Windows PowerShell 5.1 ships with Pester 3.4.0)

## Usage

```powershell
# Run the full workflow with the default settings.
& '.\Calibre Update Backup.ps1'

# Show what would happen without downloading, stopping HiDrive, backing up, installing, or deleting anything.
& '.\Calibre Update Backup.ps1' -WhatIf

# Keep the five newest backup sets and use a different 7-Zip location.
& '.\Calibre Update Backup.ps1' -CalibreBackupRetention 5 -SevenZipPath 'D:\Tools\7-Zip\7z.exe'
```

The script returns exit code 0 on success and 1 on failure, so it can be used from Task Scheduler.

## Configuration

All settings are script parameters with defaults:

| Parameter | Default | Purpose |
| --- | --- | --- |
| `-CalibreUpdateSource` | `https://calibre-ebook.com/dist/portable` | HTTPS URL of the Calibre Portable installer. |
| `-SevenZipPath` | `%ProgramFiles%\7-Zip\7z.exe` | Full path to `7z.exe`. |
| `-CalibreBackupRetention` | `3` | Number of backup sets to keep (1 to 100). |

- The Calibre and backup paths are derived from `Get-HiDriveSyncRoot`, so no hostname-specific configuration is needed.
- The installer is downloaded to `$env:TEMP\calibre-portable-installer.exe` and deleted after the installation.
- Each run writes its own log file to `<Documents>\Logs\Calibre-Update-Backup\<HOSTNAME>_<timestamp>.log` via DonGrobione.Logging, and the five newest log files are kept.

## Testing

```powershell
Import-Module Pester -RequiredVersion 3.4.0
Invoke-Pester -Path '.\Tests'
```

The tests mock every destructive or external call, including logging and the module installation, and write only to Pester's `TestDrive`, so they never touch HiDrive, Calibre, existing backups, or log files. They do not need DonGrobione.Logging or DonGrobione.StratoHiDriveUtils to be installed.

## Repository Structure

```text
.
|-- Legacy/
|   `-- Calibre Update Backup.bat
|-- Tests/
|   `-- Calibre Update Backup.Tests.ps1
|-- Calibre Update Backup.ps1
|-- CHANGELOG.md
|-- CLAUDE.md
|-- .gitignore
|-- LICENSE.md
`-- README.md
```

## File and Directory Purpose

- `Legacy/`
	Historical scripts kept for reference only; not actively maintained.
- `Legacy/Calibre Update Backup.bat`
	Older batch-file implementation.
- `Tests/`
	Pester tests, one file per script.
- `Calibre Update Backup.ps1`
	Current maintained script.
- `CHANGELOG.md`
	Version history of the script.
- `CLAUDE.md`
	Project rules, repository memory, and architecture notes for Claude Code.
- `.gitignore`
	Git ignore rules for local/runtime artifacts.
- `LICENSE.md`
	GNU AGPL v3 license file.
- `README.md`
	Project documentation.

## Legacy Notes

All files under `Legacy/` are retained for reference and migration history. Use `Calibre Update Backup.ps1` for current operations.

## License

This project is licensed under the [GNU Affero General Public License v3.0](LICENSE.md). You may use, modify, and distribute it, provided that you make the source code of modified versions available under the same license, including when you offer the software to users over a network.

## AI Usage

This project was created with support from AI/KI tooling.