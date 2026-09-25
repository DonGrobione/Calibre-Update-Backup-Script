# Calibre Update Backup Script

This repository contains a PowerShell workflow that:

1. Imports and checks for updates of the [StratoHiDriveUtils](https://github.com/DonGrobione/StratoHiDriveUtils) module via its built-in self-update function (`Update-StratoHiDriveUtils`).
2. Resolves Calibre and backup paths from the HiDrive sync root (via `Get-HiDriveSyncRoot`).
3. Downloads the latest Calibre Portable installer.
4. Stops STRATO HiDrive to avoid sync/file-lock issues.
5. Creates a split 7z backup of the Calibre Portable folder.
6. Installs the update.
7. Restarts HiDrive.
8. Deletes old backup sets and keeps only the newest configured amount.

## Main Script

- `Calibre Update Backup.ps1`
	Main script for backup, update, HiDrive stop/start, and retention cleanup.

## Requirements

- Windows PowerShell 5.1+
- [StratoHiDriveUtils](https://github.com/DonGrobione/StratoHiDriveUtils) PowerShell module installed in `PSModulePath`
- 7-Zip installed at `C:\Program Files\7-Zip\7z.exe` (default path used by script)
- STRATO HiDrive client, installed and previously synced at least once (so `Get-HiDriveSyncRoot` can resolve the sync root from HiDrive logs)
- Calibre Portable located at `<HiDriveSyncRoot>\PortableApps\Calibre Portable`, backups written to `<HiDriveSyncRoot>\Backup\Calibre`

## Configuration Notes

- The [StratoHiDriveUtils](https://github.com/DonGrobione/StratoHiDriveUtils) module self-updates via `Update-StratoHiDriveUtils` by checking latest releases on GitHub.
- `Set-CalibreBackupPath` and `Set-CalibreFolderPath` derive their paths from `Get-HiDriveSyncRoot`, no more hostname-specific configuration needed.
- Retention count is controlled by `$CalibreBackupRetention`.
- Installer is downloaded to `$env:TEMP\calibre-portable-installer.exe`.

## Repository Structure

```text
.
|-- Legacy/
|   `-- Calibre Update Backup.bat
|-- Calibre Update Backup.ps1
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
- `Calibre Update Backup.ps1`
	Current maintained script.
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