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
|-- .copilot/
|   |-- project-rules.md
|   `-- repo-memory.md
|-- Legacy/
|   `-- Calibre Update Backup.bat
|-- Calibre Update Backup.ps1
|-- .gitignore
|-- LICENSE.md
`-- README.md
```

## File and Directory Purpose

- `.copilot/`
	Repository-scoped Copilot rules and concise repository memory.
- `.copilot/project-rules.md`
	Canonical project rules for documentation, comments, PowerShell 5.1 compatibility, and validation.
- `.copilot/repo-memory.md`
	Persistent, repository-specific facts used by Copilot.
- `Legacy/`
	Historical scripts kept for reference only; not actively maintained.
- `Legacy/Calibre Update Backup.bat`
	Older batch-file implementation.
- `Calibre Update Backup.ps1`
	Current maintained script.
- `.gitignore`
	Git ignore rules for local/runtime artifacts.
- `LICENSE.md`
	CC BY-NC-SA 4.0 license file.
- `README.md`
	Project documentation.

## Legacy Notes

All files under `Legacy/` are retained for reference and migration history. Use `Calibre Update Backup.ps1` for current operations.

## License

This project is licensed under the [Creative Commons Attribution-NonCommercial-ShareAlike 4.0 International License](LICENSE.md). You may use, share, and adapt the material for non-commercial purposes, provided you give appropriate credit and distribute any changes under the same license.

## AI Usage

This project was created with support from AI/KI tooling.