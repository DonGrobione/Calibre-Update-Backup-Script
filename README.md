# Calibre Update Backup Script

This repository contains a PowerShell workflow that:

1. Resolves Calibre and backup paths based on the current hostname.
2. Downloads the latest Calibre Portable installer.
3. Stops STRATO HiDrive to avoid sync/file-lock issues.
4. Creates a split 7z backup of the Calibre Portable folder.
5. Installs the update.
6. Restarts HiDrive.
7. Deletes old backup sets and keeps only the newest configured amount.

## Main Script

- `Calibre Update Backup.ps1`
	Main script for backup, update, HiDrive stop/start, and retention cleanup.

## Requirements

- Windows PowerShell
- 7-Zip installed at `C:\Program Files\7-Zip\7z.exe` (default path used by script)
- Access to the target backup and Calibre Portable directories defined in the script
- STRATO HiDrive client (optional but supported and handled by the script)

## Configuration Notes

- Host-specific paths are configured in:
	- `Set-CalibreBackupPath`
	- `Set-CalibreFolderPath`
- Retention count is controlled by `$CalibreBackupRetention`.
- Installer is downloaded to `$env:TEMP\calibre-portable-installer.exe`.

## Repository Structure

```text
.
|-- Legacy/
|   `-- Calibre Update Backup.bat
|-- Calibre Update Backup.ps1
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
- `.gitignore`
	Git ignore rules for local/runtime artifacts.
- `LICENSE.md`
	MIT License file.
- `README.md`
	Project documentation.

## Legacy Notes

All files under `Legacy/` are retained for reference and migration history. Use `Calibre Update Backup.ps1` for current operations.

## License

This project is licensed under the [MIT License](LICENSE.md). You are free to use and modify the code as long as you include the original copyright notice and credit the author.