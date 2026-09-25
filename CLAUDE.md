# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A single Windows PowerShell 5.1 script, `Calibre Update Backup.ps1`, that backs up and updates a Calibre Portable installation stored inside a STRATO HiDrive sync folder. `Legacy/Calibre Update Backup.bat` is kept for reference only and is not maintained.

This file is the single source of project rules and repository memory. Do not create separate rule or memory files for other AI tools.

## Validation

There is no build or test suite. Before finishing a change, parse the script with Windows PowerShell 5.1 (not pwsh) and run PSScriptAnalyzer if installed. Fix any diagnostics that the change introduced:

```powershell
# Run from PowerShell; the script blocks keep $variables from expanding in the calling shell.
powershell.exe -NoProfile -Command { $errors = $null; [void][System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path '.\Calibre Update Backup.ps1'), [ref]$null, [ref]$errors); $errors }
powershell.exe -NoProfile -Command { Invoke-ScriptAnalyzer -Path '.\Calibre Update Backup.ps1' }
```

Do not run the script itself as a test: it stops the HiDrive client, force-kills running Calibre processes, overwrites the Calibre Portable installation, and deletes old backup sets.

## Architecture

The script has three regions: configuration variables, functions, and a main `try/catch/finally` block that calls the functions in sequence.

- **External dependency:** `Get-HiDriveSyncRoot`, `Stop-HiDrive`, `Start-HiDrive`, and `Update-StratoHiDriveUtils` come from the separate [StratoHiDriveUtils](https://github.com/DonGrobione/StratoHiDriveUtils) module, not from this repo. The script also needs 7-Zip and the STRATO HiDrive client. `Initialize-StratoHiDriveUtils` imports the module and self-updates it. A failed update is non-fatal as long as `Get-HiDriveSyncRoot` is still available.
- **Path resolution:** `Set-CalibreBackupPath` and `Set-CalibreFolderPath` set `$script:`-scoped variables relative to the HiDrive sync root (`Backup\Calibre` and `PortableApps\Calibre Portable`). Later functions read these script-scoped values.
- **Workflow order matters:** the installer is downloaded *before* HiDrive is stopped, so the time HiDrive is offline stays short. The `$hiDriveStopped` flag is set right after `Stop-HiDrive` and cleared after `Start-HiDrive`. The `finally` block uses it to restart HiDrive on failure. Keep this recovery behavior intact when you change the workflow.
- **Backups:** 7-Zip creates 1 GB split volumes named `CalibrePortableBackup_yyyy-MM-dd_HH-mm.7z.NNN`. 7-Zip exit codes 0 and 1 (warnings) count as success. `Remove-ExpiredBackups` groups the volumes by their timestamp and deletes whole sets beyond `$CalibreBackupRetention`. If you change the filename or `$Date` format, you must also change the regex and the `ParseExact` format there.
- **Logging:** `Write-Log` is currently defined inline and appends to `Calibre-Backup-Update.log` next to the script. The logging rules below describe a shared `Write-Log.psm1` that is imported via `$PSScriptRoot`, but that module does not exist in this repo yet.

## Project Rules

### Documentation
- Keep `README.md` current: project structure, the purpose of every user-facing file and directory, prerequisites, configuration, and usage.
- Write code, comments, comment-based help, output, log messages, and documentation in English.
- Do not break a sentence across lines in `README.md`, comment-based help, or regular comments. Start new lines only between complete sentences or list items.
- Give every commit a clear message that says what changed and why.
- After code changes, check that examples, file structure, and guidelines still match the codebase.

### Comment-Based Help And Comments
- Each `.ps1`/`.psm1` file must contain exactly one comment-based help block, as the first content in the file. Only an optional `#requires` statement may come before it.
- Comment-based help blocks inside functions or elsewhere in the file are strictly forbidden.
- All other comments must be regular single-line `#` comments.
- Every function must have a concise regular comment immediately before its definition that describes its responsibility and observable behavior.

### PowerShell Standards (Windows PowerShell 5.1)
- All `.ps1` and `.psm1` files must parse and run in Windows PowerShell 5.1. Do not use PowerShell 6+ syntax, operators, automatic variables, or cmdlets (for example `??`, `?.`, ternaries, or `&&`/`||`).
- Use approved verbs for function names (check with `Get-Verb`) and descriptive PascalCase/camelCase names without abbreviations (except loop counters).
- Keep functions single-responsibility and action-oriented, with early returns to reduce nesting.
- Validate inputs and paths at boundaries. Use `Join-Path` for constructed paths, `Test-Path` with `-PathType` before dependent operations, and `-ErrorAction Stop` for operations handled by `try/catch`.
- Avoid aliases, `Invoke-Expression`, wildcard exports, global state, and unnecessary side effects.
- Enable `Set-StrictMode -Version Latest` only after verifying that the complete script and all imported modules support it in PowerShell 5.1.

### Logging And Error Handling
- `Write-Log.psm1` is the only module allowed to write log files or emit project log entries.
- Modules and library scripts must not contain logging code, call `Write-Log`, or write log files. They return data, status objects, exception records, or error objects to the calling script.
- The orchestrator or standalone entry-point script decides which events are logged and at what level, and calls `Write-Log` for those events.
- Every log file entry must start with the name of the script that emitted it (for example `[MyScript.ps1]`).
- Child scripts must catch their own errors and rethrow them with `throw` so failures reach the orchestrator, which decides the final logging.
- The orchestrator prints errors to the terminal and writes log entries via `Write-Log`. It must not write duplicate entries for errors that were already logged.
- Scripts return exit code 0 on success and 1 on an unhandled failure. Standalone entry points and orchestrators use exit codes; sub-scripts called by an orchestrator rely on the orchestrator's exit handling.
- Log project-internal paths relative to the repository root. Log paths under `$env:TEMP` as absolute paths.
- Create temporary files under `$env:TEMP` and remove them after use.

### Security
- Never trust user input. Apply defense in depth, least privilege, and fail-secure principles.
- Never store secrets in code. Read credentials, tokens, and API keys at runtime from a `.psd1` file under `Secrets/`.
- Do not log secret values.

### Versioning
- Modules (`.psm1` with a `.psd1` manifest): the manifest is the single source of truth for the version. Do not repeat version metadata in comment-based help.
- Standalone scripts (`.ps1` without a manifest): keep a version number in the comment-based help block and update it (along with `Updated` in `.NOTES`) when behavior changes.

### Project-Specific Rules
- Every script must import `Write-Log.psm1` using a path based on `$PSScriptRoot`.
- Protect critical workflows with `try/catch/finally` and keep recovery behavior such as restarting stopped services or clients.
