# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A single Windows PowerShell 5.1 script, `Calibre Update Backup.ps1`, that backs up and updates a Calibre Portable installation stored inside a STRATO HiDrive sync folder. `Legacy/Calibre Update Backup.bat` is kept for reference only and is not maintained.

This file is the single source of project rules and repository memory. Do not create separate rule or memory files for other AI tools.

## Validation

There is no build step. Before finishing a change, parse every changed `.ps1`/`.psm1` file with Windows PowerShell 5.1 (not pwsh), run PSScriptAnalyzer, and run the affected Pester tests (see "Testing (Pester)" below). The code must produce zero PSScriptAnalyzer warnings with the standard rule set:

```powershell
# Run from PowerShell; the script blocks keep $variables from expanding in the calling shell.
powershell.exe -NoProfile -Command { $errors = $null; [void][System.Management.Automation.Language.Parser]::ParseFile((Resolve-Path '.\Calibre Update Backup.ps1'), [ref]$null, [ref]$errors); $errors }
powershell.exe -NoProfile -Command { Invoke-ScriptAnalyzer -Path . -Recurse }
powershell.exe -NoProfile -Command { Import-Module Pester -RequiredVersion 3.4.0; Invoke-Pester -Path '.\Tests' }
powershell.exe -NoProfile -Command { Test-ScriptFileInfo -Path '.\Calibre Update Backup.ps1' }
```

Do not run the script itself as a test: it stops the HiDrive client, force-kills running Calibre processes, overwrites the Calibre Portable installation, and deletes old backup sets.

- Pester 5.x is also installed on this machine but does not support the `Should Be` syntax, so always import Pester 3.4.0 (or 4.x) explicitly.
- Pester 3.x keeps a mock defined inside an `It` block until the end of the enclosing `Context` or `Describe`, so put tests that override a shared mock in their own `Context`.
- Pester 3.4 in Windows PowerShell 5.1 cannot mock `Get-Module` (its `-PSEdition` parameter collides with a read-only automatic variable) or `Start-BitsTransfer`. The script therefore does not call `Get-Module`, and the tests shadow `Start-BitsTransfer` with a stub before mocking it.
- The tests define stubs for the external module commands (`Write-Log`, `Get-HiDriveSyncRoot`, `Stop-HiDrive`, `Start-HiDrive`) and a fictitious `Update-TestModule` update command, so they run without DonGrobione.Logging or DonGrobione.StratoHiDriveUtils installed.

## Architecture

The script has a `PSScriptInfo` block, comment-based help, a validated `param()` block, a functions region, and a main region.
The main region returns early when the script is dot-sourced (`$MyInvocation.InvocationName -eq '.'`), so the tests can load the functions without running the workflow.
Otherwise it runs `Initialize-RequiredModule` for every entry in `$requiredModules`, calls `Start-Log`, logs the returned module statuses, runs `Invoke-CalibreUpdateBackup` inside `try/catch/finally`, logs a failure once at `FATAL` level, prints it with `Write-Error`, calls `Stop-Log`, and exits with 0 or 1.

- **External dependency:** `Get-HiDriveSyncRoot`, `Stop-HiDrive`, `Start-HiDrive`, and `Update-HiDriveUtility` come from the separate [StratoHiDriveUtils](https://github.com/DonGrobione/StratoHiDriveUtils) module (3.x), not from this repo. It is installed and imported under the module name `DonGrobione.StratoHiDriveUtils`. The script also needs 7-Zip and the STRATO HiDrive client. It is installed and updated like DonGrobione.Logging (see **Required modules**).
- **Path resolution:** `Invoke-CalibreUpdateBackup` calls `Get-HiDriveSyncRoot` once and passes the result to `Initialize-CalibreBackupFolder` (`Backup\Calibre`, created when missing) and `Get-CalibreFolderPath` (`PortableApps\Calibre Portable`, must exist). Functions receive all values as parameters; there are no `$script:`-scoped variables.
- **Workflow order matters:** 7-Zip is checked first, and the installer is downloaded *before* HiDrive is stopped, so the time HiDrive is offline stays short. In `Invoke-CalibreUpdateBackup`, the `$hiDriveStopped` flag is set right after `Stop-HiDrive` and cleared after `Start-HiDrive`. The `finally` block uses it to restart HiDrive on failure. Keep this recovery behavior intact when you change the workflow.
- **WhatIf:** the script and every state-changing function support `-WhatIf`. `Stop-HiDrive`, `Start-HiDrive`, and the module update commands do not inherit the script's `$WhatIfPreference`, so they are only called behind an explicit `$PSCmdlet.ShouldProcess()` check. With `-WhatIf`, a missing module is not installed and the script stops.
- **Backups:** 7-Zip creates 1 GB split volumes named `CalibrePortableBackup_yyyy-MM-dd_HH-mm.7z.NNN`. 7-Zip exit codes 0 and 1 (warnings) count as success. `Remove-ExpiredBackup` groups the volumes by their timestamp and deletes whole sets beyond `-CalibreBackupRetention`. If you change the filename or timestamp format, you must also change the `-Timestamp` validation in `New-CalibreBackup` and the regex and `ParseExact` format in `Remove-ExpiredBackup`.
- **Required modules:** `$requiredModules` in the main region lists each module with its official installer URL and update command. DonGrobione.Logging comes first because its update reloads the module, which must happen before `Start-Log`. `Initialize-RequiredModule` imports a module with `Import-Module -PassThru`. If the import fails, `Install-RequiredModule` downloads the installer to `$env:TEMP`, runs it without arguments (both installers default to the current user), and deletes it. Otherwise it runs the update command, reloads the module with `-Force`, and compares the versions. A failed update is non-fatal; a failed installation stops the script. The function runs before logging is available, so it returns a status object (`Status`, `Version`, `Message`, `ErrorRecord`) that the main region logs after `Start-Log`, and a fatal failure is only printed to the terminal.
- **Errors:** functions raise terminating errors with `$PSCmdlet.ThrowTerminatingError()` and `New-ErrorRecord`, and they do not log their own failures. Only the main region logs the final error, so each failure is logged exactly once.
- **Logging:** `Write-Log`, `Start-Log`, and `Stop-Log` come from the separate [DonGrobione.Logging](https://github.com/DonGrobione/Logging) module, which is installed in the user's module folder, not in this repo. `Start-Log -LogDirectory 'Calibre-Update-Backup'` writes one file per run to `<Documents>\Logs\Calibre-Update-Backup\<HOSTNAME>_<timestamp>.log`, and the module keeps the five newest files. Log levels are `DEBUG`, `INFO` (default), `WARN`, `ERROR`, and `FATAL`. Pass caught errors with `-ErrorRecord $_` instead of formatting the exception into the message. The module never throws.

## Project Rules

### Documentation
- Keep `README.md` current: project structure, the purpose of every user-facing file and directory, prerequisites, configuration, and usage.
- Write code, comments, comment-based help, output, log messages, and documentation in English.
- Do not break a sentence across lines in `README.md`, comment-based help, or regular comments. Start new lines only between complete sentences or list items.
- Give every commit a clear message that says what changed and why.
- After code changes, check that examples, file structure, and guidelines still match the codebase.

### Comment-Based Help And Comments
- Each `.ps1`/`.psm1` file must contain exactly one comment-based help block, as the first content in the file. Only an optional `<#PSScriptInfo ... #>` block (see "Gallery Readiness") and an optional `#requires` statement may come before it, in that order.
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

### PowerShell Quality Enforcement
- Use `[CmdletBinding()]` and `[Parameter()]` attributes with validation attributes on every function.
- State-changing functions must declare `SupportsShouldProcess` and call `$PSCmdlet.ShouldProcess()` before making changes.
- Use structured error handling: `try/catch`, `-ErrorAction Stop`, and `$PSCmdlet.ThrowTerminatingError()` for terminating errors in advanced functions.
- Do not use `Write-Host` in module code. Use `Write-Verbose` or `Write-Warning` instead.
- Do not hard-code secrets or machine-specific paths. Derive paths from environment variables, `$PSScriptRoot`, or resolved locations such as the HiDrive sync root.
- Do not add an email address anywhere it is not explicitly required.

### Testing (Pester)
- Pester tests live in `Tests/` and use Pester v4-compatible syntax (`Describe`/`Context`/`It` with `Should Be` notation) so they run in Windows PowerShell 5.1.
- When you create a new script or function, write corresponding tests in `Tests/<Name>.Tests.ps1`, one file per script or module.
- When you modify an existing script or function, update its tests to cover the change.
- After any creation or modification, run the affected test files with `Invoke-Pester -Path Tests/<Name>.Tests.ps1` and report the results to the user. If tests fail, fix the code or the tests before considering the task complete.
- Tests must mock every destructive or external call (HiDrive start/stop, process termination, 7-Zip, downloads, module installation, logging, file deletion) so they never touch the real installation, backups, or log files.

### Gallery Readiness
- The project is not published to the PowerShell Gallery, but it must stay compliant with Gallery publication requirements so it can be published later without restructuring.
- Modules (`.psm1`) need a `.psd1` manifest with `Author = 'DonGrobione'`, an explicit `FunctionsToExport` list (never `'*'`), `PowerShellVersion = '5.1'`, `Tags`, `ProjectUri`, `LicenseUri`, and `ReleaseNotes`.
- Standalone scripts (`.ps1` without a manifest) are published with a `<#PSScriptInfo ... #>` block instead of a manifest, containing the same metadata (`VERSION`, `AUTHOR DonGrobione`, `TAGS`, `PROJECTURI`, `LICENSEURI`, `RELEASENOTES`).
- Use Semantic Versioning (SemVer) for all versions.
- Maintain `README.md`, the license file (currently `LICENSE.md`), and `CHANGELOG.md` at the project root.
- Update the version, release notes, and `CHANGELOG.md` whenever a release-relevant change is made.
- The project is published on GitHub, so use `https://github.com/DonGrobione/Calibre-Update-Backup-Script` as `ProjectUri` and `https://github.com/DonGrobione/Calibre-Update-Backup-Script/blob/main/LICENSE.md` as `LicenseUri`. For a project that is not yet on GitHub, leave both fields empty and never invent placeholder URLs.

### Logging And Error Handling
- The DonGrobione.Logging module is the only component allowed to write log files or emit project log entries. Do not add a project-local logging module or write log files directly.
- Modules and library scripts must not contain logging code, call `Write-Log`, or write log files. They return data, status objects, exception records, or error objects to the calling script.
- The orchestrator or standalone entry-point script decides which events are logged and at what level, and calls `Write-Log` for those events.
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
- Standalone scripts (`.ps1` without a manifest): keep a version number in the comment-based help block and update it (along with `Updated` in `.NOTES`) when behavior changes. Once the script has a `<#PSScriptInfo ... #>` block, its `VERSION` becomes the single source of truth, and the version is removed from comment-based help.

### Project-Specific Rules
- Every entry-point script must check its required modules, install missing ones with their official installers, try to update installed ones, call `Start-Log` before its work, and call `Stop-Log` in a `finally` block.
- Protect critical workflows with `try/catch/finally` and keep recovery behavior such as restarting stopped services or clients.

### Precedence
- If the "PowerShell Quality Enforcement", "Testing (Pester)", or "Gallery Readiness" rules conflict with any other rule in this file, those three sections take precedence.
