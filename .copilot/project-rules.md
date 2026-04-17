# Copilot Project Rules

## 1. Documentation
- Always update the README.md with the current project structure and the creation of new scripts.
- Document every directory and file, including their purpose, in the README.md.
- Every change to documentation and code must have clear commit messages.
- After code changes, verify that all examples, file structure, and guidelines are accurate.

## 2. Coding Standards
- Use descriptive, consistent names (no abbreviations except for loop counters).
- Functions should have a single responsibility, ideally no more than 3 parameters, and clear, action-oriented names.
- Avoid side effects in functions when possible; return early to reduce nesting.
- Group related functionality, keep files and modules focused.
- Always validate inputs at boundaries.

## 3. Performance
- Optimize only after measurement and profiling.
- Focus on optimizing the critical path first.
- Use appropriate data structures, minimize memory allocations, batch operations where possible, and cache expensive operations when needed.
- Never sacrifice readability for performance.

## 4. Script Handling (PowerShell)
- All scripts must be written in English (code, comments, help blocks, and output messages).
- Add new scripts to the Scripts/ directory, use PowerShell Approved Verbs for filenames.
- Always import Modules\Write-Log.psm1.
- Use PowerShell Approved Verbs for function names (verify with Get-Verb).
- Error handling and logging are mandatory; use a top-level try/catch, log in catch, and rethrow with throw to bubble failures to the orchestrator.
- Avoid exit 1 in subscripts. Reserve exit handling for orchestrator or standalone entry points where process exit codes are explicitly required.
- No need to modify the main orchestrator script for new scripts.
- Every script must include a correct and up-to-date comment-based help block (about_Comment_Based_Help).
- The comment-based help block must always be updated to reflect any changes to the script, including the version number, whenever the script is modified.
- The comment-based help block must be placed at the top of the file (module-level/script-level), not inside function bodies.
- Do not use ====== separators (reserved for the main script only).
- For existing scripts: maintain the logging pattern, validate paths with Test-Path, use try-catch for critical operations.

## 5. Copilot Memory
- All Copilot memory relevant to this project must be written to `.copilot/repo-memory.md`, not to the local user-scoped Copilot memory system.
- This ensures memory is shared across all machines via version control.
- Session-only notes may still use session memory, but any insight worth keeping must be persisted to `.copilot/repo-memory.md`.
- When reading project context at the start of a task, always load `.copilot/repo-memory.md` first.

## 6. Security
- Never trust user input; always validate.
- Apply defense in depth, least privilege, and fail securely principles.
- Never store secrets in code — all passwords, usernames, credentials, API keys, tokens, and similar sensitive values must be stored in a `.psd1` file under the `Secrets/` folder and read from there at runtime.
- Secrets file naming must correspond to the consuming script: drop the verb prefix from the script name (e.g., `Mount-NetworkDrives.ps1` → `Secrets/NetworkDrives.psd1`).
- The script must reference the secrets file path explicitly (e.g., via a `$ConfigFilePath` parameter defaulting to the relative `Secrets/<Name>.psd1` path).
- The secrets file must include a comment on the first line naming the script that consumes it (e.g., `# Used by: Scripts\Mount-NetworkDrives.ps1`).
- Log security-relevant events.
- Keep dependencies up to date, use static analysis tools, conduct security reviews.
