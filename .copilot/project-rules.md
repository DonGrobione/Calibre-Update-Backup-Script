# Copilot Project Rules

## 1. Scope and Canonical Locations
- Store project rules only in `.copilot/project-rules.md`.
- Store repository memory only in `.copilot/repo-memory.md`.
- Do not create project rules or repository-memory files in any other path.

## 2. Documentation and Comments
- Keep `README.md` accurate when files, requirements, configuration, or behavior changes.
- Do not insert line breaks in the middle of a sentence in `README.md`, comment-based help, or regular comments.
- Use English for code, comments, comment-based help, log messages, and user-facing script output.
- A script-level comment-based help block may appear only once and must be the first content in the script file.
- Do not use comment-based help blocks inside functions.
- Document every function with a concise, human-readable regular comment immediately above its declaration.

## 3. PowerShell 5.1 Standards
- Keep all maintained PowerShell code compatible with Windows PowerShell 5.1.
- Use approved PowerShell verbs for function names when applicable.
- Use descriptive names, consistent casing, and single-purpose functions.
- Use `Join-Path` for constructed paths and `Test-Path` with an appropriate `-PathType` before dependent file-system operations.
- Use `-ErrorAction Stop` for operations handled by `try`/`catch` blocks.
- Provide actionable error messages and log important operations, failures, and cleanup actions.
- Protect critical workflows with `try`/`catch`/`finally`; preserve recovery behavior such as restarting stopped services or clients.
- Avoid aliases, `Invoke-Expression`, unvalidated external input, and hard-coded secrets.
- Prefer `Set-StrictMode -Version Latest` only after validating that the existing script and its dependencies support it in PowerShell 5.1.

## 4. Validation
- Parse changed PowerShell scripts with Windows PowerShell 5.1 before completing work.
- Update the script-level help version and date whenever a maintained script changes.
- Run focused static checks when available and report any checks that cannot run.

## 5. Repository Memory
- Read `.copilot/repo-memory.md` before changing this repository.
- Keep repository memory concise, factual, and current.
