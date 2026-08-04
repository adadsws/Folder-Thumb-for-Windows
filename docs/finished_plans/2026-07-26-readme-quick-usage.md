# README Quick Usage Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the current long README with a concise English usage guide while preserving enough information to run and configure the script.

**Architecture:** This is a documentation-only change. `README.md` becomes the user-facing quick-start document; `CHANGELOG.md` records the removal of maintainer-oriented material.

**Tech Stack:** Markdown, INI, Windows PowerShell 5.1

## Global Constraints

- README must remain entirely in English.
- README must contain only the title, one-sentence purpose, requirements, usage steps, and a minimal sanitized `config.ini` example.
- Commands, filenames, configuration keys, and parameters must retain their original spelling.
- The example path must not contain a real username or private directory.
- `CHANGELOG.md` must record the documentation change.
- The completed plan must move to `~archived/superpowers-plans/`.

---

### Task 1: Replace README with the minimal usage guide

**Files:**
- Modify: `README.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: `run.bat`, the interactive menu, and the `[settings]` / `[paths]` structure from `config.ini`.
- Produces: An English quick-start document that lets a Windows user launch the script and optionally configure multiple paths.

- [x] **Step 1: Replace README content**

Write exactly these sections: project title and purpose sentence, `Requirements`, `Usage`, and `config.ini`. The usage section instructs users to run `run.bat`, choose set or restore mode, then choose script directory, custom path, or configured paths. The INI example uses `fit=crop` and `path=C:\Users\YourName\Pictures\Albums`.

- [x] **Step 2: Record the change**

Add a `CHANGELOG.md` entry under `2026-07-26` stating that README now retains only concise English usage, requirements, and configuration instructions.

- [x] **Step 3: Inspect the rendered source**

Run:

```powershell
Get-Content -LiteralPath README.md -Encoding UTF8
```

Expected: no `Files`, `How it works`, `Notes`, `Repository conventions`, or `Push safety guard` sections remain.

- [x] **Step 4: Run repository checks**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryComplianceHarness.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryCompliance.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-PrePushGuard.ps1
git diff --check
```

Expected: all commands exit with code `0`.

- [x] **Step 5: Commit the documentation change**

```powershell
git add README.md CHANGELOG.md docs/superpowers/plans/2026-07-26-readme-quick-usage.md
git commit -m "docs: simplify README usage"
```

Expected: the commit contains only README, CHANGELOG, and the active plan.

### Task 2: Archive the completed plan

**Files:**
- Move: `docs/superpowers/plans/2026-07-26-readme-quick-usage.md`
- To: `~archived/superpowers-plans/2026-07-26-readme-quick-usage.md`

**Interfaces:**
- Consumes: The committed and verified Task 1 result.
- Produces: No completed plan remains in `docs/superpowers/plans/`.

- [x] **Step 1: Mark the plan complete and archive it**

Move this plan to `~archived/superpowers-plans/2026-07-26-readme-quick-usage.md` and mark all checkboxes complete.

- [x] **Step 2: Re-run repository compliance**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryCompliance.ps1
```

Expected: exit code `0`, with no active plan file reported.

- [x] **Step 3: Commit the archive**

```powershell
git add -A docs/superpowers/plans/2026-07-26-readme-quick-usage.md ~archived/superpowers-plans/2026-07-26-readme-quick-usage.md
git commit -m "chore: archive README implementation plan"
```

Expected: a rename-only archive commit plus checkbox state changes.
