# README Usage Details Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a small amount of practical detail to the concise English README without restoring implementation or repository-maintenance material.

**Architecture:** Keep the existing four README sections. Expand requirements and usage with compact lists, then record the refinement in `CHANGELOG.md`.

**Tech Stack:** Markdown, INI, Windows PowerShell 5.1

## Global Constraints

- README must remain entirely in English.
- README must retain exactly four headings: title, `Requirements`, `Usage`, and `config.ini`.
- Requirements must mention Large icons or Extra large icons.
- Usage must define mode choices `0`, `1`, `2`, target choices `1`, `2`, `3`, and supported image formats.
- README must not restore implementation details, repository governance, tests, or push instructions.
- The example path must remain sanitized.
- The completed plan must move to `~archived/superpowers-plans/`.

---

### Task 1: Add practical usage details

**Files:**
- Modify: `README.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: The interactive options implemented by `run.bat` and `set_folder_thumb.ps1`.
- Produces: A concise English README of approximately 35–40 lines.

- [x] **Step 1: Expand requirements and usage**

Add the Explorer icon-view requirement. Add compact `Mode` and `Target` lists with the exact numeric choices. Add the supported image formats `.jpg`, `.jpeg`, `.png`, `.bmp`, `.gif`, and `.webp`.

- [x] **Step 2: Update CHANGELOG**

Add a `2026-07-26` entry stating that README now includes compact menu, target, supported-format, and Explorer-view guidance.

- [x] **Step 3: Verify the README scope**

Run:

```powershell
Get-Content -LiteralPath README.md -Encoding UTF8
rg -n '^##? ' README.md
```

Expected: exactly four headings and approximately 35–40 lines, with no maintainer-oriented sections.

- [x] **Step 4: Run repository checks**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryComplianceHarness.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryCompliance.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-PrePushGuard.ps1
git diff --check
```

Expected: all commands exit with code `0`.

- [x] **Step 5: Commit**

```powershell
git add README.md CHANGELOG.md docs/superpowers/plans/2026-07-26-readme-usage-details.md
git commit -m "docs: add concise README details"
```

Expected: one verified documentation commit.

### Task 2: Archive the completed plan

**Files:**
- Move: `docs/superpowers/plans/2026-07-26-readme-usage-details.md`
- To: `~archived/superpowers-plans/2026-07-26-readme-usage-details.md`

**Interfaces:**
- Consumes: The completed Task 1 commit.
- Produces: No completed plan in the active plans directory.

- [x] **Step 1: Mark complete and archive**

Mark every task checkbox complete and move the plan to `~archived/superpowers-plans/2026-07-26-readme-usage-details.md`.

- [x] **Step 2: Verify and commit the archive**

Run repository compliance, stage the rename, and commit with:

```powershell
git commit -m "chore: archive README details plan"
```

Expected: repository compliance passes and the archive commit succeeds.
