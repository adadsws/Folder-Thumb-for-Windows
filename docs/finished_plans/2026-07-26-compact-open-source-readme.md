# Compact Open-Source README Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Expand the English README into a compact, user-focused open-source project guide with clear setup, usage, configuration, and behavior explanations.

**Architecture:** Replace the current minimal README with six short sections that mirror common open-source user documentation while staying within approximately 55–65 lines. Record the documentation change in the existing Chinese changelog, verify every documented behavior against the PowerShell script, then archive this completed plan locally.

**Tech Stack:** Markdown, INI, Windows PowerShell 5.1, Git

## Global Constraints

- `README.md` must remain entirely in English.
- README must contain: overview, `Getting Started`, `Requirements`, `Usage`, `Configuration`, and `How It Works`.
- README must remain approximately 55–65 lines and must not add a separate `Features` section.
- Configuration must explain `crop`, `letterbox`, multiple absolute `path` entries, and UTF-8 paths.
- README must not add troubleshooting, development, contributing, badges, CI, releases, license claims, sponsors, contributors, governance, tests, or publishing instructions.
- Menu values, supported formats, image processing, and restore behavior must match `set_folder_thumb.ps1`.
- Example paths must be generic and contain no private values.
- No script behavior or default configuration may change.
- The completed plan must move to `~archived/superpowers-plans/`.

---

### Task 1: Publish the compact end-user README

**Files:**
- Modify: `README.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: Interactive menu values and runtime behavior from `run.bat`, `set_folder_thumb.ps1`, and `config.ini`.
- Produces: A compact English README whose configuration example can be copied into `config.ini`.

- [x] **Step 1: Replace README with the approved content**

Replace the whole file with:

````markdown
# Folder-Thumb-for-Windows

Generate Windows folder thumbnails from the first supported image found in each subfolder.

## Getting Started

1. Download the repository as a ZIP file, or clone it:

```powershell
git clone https://github.com/adadsws/Folder-Thumb-for-Windows.git
```

2. Keep `run.bat`, `set_folder_thumb.ps1`, and `config.ini` in the same folder.
3. Double-click `run.bat`.

## Requirements

- Windows 10 or Windows 11
- Windows PowerShell 5.1
- File Explorer set to Large icons or Extra large icons

## Usage

First choose a mode:

- `1` — Set folder thumbnails
- `2` — Restore default folder icons
- `0` — Exit

Then choose a target:

- `1` — Process subfolders in the script directory
- `2` — Enter a custom absolute path
- `3` — Load multiple paths from `config.ini`

Supported images: `.jpg`, `.jpeg`, `.png`, `.bmp`, `.gif`, and `.webp`.

## Configuration

Edit `config.ini` to control image fitting and define reusable target folders:

```ini
[settings]
; crop fills the square and may trim the image edges.
; letterbox keeps the whole image and adds black padding.
fit=crop

[paths]
; Use one absolute UTF-8 path per line.
path=C:\Pictures\Albums
path=D:\Gallery
```

Use `fit=crop` to fill the 256×256 icon or `fit=letterbox` to preserve the complete image. Add one `path=` line for every folder you want target option `3` to process.

## How It Works

The script recursively scans the selected folder's subfolders and uses the first supported image found in each one. It creates a 256×256 `.ico`, writes the folder icon setting to `desktop.ini`, and restarts Explorer to refresh the result. Restore mode removes `desktop.ini` and generated icon files from the processed subfolders.
````

- [x] **Step 2: Verify README structure, language, and scope**

Run:

```powershell
$readme = Get-Content -LiteralPath README.md -Encoding UTF8
$readme.Count
$readme | Select-String -Pattern '^## '
rg -n 'Troubleshooting|Development|Contributing|badge|CI|release|license|sponsor|governance|pre-push' README.md
```

Expected:

- Line count is between `55` and `65`.
- The only level-two headings are `Getting Started`, `Requirements`, `Usage`, `Configuration`, and `How It Works`.
- `rg` exits with code `1` because no excluded topic is present.

- [x] **Step 3: Verify documented behavior against the script**

Run:

```powershell
rg -n "Select mode|Set folder thumbnails|Restore default folder icons|Select target|config.ini|\\.jpg|\\.jpeg|\\.png|\\.bmp|\\.gif|\\.webp|letterbox|256|desktop.ini|Stop-Process -Name explorer" set_folder_thumb.ps1
```

Expected: output confirms all menu values, formats, fitting modes, icon size, `desktop.ini` behavior, and Explorer refresh described by README.

- [x] **Step 4: Update CHANGELOG**

Under `## 2026-07-26` → `### 仓库治理`, add this bullet above the older README bullets:

```markdown
- 将 `README.md` 扩充为紧凑的英文开源项目用法文档，补充下载安装、完整 `config.ini` 配置和工作原理说明。
```

- [x] **Step 5: Run documentation and repository verification**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryComplianceHarness.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryCompliance.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-PrePushGuard.ps1
git diff --check
git status --short
```

Expected: all three PowerShell checks and `git diff --check` exit with code `0`; status lists only `README.md`, `CHANGELOG.md`, and this active plan.

- [x] **Step 6: Commit the documentation**

Run:

```powershell
git add README.md CHANGELOG.md docs/superpowers/plans/2026-07-26-compact-open-source-readme.md
git commit -m "docs: expand compact project README"
```

Expected: one commit authored as `adadsws <1410990013@qq.com>`.

### Task 2: Archive the completed implementation plan

**Files:**
- Move: `docs/superpowers/plans/2026-07-26-compact-open-source-readme.md`
- To: `~archived/superpowers-plans/2026-07-26-compact-open-source-readme.md`

**Interfaces:**
- Consumes: The verified documentation commit from Task 1.
- Produces: A clean active plans directory and a locally tracked, non-publishable plan archive.

- [x] **Step 1: Mark the plan complete and archive it**

Change every task checkbox in this file from `[ ]` to `[x]`, then move the file without deleting it:

```powershell
Move-Item -LiteralPath docs/superpowers/plans/2026-07-26-compact-open-source-readme.md -Destination ~archived/superpowers-plans/2026-07-26-compact-open-source-readme.md
```

- [x] **Step 2: Verify the archived state**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryCompliance.ps1
git diff --check
git status --short
```

Expected: checks exit with code `0`; Git reports the plan as moved from `docs/superpowers/plans/` to `~archived/superpowers-plans/`.

- [x] **Step 3: Commit the archive**

Run:

```powershell
git add docs/superpowers/plans/2026-07-26-compact-open-source-readme.md ~archived/superpowers-plans/2026-07-26-compact-open-source-readme.md
git commit -m "chore: archive compact README plan"
git status --short
```

Expected: the archive commit succeeds and the source repository working tree is clean. Do not push because the user has not requested publishing and the archive and reference paths are local-only.
