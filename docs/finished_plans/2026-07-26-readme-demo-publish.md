# README Demo and Main Publish Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a clickable preview image and Bilibili introduction video to the English README, then publish the complete sanitized documentation update to remote `main`.

**Architecture:** Add one standard-Markdown `Demo` section to the source README and record it in the changelog. Copy only README, changelog, and the approved image into the existing sanitized export, amend its unpublished documentation commit, verify and fast-forward push `main`, then archive this completed plan locally.

**Tech Stack:** Markdown, JPEG, Windows PowerShell 5.1, Git

## Global Constraints

- Work directly on `main` as explicitly authorized by the user.
- Keep `README.md` entirely in English and compact.
- Use `https://www.bilibili.com/video/BV1heDdBNEVU` for both preview-image and text links.
- Use `docs/202604100646.mp4_20260726_161435.607.jpg` unchanged at 760×673.
- Do not modify scripts, configuration, tests, or runtime behavior.
- Commit as `adadsws <1410990013@qq.com>`.
- Archive the completed plan under `~archived/superpowers-plans/`.
- Never copy `~archived/`, `~ref/`, `docs/superpowers/plans/`, or private secrets into the export.
- The export, source, and remote must expose only `main`; push must be fast-forward without force and without a PR.

---

### Task 1: Add the README Demo section

**Files:**
- Modify: `README.md`
- Modify: `CHANGELOG.md`
- Verify existing asset: `docs/202604100646.mp4_20260726_161435.607.jpg`

**Interfaces:**
- Consumes: The user-provided JPEG and Bilibili URL from the approved design.
- Produces: A GitHub-renderable clickable preview and redundant text link.

- [x] **Step 1: Insert the Demo section**

Insert this block immediately after the opening description and before `## Getting Started`:

```markdown
## Demo

[![Folder thumbnail preview](docs/202604100646.mp4_20260726_161435.607.jpg)](https://www.bilibili.com/video/BV1heDdBNEVU)

[Watch the introduction video on Bilibili.](https://www.bilibili.com/video/BV1heDdBNEVU)
```

- [x] **Step 2: Update CHANGELOG**

Add this bullet at the top of `## 2026-07-26` → `### 仓库治理`:

```markdown
- 在 `README.md` 中加入可点击的效果预览图和 Bilibili 项目介绍视频。
```

- [x] **Step 3: Verify the source documentation**

Run:

```powershell
$readme = Get-Content -LiteralPath README.md -Encoding UTF8
$readme.Count
$readme | Select-String -Pattern '^## '
rg -n -F 'https://www.bilibili.com/video/BV1heDdBNEVU' README.md
Get-Item -LiteralPath docs/202604100646.mp4_20260726_161435.607.jpg
git diff --check
```

Expected: README has `64` lines; `Demo` is the first level-two heading; the URL occurs twice; the image exists; diff check exits `0`.

- [x] **Step 4: Run source repository checks**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryComplianceHarness.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryCompliance.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-PrePushGuard.ps1
```

Expected: all three commands exit with code `0`.

- [x] **Step 5: Commit the source documentation**

Run:

```powershell
git add README.md CHANGELOG.md docs/superpowers/plans/2026-07-26-readme-demo-publish.md
git commit -m "docs: add README demo video"
```

Expected: one source commit authored as `adadsws <1410990013@qq.com>`.

### Task 2: Update the sanitized export

**Files:**
- Modify: `output/github-export/README.md`
- Modify: `output/github-export/CHANGELOG.md`
- Create: `output/github-export/docs/202604100646.mp4_20260726_161435.607.jpg`

**Interfaces:**
- Consumes: Verified source README, changelog, and unchanged JPEG.
- Produces: One sanitized export commit ahead of `origin/main`.

- [x] **Step 1: Copy the three publishable files**

Run:

```powershell
Copy-Item -LiteralPath README.md -Destination output/github-export/README.md
Copy-Item -LiteralPath CHANGELOG.md -Destination output/github-export/CHANGELOG.md
New-Item -ItemType Directory -Path output/github-export/docs -Force
Copy-Item -LiteralPath docs/202604100646.mp4_20260726_161435.607.jpg -Destination output/github-export/docs/202604100646.mp4_20260726_161435.607.jpg
```

- [x] **Step 2: Confirm export scope and privacy**

Run:

```powershell
git -C output/github-export status --short
git -C output/github-export diff --check
git -C output/github-export diff --name-status
git -C output/github-export ls-files |
    rg '^(~archived/|~ref/|docs/superpowers/plans/|secrets/(?!.*\.example$))' --pcre2
```

Expected: only `README.md`, `CHANGELOG.md`, and the JPEG are changed or untracked; diff check passes; blocked-path scan returns no matches.

- [x] **Step 3: Amend the unpublished export commit**

Run:

```powershell
git -C output/github-export add -- README.md CHANGELOG.md docs/202604100646.mp4_20260726_161435.607.jpg
git -C output/github-export commit --amend --no-edit
```

Expected: export remains exactly one commit ahead of `origin/main`, authored as `adadsws <1410990013@qq.com>`.

- [x] **Step 4: Run export repository checks**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File output/github-export/tests/Test-RepositoryComplianceHarness.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File output/github-export/tests/Test-RepositoryCompliance.ps1 -RepositoryRoot output/github-export
Push-Location output/github-export
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-PrePushGuard.ps1
git diff --check
git status --short
Pop-Location
```

Expected: all checks pass and the export working tree is clean.

### Task 3: Fast-forward push and verify remote main

**Files:**
- No source file changes.
- Remote target: `origin/main`

**Interfaces:**
- Consumes: One verified sanitized export commit.
- Produces: Remote `main` containing the compact README, Demo section, preview JPEG, and changelog updates.

- [x] **Step 1: Refresh and verify fast-forward state**

Run:

```powershell
git -C output/github-export fetch origin main
git -C output/github-export rev-list --count main..origin/main
git -C output/github-export rev-list --count origin/main..main
git -C output/github-export diff --name-only origin/main..main
git -C output/github-export push --dry-run origin main:main
```

Expected: behind count `0`, ahead count `1`, outgoing files are exactly README, changelog, and JPEG, and dry-run succeeds through the actual pre-push hook.

- [x] **Step 2: Push main**

Run:

```powershell
git -C output/github-export push origin main:main
```

Expected: a non-force fast-forward update of remote `main`.

- [x] **Step 3: Verify source, export, and remote**

Run:

```powershell
git -C output/github-export fetch origin main
git -C output/github-export rev-parse main
git -C output/github-export rev-parse origin/main
git -C output/github-export show origin/main:README.md
git -C output/github-export cat-file -e origin/main:docs/202604100646.mp4_20260726_161435.607.jpg
git -C output/github-export ls-remote --heads origin
git branch --format='%(refname:short)'
git -C output/github-export branch --format='%(refname:short)'
git status --short
git -C output/github-export status --short
```

Expected: local export `main` equals `origin/main`; remote README contains `Demo` and both Bilibili links; the JPEG exists; source, export, and remote each expose only `main`; the export is clean and the source lists only this active plan's checkbox progress.

### Task 4: Archive the completed implementation plan

**Files:**
- Move: `docs/superpowers/plans/2026-07-26-readme-demo-publish.md`
- To: `~archived/superpowers-plans/2026-07-26-readme-demo-publish.md`

**Interfaces:**
- Consumes: Completed source documentation and verified remote publication.
- Produces: No completed active plan and one locally tracked archive.

- [x] **Step 1: Complete and move the plan**

Change every task-step checkbox to `[x]`, then move the plan to:

```text
~archived/superpowers-plans/2026-07-26-readme-demo-publish.md
```

- [x] **Step 2: Stage before compliance verification**

Run:

```powershell
git add -A -- docs/superpowers/plans/2026-07-26-readme-demo-publish.md ~archived/superpowers-plans/2026-07-26-readme-demo-publish.md
```

This makes the destination locally tracked before the compliance script inspects untracked paths.

- [x] **Step 3: Verify and commit the archive**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryCompliance.ps1
git diff --cached --check
git commit -m "chore: archive README demo plan"
git status --short
```

Expected: checks pass, the archive commit succeeds, and the source working tree is clean.
