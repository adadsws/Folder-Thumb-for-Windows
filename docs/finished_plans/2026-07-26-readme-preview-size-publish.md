# README Preview Size Publish Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Display the clickable README preview image at 20% of the README content width and publish the sanitized update to remote `main`.

**Architecture:** Replace one Markdown image link with an HTML anchor and image carrying `width="20%"`, while retaining the existing text link. Record the change, verify the source, publish only README and changelog through the sanitized export, then archive this plan.

**Tech Stack:** Markdown, HTML, Windows PowerShell 5.1, Git

## Global Constraints

- Work directly on `main` as previously authorized.
- Keep README text and HTML attributes in English.
- Preserve the existing image file, path, alternative text, Bilibili URL, and text link.
- Use exactly `width="20%"`; do not use a fixed pixel width.
- Do not modify scripts, configuration, tests, or the JPEG.
- Commit as `adadsws <1410990013@qq.com>`.
- Push only a sanitized, fast-forward `main` update without force, another branch, or a PR.
- Never publish `~archived/`, `~ref/`, `docs/superpowers/plans/`, or private secrets.
- Archive this completed plan under `~archived/superpowers-plans/`.

---

### Task 1: Resize the README preview

**Files:**
- Modify: `README.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: Existing clickable preview Markdown in the `Demo` section.
- Produces: A responsive HTML image that remains linked to the same video.

- [x] **Step 1: Replace the preview markup**

Replace:

```markdown
[![Folder thumbnail preview](docs/202604100646.mp4_20260726_161435.607.jpg)](https://www.bilibili.com/video/BV1heDdBNEVU)
```

with:

```html
<a href="https://www.bilibili.com/video/BV1heDdBNEVU"><img src="docs/202604100646.mp4_20260726_161435.607.jpg" alt="Folder thumbnail preview" width="20%"></a>
```

- [x] **Step 2: Update CHANGELOG**

Add this bullet at the top of `## 2026-07-26` → `### 仓库治理`:

```markdown
- 将 `README.md` 中的可点击预览图调整为响应式 20% 宽度。
```

- [x] **Step 3: Verify source content**

Run:

```powershell
rg -n -F 'width="20%"' README.md
rg -n -F 'https://www.bilibili.com/video/BV1heDdBNEVU' README.md
rg -n -F 'alt="Folder thumbnail preview"' README.md
git diff --check
```

Expected: width occurs once, video URL occurs twice, alt text occurs once, and diff check passes.

- [x] **Step 4: Run source checks and commit**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryComplianceHarness.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryCompliance.ps1
'' | powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-PrePushGuard.ps1
git add README.md CHANGELOG.md docs/superpowers/plans/2026-07-26-readme-preview-size-publish.md
git commit -m "docs: resize README preview image"
```

Expected: all checks pass and the source documentation commit succeeds.

### Task 2: Publish through the sanitized export

**Files:**
- Modify: `output/github-export/README.md`
- Modify: `output/github-export/CHANGELOG.md`

**Interfaces:**
- Consumes: Verified source documentation from Task 1.
- Produces: One fast-forward sanitized export commit on remote `main`.

- [x] **Step 1: Synchronize export and copy publishable files**

Run:

```powershell
git -C output/github-export fetch origin main
git -C output/github-export pull --ff-only origin main
Copy-Item -LiteralPath README.md -Destination output/github-export/README.md
Copy-Item -LiteralPath CHANGELOG.md -Destination output/github-export/CHANGELOG.md
```

Expected: export begins at remote `main`; only README and changelog change.

- [x] **Step 2: Verify scope, privacy, and content**

Run:

```powershell
git -C output/github-export status --short
git -C output/github-export diff --check
git -C output/github-export diff --name-only
git -C output/github-export ls-files |
    rg '^(~archived/|~ref/|docs/superpowers/plans/|secrets/(?!.*\.example$))' --pcre2
rg -n -F 'width="20%"' output/github-export/README.md
```

Expected: only README and changelog differ; blocked-path scan has no matches; width occurs once.

- [x] **Step 3: Commit and verify export**

Run:

```powershell
git -C output/github-export add -- README.md CHANGELOG.md
git -C output/github-export commit -m "Scale README preview image"
powershell -NoProfile -ExecutionPolicy Bypass -File output/github-export/tests/Test-RepositoryComplianceHarness.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File output/github-export/tests/Test-RepositoryCompliance.ps1 -RepositoryRoot output/github-export
Push-Location output/github-export
'' | powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-PrePushGuard.ps1
git diff --check
git status --short
Pop-Location
```

Expected: export commit is authored as `adadsws <1410990013@qq.com>`; all checks pass; export is clean.

- [x] **Step 4: Dry-run and push**

Run:

```powershell
git -C output/github-export fetch origin main
git -C output/github-export rev-list --count main..origin/main
git -C output/github-export rev-list --count origin/main..main
git -C output/github-export diff --name-only origin/main..main
git -C output/github-export push --dry-run origin main:main
git -C output/github-export push origin main:main
```

Expected: export is zero commits behind and one commit ahead; outgoing paths are README and changelog; actual push is fast-forward.

- [x] **Step 5: Verify remote publication**

Run:

```powershell
git -C output/github-export fetch origin main
git -C output/github-export rev-parse main
git -C output/github-export rev-parse origin/main
git -C output/github-export show origin/main:README.md
git -C output/github-export ls-remote --heads origin
git branch --format='%(refname:short)'
git -C output/github-export branch --format='%(refname:short)'
```

Expected: export `main` equals remote `main`; remote README contains one `width="20%"`; source, export, and remote expose only `main`.

### Task 3: Archive the completed plan

**Files:**
- Move: `docs/superpowers/plans/2026-07-26-readme-preview-size-publish.md`
- To: `~archived/superpowers-plans/2026-07-26-readme-preview-size-publish.md`

**Interfaces:**
- Consumes: Verified remote publication.
- Produces: A clean source working tree with no completed active plan.

- [x] **Step 1: Complete and move the plan**

Mark every task-step checkbox `[x]`, then move the file to:

```text
~archived/superpowers-plans/2026-07-26-readme-preview-size-publish.md
```

- [x] **Step 2: Stage, verify, and commit the archive**

Run:

```powershell
git add -A -- docs/superpowers/plans/2026-07-26-readme-preview-size-publish.md ~archived/superpowers-plans/2026-07-26-readme-preview-size-publish.md
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryCompliance.ps1
git diff --cached --check
git commit -m "chore: archive README preview plan"
git status --short
```

Expected: checks pass, archive commit succeeds, and the source working tree is clean.
