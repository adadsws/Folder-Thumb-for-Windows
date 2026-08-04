# 全局 AGENTS.md 更新合规实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 让当前仓库的自动检查、push 守卫和维护文档完整符合 2026-07-26 更新后的全局 `AGENTS.md`。

**Architecture:** 保留现有源码行为，只扩展 PowerShell 合规检查的可测试入口和 Git 隐藏文件检测，并修正发布指引。使用独立临时 Git 仓库验证 `.git/info/exclude` 不能隐藏本应本地跟踪的文件。

**Tech Stack:** Git、Windows PowerShell 5.1、Markdown

## Global Constraints

- `output/`、`~temp/` 必须写入 `.gitignore`，不纳入 Git。
- 除 `output/` 和 `~temp/` 外，所有内容必须纳入本地 Git，不得通过 `.gitignore`、`.git/info/exclude` 或 global gitignore 排除。
- `~archived/`、真实 `secrets/**`、`~ref/`、`docs/superpowers/plans/` 不得 push；`secrets/**/*.example` 可在脱敏后 push。
- 发布只能使用位于 `output/github-export/`、基于远程 `main` 的独立脱敏导出仓库；只保留并 fast-forward push `main`。
- 修改完成后必须运行验证并创建本地 Git commit。
- 已完成计划必须移入 `~archived/superpowers-plans/`，不得删除。
- 当前项目没有 WebUI；如未来新增，本地服务的 `api_version` 固定为 `1`。

---

### Task 1: 让仓库合规检查可在隔离仓库中验证

**Files:**
- Create: `tests/Test-RepositoryComplianceHarness.ps1`
- Modify: `tests/Test-RepositoryCompliance.ps1`

**Interfaces:**
- Consumes: `Test-RepositoryCompliance.ps1 -RepositoryRoot <绝对路径>`。
- Produces: 被 `.git/info/exclude` 隐藏的非输出文件导致合规检查退出码非零，并在错误中报告相对路径。

- [x] **Step 1: 写入失败的回归测试**

测试在 `~temp/` 下创建唯一临时 Git 仓库，写入合规的 `.gitignore`、`README.md`、`CHANGELOG.md`，再通过 `.git/info/exclude` 隐藏 `private-local.txt`。测试调用主合规脚本的 `-RepositoryRoot` 参数并断言退出码非零且输出包含 `private-local.txt`。

- [x] **Step 2: 运行回归测试并确认失败**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryComplianceHarness.ps1`

Expected: FAIL，因为现有合规脚本没有 `RepositoryRoot` 参数，也不会报告被其他 Git ignore 来源隐藏的文件。

- [x] **Step 3: 实现最小合规检测**

为主合规脚本增加可选 `RepositoryRoot` 参数；调用 `git ls-files --others --ignored --exclude-standard`，只允许结果位于 `output/` 或 `~temp/`，其他路径加入失败列表。同时调用 `git ls-files --others --exclude-standard`，报告所有未跟踪且未忽略的文件。

- [x] **Step 4: 运行回归测试和主检查**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryComplianceHarness.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryCompliance.ps1
```

Expected: 两个命令均以退出码 `0` 结束。

### Task 2: 对齐新版发布约束

**Files:**
- Modify: `.githooks/pre-push.ps1`
- Modify: `tests/Test-PrePushGuard.ps1`
- Modify: `README.md`
- Modify: `CHANGELOG.md`

**Interfaces:**
- Consumes: pre-push 输入的 outgoing commit 范围或 `-CheckPath` 测试路径。
- Produces: `~archived/`、`~ref/`、执行中计划和真实 secrets 被拒绝；脱敏 `secrets/**/*.example` 被允许；失败指引只推荐独立脱敏导出仓库。

- [x] **Step 1: 扩充守卫行为检查**

在既有测试中分别调用 `-CheckPath '~ref/reference/file.txt'` 与 `-CheckPath 'docs/superpowers/plans/active.md'`，断言均被拒绝并报告对应路径。

- [x] **Step 2: 修正 pre-push 失败指引**

将关于 orphan branch 的旧指引替换为 `output/github-export/` 独立脱敏导出仓库流程，保持现有路径判定逻辑不变。

- [x] **Step 3: 同步 README 与 CHANGELOG**

README 明确只有 `output/`、`~temp/` 可以被任一 Git ignore 来源忽略；CHANGELOG 记录本次自动检查与发布指引更新。

- [x] **Step 4: 运行守卫测试**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-PrePushGuard.ps1`

Expected: PASS，输出 `Pre-push guard checks passed.`。

### Task 3: 完整验证、归档与提交

**Files:**
- Move: `docs/superpowers/plans/2026-07-26-global-agents-compliance.md`
- To: `~archived/superpowers-plans/2026-07-26-global-agents-compliance.md`

**Interfaces:**
- Consumes: Task 1 与 Task 2 的所有改动。
- Produces: 工作树中没有执行完毕的活动计划，且本次改动形成一个独立本地 commit。

- [x] **Step 1: 运行完整验证**

Run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryComplianceHarness.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryCompliance.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-PrePushGuard.ps1
powershell -NoProfile -Command "[void][scriptblock]::Create((Get-Content -LiteralPath 'set_folder_thumb.ps1' -Raw -Encoding UTF8))"
git diff --check
```

Expected: 所有命令退出码为 `0`。

- [x] **Step 2: 归档计划并重新检查仓库**

把本计划移入 `~archived/superpowers-plans/`，再运行主合规检查，确认 `docs/superpowers/plans/` 没有遗留文件。

- [x] **Step 3: 创建本地提交**

```powershell
git add .githooks/pre-push.ps1 README.md CHANGELOG.md tests/Test-PrePushGuard.ps1 tests/Test-RepositoryCompliance.ps1 tests/Test-RepositoryComplianceHarness.ps1 ~archived/superpowers-plans/2026-07-26-global-agents-compliance.md
git commit -m "chore: align repository with updated agent rules"
```

Expected: 提交仅包含本计划列出的治理、测试、文档和归档计划文件。
