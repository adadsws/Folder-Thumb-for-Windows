# 仓库全局 AGENTS.md 合规实施计划（已完成）

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 让当前仓库的忽略规则、配置示例、文档和版本记录满足全局 `AGENTS.md` 的可验证要求。

**Architecture:** 以最小改动调整仓库治理文件，不改变文件夹缩略图脚本的运行行为。新增一个 PowerShell 合规测试，自动检查忽略规则、必需文档以及隐私文件示例约束。

**Tech Stack:** Git、PowerShell 5.1、Markdown、INI

## Global Constraints

- `output/` 只存输出，`~temp/` 只存临时文件；两者必须写入 `.gitignore`，不提交。
- `data/` 只存输入；发布时，隐私数据必须替换为示例数据。
- 永远归档而不删除；不得使用删除文件的程序或命令。
- 每个真实隐私文件必须有不含真实值、保留结构与说明的 `.example`。
- 除 `output/` 和 `~temp/` 外，不使用 Git ignore 排除仓库内容。
- 重大更新须更新或创建根目录的 `README.md` 和 `CHANGELOG.md`。
- PowerShell 查看中文文件时显式使用 UTF-8。
- WebUI 本地服务返回的 `api_version` 固定为 `1`；当前项目没有 WebUI。

---

### Task 1: 建立可执行的仓库合规检查

**Files:**
- Create: `tests/Test-RepositoryCompliance.ps1`

**Interfaces:**
- Consumes: 仓库根目录的 `.gitignore`、`README.md`、`CHANGELOG.md`、可选的 `secrets/`。
- Produces: 退出码 `0` 表示合规；发现问题时输出错误并以退出码 `1` 结束。

- [x] **Step 1: 写入初始合规测试**

```powershell
$requiredIgnoreRules = @('output/', '~temp/')
$actualIgnoreRules = Get-Content -LiteralPath $gitIgnorePath -Encoding UTF8 |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ -and -not $_.StartsWith('#') }
```

测试必须断言有效规则恰好为 `output/` 与 `~temp/`，并检查 `README.md`、`CHANGELOG.md` 存在；若 `secrets/` 存在，则其中每个非 `.example` 文件都必须有同名 `.example`。

- [x] **Step 2: 运行测试并确认当前仓库失败**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryCompliance.ps1`

Expected: FAIL，指出 `.gitignore` 缺少 `output/`、`~temp/` 且包含额外规则，并指出缺少 `CHANGELOG.md`。

### Task 2: 修复治理文件并记录规则

**Files:**
- Modify: `.gitignore`
- Modify: `README.md`
- Create: `CHANGELOG.md`
- Create: `AGENTS.md`
- Create: `config.ini.example`

**Interfaces:**
- Consumes: Task 1 的合规断言。
- Produces: 只忽略 `output/` 与 `~temp/` 的 Git 规则，以及供维护者和后续代理遵循的简体中文说明。

- [x] **Step 1: 将 `.gitignore` 改为唯一允许的目录规则**

```gitignore
output/
~temp/
```

- [x] **Step 2: 添加仓库级代理约束**

在 `AGENTS.md` 中明确目录职责、隐私示例、归档代替删除、文档更新、Git 提交和禁止直接推送隐私历史的要求。

- [x] **Step 3: 更新项目文档**

在 `README.md` 增加“仓库约定”章节，说明 `data/`、`output/`、`~temp/`、`secrets/`、`~archived/` 和 `~ref/` 的用途及配置方法；创建 `config.ini.example` 作为脱敏配置模板；创建 `CHANGELOG.md`，记录本次合规调整。

- [x] **Step 4: 运行合规测试**

Run: `powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryCompliance.ps1`

Expected: PASS，输出“仓库合规检查通过”。

- [x] **Step 5: 检查项目脚本语法**

Run: `powershell -NoProfile -Command "[void][scriptblock]::Create((Get-Content -LiteralPath 'set_folder_thumb.ps1' -Raw -Encoding UTF8))"`

Expected: 退出码 `0`。

- [x] **Step 6: 审查 Git 范围并提交**

```powershell
git status --short
git diff --check
git add .gitignore AGENTS.md CHANGELOG.md README.md config.ini.example tests/Test-RepositoryCompliance.ps1 docs/superpowers/plans/2026-07-25-repository-compliance.md
git commit -m "chore: align repository with agent requirements"
```

Expected: 提交仅包含本计划列出的治理、测试和文档文件。
