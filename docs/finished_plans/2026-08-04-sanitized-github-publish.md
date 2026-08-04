# GitHub 脱敏发布计划

**目标：** 将当前允许公开的项目状态通过 `~temp/github-export/` 安全发布到 `origin/main`，不推送源仓库的本地专用历史。

**设计：** 从最新远端 `main` 创建独立导出仓库，清空其工作树后写入当前源提交的允许公开文件，并显式恢复三个 submodule Git link。导出仓库运行全部测试、敏感信息检查与 pre-push 守卫，只做 fast-forward push；源仓库保持私有历史和分叉状态。

**技术栈：** Git、GitHub、Windows PowerShell 5.1、Git submodule

## 已确认决策

- 用户明确请求 `push`，授权向配置的 `origin` 远程发布。
- 项目规则禁止直接推送源仓库，因此固定使用 `~temp/github-export/`。
- 目标分支采用远端现有默认分支 `main`，不创建分支、PR 或 force push。
- `~archive/`、`~outputs/`、`~temp/` 和真实 secrets 不进入导出仓库。
- `.gitmodules`、三个 Git link、`reference/README.md` 和允许公开的普通源码、测试、文档必须发布。

## 任务 1：修正发布仓库的合规检查

**文件：**

- 修改 `tests/Test-RepositoryCompliance.ps1`
- 修改 `CHANGELOG.md`

**步骤：**

1. 在不含 `~archive/` 的 `~temp/github-export/` 运行主合规检查，确认当前检查因强制要求本地归档目录而失败。
2. 从必需项目路径中移除 `~archive`；该目录存在时仍由 Git 跟踪与 pre-push 守卫约束，不存在时视为没有待归档内容。
3. CHANGELOG 的 `2026-08-04` 节记录脱敏导出仓库不再需要创建空的本地归档目录。
4. 运行源仓库全部测试并创建本地提交 `chore: support sanitized publication exports`。

## 任务 2：建立独立脱敏导出仓库

**路径：**

- 创建 `~temp/github-export/`
- 创建可重建的临时归档 `~temp/github-export-tree.tar`

**步骤：**

1. 执行 `git fetch origin main`，确认远端可访问并记录最新 `origin/main`；如果认证或网络失败则停止。
2. 核对 `~temp/github-export/` 不存在；若存在，只在验证绝对路径位于当前项目 `~temp/` 后删除这个可重建目录。
3. 从 `origin` 克隆 `main` 到 `~temp/github-export/`，确认导出仓库 HEAD 等于刚获取的远端 SHA。
4. 在导出仓库清除旧工作树内容；从源仓库当前 HEAD 创建排除 `~archive/` 的归档并解压到导出仓库。
5. `git add -A` 后使用 `git update-index --add --cacheinfo 160000,<SHA>,<path>` 固定三个 Git link：
   - `reference/PowerToys` → `fc680d350f74f1f4eec8a64296420e614724e74d`
   - `reference/mattpocock-skills` → `ed37663cc5fbef691ddfecd080dff42f7e7e350d`
   - `reference/winutil` → `50be7390e586f664df76d7fed41fc3c39252288c`
6. 初始化 submodule，确认三个上游可访问、HEAD 与 Git link 一致且工作区干净。
7. 配置导出仓库 `core.hooksPath=.githooks`，确保 `.githooks/pre-push` 为 `100755`，创建脱敏导出提交。

## 任务 3：验证并首次发布

**验证命令：**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-FileCleanup.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-Launcher.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryComplianceHarness.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryCompliance.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-PrePushGuard.ps1
powershell -NoProfile -Command "[void][scriptblock]::Create((Get-Content -LiteralPath 'scripts/file_cleanup.ps1' -Raw -Encoding UTF8)); [void][scriptblock]::Create((Get-Content -LiteralPath 'scripts/set_folder_thumb.ps1' -Raw -Encoding UTF8))"
git diff --check HEAD^
```

**步骤：**

1. 在导出仓库运行上述全部验证。
2. 检查所有 refs 的路径和当前树，确认不存在 `~archive/`、`~outputs/`、`~temp/`、真实 `secrets/**`、凭据或密钥文件；检查 submodule URL、SHA 和许可证说明。
3. 确认 `main` 只包含本次尚未发布的脱敏提交，执行普通 `git push origin main`；若非 fast-forward、hook、认证或网络失败则停止，不 force push。
4. fetch 后确认远端 `main` SHA 等于导出仓库 HEAD。

## 任务 4：完成计划并同步远端

**步骤：**

1. 首次发布成功后，将本计划移动到 `docs/finished_plans/`，运行源仓库全部验证并提交 `chore: finish sanitized github publish`。
2. 用源仓库新 HEAD 再次刷新导出工作树，只产生计划路径移动和相关 Git 元数据差异。
3. 在导出仓库重新运行合规检查、pre-push 测试、语法和 `git diff --check`，提交 `Finish sanitized GitHub publication plan`。
4. 确认导出分支仍为远端 `main` 的 fast-forward，执行第二次普通 push。
5. fetch 并核对远端 SHA、分支列表和导出仓库干净状态；源仓库不得与导出历史合并。

## 失败边界

- 不使用失效的 `gh` 会话创建 PR；本次只发布远端 `main`。
- 不直接 push 当前源仓库，不 pull/rebase 源 `main`，不改写历史。
- 不 force push，不删除远端分支，不发布任何本地专用目录。
- 任一测试、隐私检查、submodule 初始化、hook 或远端核对失败都停止发布并保留 `~temp/github-export/` 供审计。
