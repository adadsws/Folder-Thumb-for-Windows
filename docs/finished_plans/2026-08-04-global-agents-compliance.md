# 全局 AGENTS.md 合规整改计划

**目标：** 在不改变文件夹缩略图功能行为的前提下，使当前工作树、项目文档、目录约定、上游引用、自动检查和本地发布守卫满足 2026-08-04 生效的全局 `AGENTS.md`。

**设计：** 以全局规则定义的标准目录替换旧治理目录，同时保留可审计的历史资料和可重建的上游固定版本。通过更新现有 PowerShell 合规测试与 pre-push 守卫，把目录、文档、Git 跟踪和发布边界固化为可重复验证的规则。

**技术栈：** Windows PowerShell 5.1、Git、Git submodule、Markdown、批处理脚本

## 审计结论

- 当前源码功能验证全部通过，但现有合规测试仍验证旧规则，因此“测试通过”不代表符合现行全局规则。
- 旧目录 `~archived/`、`~ref/`、`output/` 分别应迁移为 `~archive/`、`reference/`、`~outputs/`。
- `~ref/` 下三个上游仓库已经锁定完整 commit 且工作区干净，但仓库缺少 `.gitmodules`、许可证说明和重建说明。
- 根目录缺少必需的 `AGENT_CONTEXT.md`；项目特有的测试、结构和发布约束也没有活动的根 `AGENTS.md` 承载。
- 普通源码 `set_folder_thumb.ps1` 位于根目录，不符合“普通源码原则上放入子目录”；`run.bat` 可以继续作为根入口。
- 已完成计划仍位于旧归档目录，未使用 `docs/finished_plans/`。
- `.gitignore` 仍忽略旧 `output/`，且尚未覆盖规则明确不跟踪的 `.worktrees/` 和 `.venv/`。
- 当前 `main` 相对 `origin/main` 为 ahead 38、behind 4；本计划不执行 pull、rebase、push 或远程发布。
- 仓库不存在 `secrets/`，没有触发 `git-crypt`；没有目录达到大文件阈值；项目不是 Python WebUI。

## 已确认决策

用户于 2026-08-04 回复 `v`，采用三个推荐项并授权实施：

1. 将三个现有 Git 链接迁移到 `reference/`，补充 `.gitmodules` 和 `reference/README.md`，保持现有完整 SHA，不联网升级。
2. 将旧 `output/github-export/` 原样移动到 `~outputs/github-export/`，保留后续审计和复用能力。
3. 新建精简的根 `AGENTS.md`，只记录本项目的验证命令、入口结构和禁止直接发布当前私有历史等项目特有规则，不复制全局规则。

## 实施任务

### 任务 1：先用测试表达现行合规规则

**修改文件：**

- `tests/Test-RepositoryCompliance.ps1`
- `tests/Test-RepositoryComplianceHarness.ps1`
- `tests/Test-PrePushGuard.ps1`

**实施：**

- 将允许的 Git ignore 路径精确改为 `~outputs/`、`~temp/`、`.worktrees/`、`.venv/`。
- 将 `README.md`、`CHANGELOG.md`、`AGENT_CONTEXT.md` 加入必需文档检查；选择 3a 时也检查根 `AGENTS.md`。
- 增加旧目录名不存在、普通源码不滞留根目录、已完成计划位于 `docs/finished_plans/` 的断言。
- 选择 1a 时，检查 `.gitmodules` 的三个路径、URL、Git link 类型和锁定 SHA，并检查 `reference/README.md` 包含来源、完整 SHA、MIT 许可证和重建命令。
- 将 pre-push 路径测试改为拒绝 `~archive/`、`~outputs/` 和非 example 的 `secrets/**`，允许 `reference/`、`docs/plans/`、`docs/finished_plans/` 与脱敏 example。
- 先运行三个测试并确认它们因旧目录和缺失文档而失败，再开始迁移。

### 任务 2：迁移目录、入口和固定上游

**创建或移动文件：**

- `set_folder_thumb.ps1` → `scripts/set_folder_thumb.ps1`
- `~archived/AGENTS.md` → `~archive/AGENTS.md`
- `~archived/config.ini.example` → `~archive/config.ini.example`
- `~archived/superpowers-plans/*.md` → `docs/finished_plans/*.md`
- 选择 1a：`~ref/{PowerToys,mattpocock-skills,winutil}` → `reference/{PowerToys,mattpocock-skills,winutil}`
- 选择 1a：创建 `.gitmodules` 与 `reference/README.md`
- 选择 1b：将三个旧引用移动到 `~archive/reference/`
- 选择 2a：`output/github-export/` → `~outputs/github-export/`
- 选择 2b：核对后删除 `output/github-export/`

**固定上游元数据（选择 1a）：**

| 路径 | 上游 URL | 锁定 SHA | 许可证 |
|---|---|---|---|
| `reference/PowerToys` | `https://github.com/microsoft/PowerToys.git` | `fc680d350f74f1f4eec8a64296420e614724e74d` | MIT |
| `reference/mattpocock-skills` | `https://github.com/mattpocock/skills.git` | `ed37663cc5fbef691ddfecd080dff42f7e7e350d` | MIT |
| `reference/winutil` | `https://github.com/ChrisTitusTech/winutil.git` | `50be7390e586f664df76d7fed41fc3c39252288c` | MIT |

**入口调整：**

- 修改 `run.bat`，从根目录调用 `scripts/set_folder_thumb.ps1`。
- 修改脚本，使运行目录仍以项目根目录为基准读取 `config.ini`，保持双击入口和三种目标选择的公开行为不变。
- 迁移前后分别执行 PowerShell 语法解析；通过替身脚本或最小可测试参数验证 `run.bat` 指向新路径，避免启动交互式 Explorer 流程。

### 任务 3：补齐职责分离的项目文档

**创建或修改文件：**

- 创建 `AGENT_CONTEXT.md`
- 选择 3a：创建 `AGENTS.md`
- 修改 `README.md`
- 修改 `CHANGELOG.md`

**内容边界：**

- `README.md` 继续面向用户，更新安装文件布局和运行方法，不加入内部代理流程。
- `AGENT_CONTEXT.md` 只记录技术架构、项目结构、开发验证流程、Windows/Explorer 副作用和已知问题，并链接 README 的用户说明。
- `AGENTS.md`（如选择 3a）只记录本项目专属的验证命令、不得在测试中真实重启 Explorer、当前分支历史不得直接 push 等约束，并链接其他文档。
- `CHANGELOG.md` 新增 `2026-08-04` 小节，记录目录治理、固定引用、入口移动、文档和自动检查变化。

### 任务 4：更新自动检查与发布边界

**修改文件：**

- `.gitignore`
- `.githooks/pre-push.ps1`
- `.githooks/pre-push`
- 必要时 `.gitattributes`

**实施：**

- `.gitignore` 仅保留已经明确不跟踪的 `~outputs/`、`~temp/`、`.worktrees/`、`.venv/`；只有实际出现工具原地缓存时才增加精确且带相邻说明的规则。
- 将本地专用路径识别更新为 `~archive/`、`~outputs/` 和真实 `secrets/**`；不再阻止现行规则要求跟踪和可发布的计划与 `reference/`。
- 将失败指引更新为从 `~temp/github-export/` 执行脱敏导出，继续禁止从当前源仓库直接发布本地专用内容。
- 保持 shell 包装器和 PowerShell 实现的 LF/UTF-8 约束。

### 任务 5：完整验证、完成计划和提交

**验证命令：**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryComplianceHarness.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryCompliance.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-PrePushGuard.ps1
powershell -NoProfile -Command "[void][scriptblock]::Create((Get-Content -LiteralPath 'scripts/set_folder_thumb.ps1' -Raw -Encoding UTF8))"
git -c safe.directory=D:/CODE/PY/Folder-Thumb-for-Windows diff --check
git -c safe.directory=D:/CODE/PY/Folder-Thumb-for-Windows status --short
```

**通过标准：**

- 三个测试退出码均为 `0`，脚本语法解析成功，`git diff --check` 无输出。
- `git status` 只包含本计划列出的整改文件，没有修改用户无关内容。
- 选择 1a 时，三个 submodule 工作区干净，Git link SHA 与表格一致，`.gitmodules` 可用于重新初始化。
- 没有真实 secret、仓库外密钥、`~archive/` 或 `~outputs/` 内容进入远程发布流程。
- 将本计划从 `docs/plans/` 移至 `docs/finished_plans/`，再次运行合规检查。
- 创建一个只包含本次整改的本地 commit；不 push，不处理 `main` 与 `origin/main` 的分叉。

## 明确不在范围内

- 不改变缩略图生成、恢复图标、Explorer 重启或配置语义。
- 不升级三个上游仓库，不访问 GitHub 调研，不改写 Git 历史。
- 不创建 tag、固定 worktree、脱敏导出仓库或远程发布；这些操作必须由后续明确请求和确认触发。
- 不启用 `git-crypt`，因为当前仓库没有真实 `secrets/**`。
