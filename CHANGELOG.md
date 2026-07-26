# 更新日志

本项目的主要变更记录在此文件中。

## 2026-07-26

### 仓库治理

- 修复 Git 中文路径转义导致 pre-push 守卫漏判本地专用文件、合规检查误判输出文件的问题。
- 增加隔离仓库回归测试，确保 `.git/info/exclude` 或 global gitignore 不能隐藏应由本地 Git 跟踪的文件。
- 扩充 pre-push 守卫验证，覆盖 `~ref/` 与执行中的 Superpowers 计划，并移除 orphan branch 发布指引。
- 将脱敏发布流程收敛为 `output/github-export/` 中仅含 `main` 的独立导出仓库和 fast-forward push。
- 将原为英文的 `README.md` 中新增的中文章节统一改为英文，保持文件内语言一致。
- 将发布说明更新为基于远程 `main` 的独立脱敏导出流程。
- 同步更新后的 secrets 规则：允许已脱敏的 `secrets/**/*.example` 推送，继续阻止真实 secrets 文件。
- 记录大文件夹的本地保留与远程同名 Markdown 替代规则；当前项目无此类目录。
- 新增版本化 pre-push 守卫，阻止直接推送包含本地专用路径的历史。
- 更新合规检查：项目可仅使用全局 `AGENTS.md`，无需在根目录复制一份。
- 明确禁止推送 `~archived/`、`secrets/` 和 `~ref/` 及其历史。
- 将已完成的 Superpowers 实施计划移入 `~archived/superpowers-plans/`。
- 将 `.gitignore` 收敛为仅忽略 `output/` 和 `~temp/`。
- 新增仓库级 `AGENTS.md`，固化目录、隐私、归档、文档、Git 提交和安全发布约束。
- 新增 `config.ini.example`，提供不含真实路径的配置结构示例。
- 新增 PowerShell 仓库合规检查。
- 在 `README.md` 中补充配置方法和仓库目录约定。
