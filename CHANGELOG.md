# 更新日志

本项目的主要变更记录在此文件中。

## 2026-08-04

### 仓库治理

- 将运行脚本移入 `scripts/`，保留根目录 `run.bat` 入口及原有配置行为。
- 将旧归档、输出和上游引用目录迁移到 `~archive/`、`~outputs/` 和 `reference/`。
- 为三个上游引用补充可重建的 submodule 元数据、完整锁定 SHA 和许可证说明。
- 新增 `AGENTS.md` 与 `AGENT_CONTEXT.md`，分离项目代理规则、技术架构和用户说明。
- 将已完成计划迁移到 `docs/finished_plans/`，并更新合规检查、入口测试与 pre-push 发布边界。
- 恢复默认图标时将 `desktop.ini`、生成图标和旧 `folder.ico` 送入 Windows 回收站。
- 回收失败时保留原文件并显示警告，不回退为永久删除；重新生成时的旧临时图标仍直接清理。
- 脱敏发布仓库不再需要创建不可发布的空 `~archive/` 目录；已有归档内容仍由 Git 跟踪和 pre-push 守卫保护。
- pre-push shell 包装器移除对外部 `dirname` 的依赖，在精简 `PATH` 的 Git for Windows 环境中仍可定位守卫脚本。
- 合规检查要求 pre-push shell 包装器以 `100755` 记录，避免 Unix 克隆忽略不可执行的 hook。

## 2026-07-26

### 仓库治理

- 将与 `config.ini` 内容重复的 `config.ini.example` 移入本地归档，并从发布仓库移除。
- 将 `README.md` 中的可点击预览图调整为响应式 20% 宽度。
- 在 `README.md` 中加入可点击的效果预览图和 Bilibili 项目介绍视频。
- 将 `README.md` 扩充为紧凑的英文开源项目用法文档，补充下载安装、完整 `config.ini` 配置和工作原理说明。
- 为精简后的 `README.md` 补充菜单、目标、支持格式和 Explorer 图标视图说明。
- 将 `README.md` 精简为英文的运行要求、使用步骤和最小配置示例。
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
