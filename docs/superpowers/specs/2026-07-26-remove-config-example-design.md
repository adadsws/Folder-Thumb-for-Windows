# 移除重复配置示例设计

## 目标

移除与 `config.ini` 配置文本相同、仅换行编码不同的 `config.ini.example`，让用户只维护和使用一个不含隐私值、可直接运行的配置文件。

## 源仓库处理

- 将 `config.ini.example` 原样移动到 `~archived/config.ini.example`。
- 不删除文件内容，遵守“归档而不删除”的本地规则。
- `~archived/config.ini.example` 继续由本地 Git 跟踪，但不得 push。
- 根目录只保留 `config.ini`；其现有内容使用通用注释示例，不含真实用户路径。

## 发布仓库处理

- 从脱敏导出仓库的根目录删除 `config.ini.example`。
- 保留并发布 `config.ini`，使其成为唯一配置模板和实际运行配置。
- 使用独立导出仓库从远程 `main` fast-forward push，不 force push，不创建分支或 PR。

## 文档处理

- README 当前只引用 `config.ini`，无需修改。
- 在 `CHANGELOG.md` 当前日期下新增说明，记录重复的 `config.ini.example` 已归档并从发布仓库移除。
- 保留 CHANGELOG 中“曾新增 `config.ini.example`”的旧记录，因为它描述的是历史事件。

## 不受影响的内容

- `secrets/**/*.example` 规则保持不变；该规则只处理真实隐私文件的脱敏示例。
- 不修改 `set_folder_thumb.ps1`、`run.bat`、测试或 `config.ini` 内容。
- 不修改其他 `.example` 相关测试与仓库治理规则。

## 验证

- 规范化 CRLF/LF 后确认 `config.ini` 与 `config.ini.example` 文本一致。
- 确认归档文件与移动前的 `config.ini.example` 字节哈希一致。
- 确认源仓库根目录和远程 `main` 均不存在 `config.ini.example`。
- 确认源仓库归档路径存在并由本地 Git 跟踪，导出历史不包含该归档路径。
- 运行源仓库与导出仓库的合规检查和 pre-push 守卫。
- push 后确认源仓库、导出仓库和远程仍仅有 `main`。
