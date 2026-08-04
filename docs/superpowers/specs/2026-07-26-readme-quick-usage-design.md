# README 紧凑开源项目设计

## 目标与受众

将根目录 `README.md` 改为紧凑的英文开源项目说明，面向只想下载和使用工具的 Windows 用户。README 应比纯命令清单多一点背景和配置解释，但不扩展为开发者或社区治理文档。

## 信息架构

README 按以下顺序组织，总体控制在约 55–65 行：

1. 项目标题和简短概述：说明工具会根据文件夹中的第一张受支持图片设置 Windows 文件夹缩略图。
2. `Getting Started`：说明可以下载仓库 ZIP 或使用 `git clone`，保持仓库文件在同一目录，然后双击 `run.bat`。
3. `Requirements`：列出 Windows 10/11、Windows PowerShell 5.1，以及在 Explorer 中使用 Large icons 或 Extra large icons 以便观察缩略图。
4. `Usage`：紧凑说明模式 `0`、`1`、`2`，目标 `1`、`2`、`3`，以及支持的 `.jpg`、`.jpeg`、`.png`、`.bmp`、`.gif`、`.webp` 格式。
5. `Configuration`：提供一份带英文行内注释的完整 `config.ini` 示例，展示 `crop`、`letterbox` 两种 `fit` 值以及多个绝对 `path`。正文明确路径必须为绝对路径、可使用 UTF-8 字符，并说明每个 `path` 会作为独立目标处理。
6. `How It Works`：用一个短段落说明程序选择第一张受支持图片，生成 256×256 `.ico`，写入 `desktop.ini`，随后刷新 Explorer；恢复模式会移除程序生成的文件夹缩略图设置。

不单独设置 `Features`，避免与概述、用法和工作原理重复。

## 明确排除

README 不加入以下内容：

- `Troubleshooting`
- `Development`
- `Contributing`
- contributors、sponsors 或社区治理
- badges、CI、release 或 license 声明
- 仓库内部治理约定、测试命令、pre-push 守卫或脱敏发布流程

这些内容当前缺少对应的用户需求或仓库基础设施，加入后会使 README 超出“紧凑的终端用户文档”范围。

## 文案与示例约束

- README 全文保持英文，语气直接、友好，采用常见开源项目的标题和短段落结构。
- 命令、文件名、配置键和参数保持原始拼写。
- 不承诺脚本没有实现的行为；菜单值和配置规则必须与现有脚本一致。
- `config.ini` 示例不得包含真实用户名、目录或其他隐私值，使用如 `C:\Pictures\Wallpapers` 的通用示例。
- 配置示例中的注释必须符合当前解析逻辑，不得让注释文本被误识别为配置值。

## 变更范围

- 更新根目录 `README.md`。
- 按仓库文档约定在 `CHANGELOG.md` 记录本次 README 扩充。
- 不修改脚本行为、`config.ini` 默认值或其他用户功能。

## 验证

- 对照脚本确认菜单、支持格式、图片选择、图标尺寸、恢复行为和配置规则。
- 检查 README 章节完整、英文一致、示例已脱敏，且没有明确排除的章节。
- 验证示例中的 `fit` 和 `path` 写法可被现有脚本解析。
- 运行适用于文档变更的仓库检查和 pre-push hook。
