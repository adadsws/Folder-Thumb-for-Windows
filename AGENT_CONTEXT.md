# 代理上下文

## 技术架构

项目是面向 Windows PowerShell 5.1 的交互式脚本工具。根目录 `run.bat` 负责设置 UTF-8 code page 并启动 `scripts/set_folder_thumb.ps1`；脚本读取根目录 `config.ini`，用 `System.Drawing` 生成 256×256 ICO，并通过 `desktop.ini` 配置文件夹图标。`scripts/file_cleanup.ps1` 使用 `Microsoft.VisualBasic.FileIO` 实现恢复模式的回收站操作，并保留独立的永久清理函数供缩略图重新生成使用。

用户安装和操作方法见 [README.md](README.md)。

## 项目结构

- `scripts/`：运行源码和文件清理边界。
- `tests/`：仓库合规、启动入口和 pre-push 守卫测试。
- `docs/plans/`：已批准但尚未完成的实施计划。
- `docs/finished_plans/`：已完成计划。
- `reference/`：不参与运行的固定上游 submodule 与重建说明。
- `~archive/`：只在本地 Git 保留、不得远程发布的旧文件。
- `~outputs/`：可审计或复用的本地输出，不纳入 Git。
- `~temp/`：可随时重建的临时文件，不纳入 Git。

## 开发与验证

从项目根目录运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-FileCleanup.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-Launcher.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryComplianceHarness.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryCompliance.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-PrePushGuard.ps1
powershell -NoProfile -Command "[void][scriptblock]::Create((Get-Content -LiteralPath 'scripts/set_folder_thumb.ps1' -Raw -Encoding UTF8))"
```

测试必须隔离 Explorer 和用户文件夹副作用。维护规则和发布边界见 [AGENTS.md](AGENTS.md)。

## 已知问题

- 实际运行会终止并重启 Explorer，未保存的 Explorer 窗口状态可能丢失。
- 图标刷新依赖 Windows Explorer、`desktop.ini` 和文件系统属性，在非 Windows 环境不可用。
- 网络共享等位置可能不支持 Windows 回收站；恢复模式会警告并保留无法回收的文件。
- 当前 `main` 与 `origin/main` 已分叉；源仓库包含本地专用历史，不能直接 push。
