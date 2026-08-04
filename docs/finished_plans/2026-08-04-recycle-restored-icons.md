# 恢复图标文件进入回收站实施计划

**目标：** 恢复默认文件夹图标时，将 `desktop.ini`、`fi_*.ico` 和 `folder.ico` 送入 Windows 回收站；重新生成缩略图时清理旧 `fi_*.ico` 仍永久删除。

**设计：** 新建无交互顶层副作用的 `scripts/file_cleanup.ps1`，分别提供回收站和永久清理函数。主脚本按业务场景调用对应函数；回收失败时警告、保留原文件并继续处理，绝不回退为永久删除。

**技术栈：** Windows PowerShell 5.1、`.NET Microsoft.VisualBasic.FileIO`、PowerShell 隔离测试

## 已确认决策

- 用户选择仅在“恢复默认图标”时使用回收站。
- 重新生成缩略图时的旧 `fi_*.ico` 继续永久清理，避免回收站积累临时文件。
- 不支持回收站或回收失败时保留文件并报告警告，不自动永久删除。

## 全局约束

- 自动化测试不得真实终止或启动 Explorer，不得修改用户文件夹图标或真实回收站。
- 保持 Windows PowerShell 5.1 兼容，不增加外部模块依赖。
- 不改变菜单、配置、缩略图生成和 Explorer 刷新行为。
- 同步更新 README、CHANGELOG、AGENT_CONTEXT，并运行项目全部验证命令。
- 完成后将本计划移动到 `docs/finished_plans/` 并创建本地 commit；不 push。

## 任务 1：用隔离测试定义两种删除行为

**文件：**

- 创建 `tests/Test-FileCleanup.ps1`
- 创建 `scripts/file_cleanup.ps1`

**接口：**

- `Move-FileToRecycleBin -LiteralPath <string> [-RecycleAction <scriptblock>]`：成功返回 `$true`；失败写 warning、返回 `$false` 并保留文件。
- `Remove-FilePermanently -LiteralPath <string>`：使用 `Remove-Item -LiteralPath -Force` 永久删除存在的文件。

**步骤：**

1. 创建失败测试：在 `~temp/` 建立源文件和替身回收目录；向 `Move-FileToRecycleBin` 注入只把文件移动到替身目录的 `RecycleAction`，断言源文件消失、替身目录出现文件且返回 `$true`。
2. 增加失败场景：注入抛出异常的 `RecycleAction`，断言返回 `$false` 且源文件仍存在。
3. 增加永久清理场景：调用 `Remove-FilePermanently`，断言临时文件被删除。
4. 运行 `powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-FileCleanup.ps1`，确认因函数脚本不存在而失败。
5. 最小实现 `scripts/file_cleanup.ps1`：默认回收动作调用 `Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile`，使用 `UIOption.OnlyErrorDialogs` 与 `RecycleOption.SendToRecycleBin`；异常只由函数捕获并转为 warning/`$false`。
6. 重新运行测试并确认通过。

## 任务 2：主脚本按场景调用正确删除方式

**文件：**

- 修改 `scripts/set_folder_thumb.ps1`
- 修改 `tests/Test-FileCleanup.ps1`

**步骤：**

1. 在 `~temp/` 复制主脚本和清理脚本，创建含 `desktop.ini`、`fi_old.ico`、`folder.ico` 的隔离项目；子 PowerShell 进程用顺序输入替身选择“恢复 + 项目目录 + 退出”，并替换 `attrib`、Explorer 启停和等待命令。
2. 向隔离运行注入只移动到替身回收目录的 `RecycleAction`，断言三类文件都离开源目录并出现在替身回收目录；测试不得访问真实回收站或 Explorer。运行测试并确认主脚本仍永久删除文件，因此失败。
3. 在主脚本开头 dot-source `scripts/file_cleanup.ps1`。
4. 将恢复分支的三处永久删除替换为 `Move-FileToRecycleBin`，保持移除隐藏/系统属性的顺序；根据布尔返回值输出 `Recycled` 或保留文件的 warning。
5. 将生成分支清理旧 `fi_*.ico` 改为 `Remove-FilePermanently`，明确该处仍为永久清理。
6. 运行文件清理测试与 PowerShell 语法解析并确认通过。

## 任务 3：同步用户和开发文档

**文件：**

- 修改 `README.md`
- 修改 `CHANGELOG.md`
- 修改 `AGENT_CONTEXT.md`
- 修改 `AGENTS.md`

**步骤：**

1. README 的工作原理说明改为：恢复模式将图标配置文件送入回收站；重新生成时的旧生成图标仍永久清理。
2. CHANGELOG 的 `2026-08-04` 节记录可恢复删除与失败不回退的安全行为。
3. AGENT_CONTEXT 的架构和已知问题记录 `Microsoft.VisualBasic.FileIO` 依赖、网络共享等位置可能不支持回收站。
4. AGENTS 增加项目特有规则：测试必须使用 `RecycleAction` 替身，不得操作真实回收站。

## 任务 4：验证、归档与提交

**验证命令：**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-FileCleanup.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-Launcher.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryComplianceHarness.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryCompliance.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-PrePushGuard.ps1
powershell -NoProfile -Command "[void][scriptblock]::Create((Get-Content -LiteralPath 'scripts/file_cleanup.ps1' -Raw -Encoding UTF8)); [void][scriptblock]::Create((Get-Content -LiteralPath 'scripts/set_folder_thumb.ps1' -Raw -Encoding UTF8))"
git -c safe.directory=D:/CODE/PY/Folder-Thumb-for-Windows diff --check
```

**通过标准：**

- 文件清理测试覆盖回收成功、回收失败保留、永久清理和主脚本路由，且不访问真实回收站。
- 原有四项测试继续通过，两个 PowerShell 脚本语法有效，diff 检查无错误。
- 将计划移至 `docs/finished_plans/2026-08-04-recycle-restored-icons.md` 后重新运行合规检查。
- 工作树只包含本计划列出的修改，创建一个本地 commit，不执行 push。
