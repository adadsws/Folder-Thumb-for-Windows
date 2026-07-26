# Folder-Thumb-for-Windows

Batch-set Windows folder icons to the first image found inside each subfolder.

Instead of the default yellow folder icon, each folder displays its own image — converted to a 256×256 ICO and applied via `desktop.ini`.

## Files

| File | Description |
|---|---|
| `run.bat` | Double-click to launch |
| `set_folder_thumb.ps1` | Core script (called by run.bat) |
| `config.ini` | Fit mode setting and optional path list |
| `config.ini.example` | Sanitized configuration template |

## Usage

Double-click **run.bat**. The menu loops until you choose **0 to exit**:

```
Select mode:
  1  Set folder thumbnails
  2  Restore default folder icons
  0  Exit
```

For both modes you then pick the target location:

```
Select target:
  1  Script directory
  2  Enter a custom path
  3  Read paths from config.ini
```

| Target | Description |
|---|---|
| 1 | All subfolders inside the same directory as `run.bat` |
| 2 | Any absolute folder path, entered interactively |
| 3 | Multiple paths listed in `config.ini` under `[paths]` |

## config.ini

`config.ini` is the local runtime configuration. Keep `config.ini.example`
free of real user names and private paths so it remains safe to publish. When
creating or restructuring local configuration, update the example at the same
time.

```ini
[settings]
; How to fit the image into the 256x256 icon square:
;   crop      - center-crop (fills the square, may cut edges)
;   letterbox - fit whole image, pad remaining area with black
fit=crop

[paths]
; Add one path= line per folder to process (used by target 3).
; Lines starting with ; or [ are ignored.
path=C:\Users\YourName\Pictures\Albums
path=D:\Gallery
```

## How it works

1. For each subfolder the first image file (`.jpg` / `.jpeg` / `.png` / `.bmp` / `.gif` / `.webp`) is found.
2. The image is scaled to 256×256 using the configured fit mode and saved as a timestamped `fi_*.ico` file inside that folder. The timestamp in the filename guarantees Windows always loads the new icon instead of a cached version.
3. A `desktop.ini` is written with `IconResource=fi_*.ico,0`, telling Explorer to use that ICO as the folder icon.
4. Explorer is restarted and a global shell notification is sent so changes take effect immediately.

**Restore mode** removes all generated `fi_*.ico` files and `desktop.ini` files from processed subfolders, reverting them to the default yellow folder icon.

## Requirements

- Windows 10 / 11
- Windows PowerShell 5.1 (built-in, no extra install needed)
- Explorer view set to **Large icons** or **Extra large icons** to see custom folder icons

## Notes

- Each processed subfolder contains two hidden files: `fi_*.ico` and `desktop.ini`. Deleting them reverts the icon to the default — use mode 2 (Restore) instead.
- Re-run mode 1 after adding new subfolders or replacing images; the old ICO is replaced automatically.

## 仓库约定

- `data/` 只存输入数据；发布前必须把其中的隐私值替换为示例值。
- `output/` 只存可重新生成的输出，`~temp/` 只存临时文件；这两个目录不会提交。
- `secrets/` 可保存本地真实配置，但每个文件必须有同名、无真实值的 `.example`。
- 无用或旧文件经确认后移入所属子项目的 `~archived/`，不直接删除。
- 设计或实现前调研的成熟 GitHub 项目放入 `~ref/`。
- `~archived/`、`~ref/` 和 `secrets/` 内的真实文件只在本地跟踪和提交，禁止推送到远程。
- `secrets/**/*.example` 保留原路径，经确认不含真实值后允许推送。
- `docs/superpowers/plans/` 只保留执行中的计划；已完成计划及时移入 `~archived/superpowers-plans/`。
- 当前项目没有因总体积或文件数量而不适合推送的大文件夹；以后出现时，本地保留完整目录，远程脱敏导出用原位置的同名 `.md` 说明替代。
- 除 `output/` 和 `~temp/` 外，仓库内容不得通过 Git ignore 排除。

修改治理文件后可运行：

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryCompliance.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-PrePushGuard.ps1
```

## Push 安全守卫

当前仓库使用版本化的 `.githooks/pre-push` 阻止直接推送包含以下本地专用路径的提交历史：

- `~archived/`
- `secrets/` 中非 `.example` 的文件
- `~ref/`
- `docs/superpowers/plans/`

首次使用此工作树时启用守卫：

```powershell
git config core.hooksPath .githooks
```

如果待发布历史包含上述路径，请从不含这些路径及其历史的独立脱敏导出仓库或无历史 orphan 分支发布，不要绕过守卫。
