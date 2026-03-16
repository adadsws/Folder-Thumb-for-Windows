# Folder-Thumb-for-Windows

Batch-set Windows folder icons to the first image found inside each subfolder.

Instead of the default yellow folder icon, each folder displays its own image — converted to a 256×256 ICO and applied via `desktop.ini`.

## Files

| File | Description |
|---|---|
| `run.bat` | Double-click to launch |
| `set_folder_thumb.ps1` | Core script (called by run.bat) |
| `config.ini` | Fit mode setting and optional path list |

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
