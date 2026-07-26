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

## Repository conventions

- `data/` contains inputs only; replace private values with sample data before publishing.
- `output/` contains reproducible outputs and `~temp/` contains temporary files; neither directory is committed.
- `secrets/` may contain real local configuration, but every private file must have an adjacent, sanitized `.example`.
- Move obsolete files to the relevant project's `~archived/` after confirmation; do not delete them.
- Mature GitHub projects used for implementation research belong in `~ref/`.
- Track and commit `~archived/`, `~ref/`, and real secrets locally, but never push them.
- Sanitized `secrets/**/*.example` files may keep their paths and be pushed after verification.
- Keep only active plans in `docs/superpowers/plans/`; move completed plans to `~archived/superpowers-plans/`.
- This project currently has no folders too large to push. If one is added, keep it locally and replace it with a same-name `.md` description in the sanitized remote export.
- Do not hide repository content through `.gitignore`, `.git/info/exclude`, or a global Git ignore file; only `output/` and `~temp/` may be ignored.

After changing governance files, run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryComplianceHarness.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-RepositoryCompliance.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/Test-PrePushGuard.ps1
```

## Push safety guard

The version-controlled `.githooks/pre-push` blocks outgoing history that contains:

- `~archived/`
- files under `secrets/` other than `.example` files
- `~ref/`
- `docs/superpowers/plans/`

Enable the guard when using this working tree for the first time:

```powershell
git config core.hooksPath .githooks
```

If outgoing history contains any blocked path, publish through an independent sanitized export:

1. Clone the current remote `main` into `output/github-export/`.
2. Copy only publishable files from this working tree, overwriting matching paths in the export.
3. Inspect the export diff and history for private or local-only content, then run the tests and pre-push hook.
4. Keep only `main`; synchronize remote updates and use a fast-forward push without force.
5. After pushing, confirm the source repository, export repository, and remote expose only `main`.

The source repository keeps its private local history and does not merge the sanitized export history. Do not bypass the guard or publish through another branch.
