<#
.SYNOPSIS
    Set folder thumbnail to the first image found inside each subfolder.

.DESCRIPTION
    Converts the first image (jpg/jpeg/png/bmp/gif/webp) in each subfolder
    to a 256x256 ICO file and writes a desktop.ini to replace the default
    yellow folder icon with that image.

    Three modes:
      1  Process the directory where this script lives
      2  Enter a path interactively
      3  Read paths from paths.ini (path= lines under [paths])
      4  Restore default folder icons (remove desktop.ini and ico files)

.NOTES
    Requires Windows PowerShell 5.1+ and .NET System.Drawing.
    Run via run.bat (double-click) or directly in PowerShell.
#>

#Requires -Version 5.1

Add-Type -AssemblyName System.Drawing

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$exts      = @('.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp')

# ---------------------------------------------------------------------------
# Path selection helper
# ---------------------------------------------------------------------------
function Get-TargetDirs {
    param([string]$Prompt = "Enter 1 / 2 / 3")
    Write-Host ""
    Write-Host "Select target:"
    Write-Host "  1  Script directory  : $scriptDir"
    Write-Host "  2  Enter a custom path"
    Write-Host "  3  Read paths from config.ini"
    Write-Host ""
    $sel = Read-Host $Prompt
    switch ($sel) {
        '1' { return @($scriptDir) }
        '2' {
            do {
                $p = (Read-Host "Enter path").Trim()
                if ($p -eq '') {
                    Write-Host "Path cannot be empty, please try again."
                } elseif (-not (Test-Path -LiteralPath $p -PathType Container)) {
                    Write-Host "Path not found: $p, please try again."
                    $p = ''
                }
            } while ($p -eq '')
            return @($p)
        }
        '3' {
            $iniFile = Join-Path $scriptDir 'config.ini'
            if (-not (Test-Path -LiteralPath $iniFile)) {
                Write-Host "config.ini not found."
                exit 1
            }
            $result = [IO.File]::ReadAllLines($iniFile, [Text.Encoding]::UTF8) |
                ForEach-Object  { $_.Trim() } |
                Where-Object    { $_ -match '^path\s*=\s*(.+)' } |
                ForEach-Object  { $matches[1].Trim() } |
                Where-Object    { Test-Path -LiteralPath $_ -PathType Container }
            if ($result.Count -eq 0) {
                Write-Host "No valid paths found in config.ini."
                Write-Host "Add  path=C:\your\folder  lines under [paths]."
                exit 1
            }
            Write-Host "Paths loaded from config.ini:"
            $result | ForEach-Object { Write-Host "  $_" }
            return $result
        }
        default { Write-Host "Invalid choice."; exit 1 }
    }
}

# ---------------------------------------------------------------------------
# Main loop
# ---------------------------------------------------------------------------
while ($true) {

Write-Host ""
Write-Host "Select mode:"
Write-Host "  1  Set folder thumbnails"
Write-Host "  2  Restore default folder icons"
Write-Host "  0  Exit"
Write-Host ""
$mode = Read-Host "Enter 0 / 1 / 2"

if ($mode -eq '0') { exit 0 }

if ($mode -eq '2') {
    $dirs = Get-TargetDirs
    Write-Host ""
    $dirs | ForEach-Object {
        Write-Host "Restoring: $_"
        Get-ChildItem -LiteralPath $_ -Directory -Recurse | ForEach-Object {
            $f   = $_.FullName
            $ini = Join-Path $f 'desktop.ini'
            if (Test-Path -LiteralPath $ini) {
                attrib -s -h $ini
                Remove-Item -LiteralPath $ini -Force
                Write-Host "  Removed desktop.ini : $($_.Name)"
            }
            Get-ChildItem -LiteralPath $f -Filter 'fi_*.ico' -Force |
                ForEach-Object { attrib -h $_.FullName; Remove-Item $_.FullName -Force }
            $legacyIco = Join-Path $f 'folder.ico'
            if (Test-Path -LiteralPath $legacyIco) {
                attrib -h $legacyIco
                Remove-Item -LiteralPath $legacyIco -Force
                Write-Host "  Removed folder.ico  : $($_.Name)"
            }
            attrib -s $f
        }
    }
    Write-Host "`nRestarting Explorer..."
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Sleep 2
    $dirs | ForEach-Object { Start-Process explorer $_ }
    Write-Host "Done!"
    continue
}

if ($mode -ne '1') { Write-Host "Invalid choice."; continue }

$dirs = Get-TargetDirs
Read-Host "`nPress Enter to start"

# Read fit mode from config.ini (default: crop)
$fitMode = 'crop'
$cfgFile = Join-Path $scriptDir 'config.ini'
if (Test-Path -LiteralPath $cfgFile) {
    $fitValue = [IO.File]::ReadAllLines($cfgFile, [Text.Encoding]::UTF8) |
        ForEach-Object { if ($_ -match '^fit\s*=\s*(.+)') { $matches[1].Trim() } } |
        Select-Object -First 1
    if ($fitValue -eq 'letterbox') { $fitMode = 'letterbox' }
}
Write-Host "Fit mode: $fitMode"
# ---------------------------------------------------------------------------
# Helper: convert an image to a 256x256 ICO (PNG-compressed, single entry)
# ---------------------------------------------------------------------------
function New-FolderIco {
    param(
        [Parameter(Mandatory)][string]$ImagePath,
        [Parameter(Mandatory)][string]$IcoPath,
        [string]$FitMode = 'crop'   # 'crop' or 'letterbox'
    )

    $src = [System.Drawing.Image]::FromFile($ImagePath)
    $bmp = New-Object System.Drawing.Bitmap(256, 256)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic

    if ($FitMode -eq 'letterbox') {
        # Fill background with black, then draw image scaled to fit entirely
        $g.Clear([System.Drawing.Color]::Black)
        $scale   = [Math]::Min(256.0 / $src.Width, 256.0 / $src.Height)
        $dstW    = [int]($src.Width  * $scale)
        $dstH    = [int]($src.Height * $scale)
        $dstX    = [int]((256 - $dstW) / 2)
        $dstY    = [int]((256 - $dstH) / 2)
        $g.DrawImage(
            $src,
            [System.Drawing.Rectangle]::new($dstX, $dstY, $dstW, $dstH)
        )
    } else {
        # Center-crop: cut the longer axis so the image fills 256x256
        $size = [Math]::Min($src.Width, $src.Height)
        $srcX = [int](($src.Width  - $size) / 2)
        $srcY = [int](($src.Height - $size) / 2)
        $g.DrawImage(
            $src,
            [System.Drawing.Rectangle]::new(0, 0, 256, 256),
            [System.Drawing.Rectangle]::new($srcX, $srcY, $size, $size),
            [System.Drawing.GraphicsUnit]::Pixel
        )
    }
    $g.Dispose(); $src.Dispose()

    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $pngBytes = $ms.ToArray()
    $ms.Dispose(); $bmp.Dispose()

    # ICO file format: 6-byte header + 16-byte directory entry + PNG data
    $fs = New-Object System.IO.FileStream($IcoPath, [System.IO.FileMode]::Create)
    $bw = New-Object System.IO.BinaryWriter($fs)
    $bw.Write([uint16]0)                 # Reserved
    $bw.Write([uint16]1)                 # Type: ICO
    $bw.Write([uint16]1)                 # Image count: 1
    $bw.Write([byte]0)                   # Width  (0 = 256)
    $bw.Write([byte]0)                   # Height (0 = 256)
    $bw.Write([byte]0)                   # Color count
    $bw.Write([byte]0)                   # Reserved
    $bw.Write([uint16]1)                 # Planes
    $bw.Write([uint16]32)                # Bits per pixel
    $bw.Write([uint32]$pngBytes.Length)  # Data size
    $bw.Write([uint32]22)                # Data offset (6 + 16)
    $bw.Write($pngBytes)
    $bw.Dispose(); $fs.Dispose()
}

# ---------------------------------------------------------------------------
# Main processing loop
# ---------------------------------------------------------------------------
foreach ($targetDir in $dirs) {
    Write-Host "`n--- Processing: $targetDir ---"

    Get-ChildItem -LiteralPath $targetDir -Directory -Recurse | ForEach-Object {
        $folder = $_
        $img = Get-ChildItem -LiteralPath $folder.FullName -File |
            Where-Object { $exts -contains $_.Extension.ToLower() } |
            Select-Object -First 1

        if (-not $img) { return }

        $ini = Join-Path $folder.FullName 'desktop.ini'

        # Use a timestamp-based ico filename so Explorer always sees a new file
        # and is forced to re-render (bypasses icon cache entirely)
        $ts  = [DateTime]::Now.ToString('yyyyMMddHHmmss')
        $ico = Join-Path $folder.FullName "fi_$ts.ico"

        # Remove old generated ico files and unprotect desktop.ini
        if (Test-Path -LiteralPath $ini) { attrib -s -h $ini }
        Get-ChildItem -LiteralPath $folder.FullName -Filter 'fi_*.ico' -Force |
            ForEach-Object { attrib -h $_.FullName; Remove-Item $_.FullName -Force }

        New-FolderIco -ImagePath $img.FullName -IcoPath $ico -FitMode $fitMode
        attrib +h $ico

        $icoName = [IO.Path]::GetFileName($ico)
        $content = "[.ShellClassInfo]`r`nIconResource=$icoName,0`r`n"
        [IO.File]::WriteAllText($ini, $content, [Text.Encoding]::Unicode)
        attrib +s +h $ini
        attrib +s $folder.FullName

        Write-Host "  Set: $($folder.Name)  ->  $($img.Name)"
    }
}

# ---------------------------------------------------------------------------
# Refresh Explorer
# ---------------------------------------------------------------------------
Write-Host "`nRestarting Explorer..."
Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
Start-Sleep 2

# Global shell notification - forces all folder icons to re-read
Add-Type -MemberDefinition @'
    [DllImport("shell32.dll", CharSet=CharSet.Auto)]
    public static extern void SHChangeNotify(int wEventId, uint uFlags, IntPtr dwItem1, IntPtr dwItem2);
'@ -Name ShellNotify2 -Namespace Win32 -ErrorAction SilentlyContinue
[Win32.ShellNotify2]::SHChangeNotify(0x08000000, 0x0000, [IntPtr]::Zero, [IntPtr]::Zero)

$dirs | ForEach-Object { Start-Process explorer $_ }
Write-Host "Done!"

} # end while
