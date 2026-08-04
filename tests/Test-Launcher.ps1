[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$fixtureRoot = Join-Path (
    Join-Path $repositoryRoot '~temp'
) ('launcher-{0}' -f [guid]::NewGuid().ToString('N'))
$capturePath = Join-Path $fixtureRoot 'powershell-arguments.txt'

[void](New-Item -ItemType Directory -Path $fixtureRoot -Force)
Copy-Item -LiteralPath (Join-Path $repositoryRoot 'run.bat') -Destination $fixtureRoot
Set-Content -LiteralPath (Join-Path $fixtureRoot 'powershell.cmd') -Encoding ASCII -Value @(
    '@echo off'
    '> "%FOLDER_THUMB_CAPTURE%" echo %*'
    'exit /b 0'
)

$previousLocation = Get-Location
$previousCapture = $env:FOLDER_THUMB_CAPTURE
$env:FOLDER_THUMB_CAPTURE = $capturePath
Set-Location -LiteralPath $fixtureRoot
try {
    & cmd.exe /d /c 'run.bat < nul' | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Error ('run.bat exited with code {0}.' -f $LASTEXITCODE)
        exit 1
    }
}
finally {
    Set-Location -LiteralPath $previousLocation
    $env:FOLDER_THUMB_CAPTURE = $previousCapture
}

$capturedArguments = Get-Content -LiteralPath $capturePath -Raw -Encoding ASCII
if ($capturedArguments -notmatch 'scripts[\\/]set_folder_thumb\.ps1') {
    Write-Error (
        'run.bat did not launch scripts/set_folder_thumb.ps1: {0}' -f
        $capturedArguments.Trim()
    )
    exit 1
}

Write-Host 'Launcher check passed.'
