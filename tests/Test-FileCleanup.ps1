[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$cleanupScript = Join-Path $repositoryRoot 'scripts\file_cleanup.ps1'
$mainScript = Join-Path $repositoryRoot 'scripts\set_folder_thumb.ps1'
$fixtureRoot = Join-Path (
    Join-Path $repositoryRoot '~temp'
) ('file-cleanup-{0}' -f [guid]::NewGuid().ToString('N'))

if (-not (Test-Path -LiteralPath $cleanupScript -PathType Leaf)) {
    Write-Error 'scripts/file_cleanup.ps1 is missing.'
    exit 1
}

. $cleanupScript

[void](New-Item -ItemType Directory -Path $fixtureRoot -Force)

$recycleSource = Join-Path $fixtureRoot 'recycle-source.ico'
$fakeRecycle = Join-Path $fixtureRoot 'fake-recycle'
[void](New-Item -ItemType Directory -Path $fakeRecycle -Force)
Set-Content -LiteralPath $recycleSource -Encoding ASCII -Value 'recycle me'
$recycleAction = {
    param([string]$Path)
    Move-Item -LiteralPath $Path -Destination $fakeRecycle
}.GetNewClosure()

$recycleResult = Move-FileToRecycleBin `
    -LiteralPath $recycleSource `
    -RecycleAction $recycleAction
if (-not $recycleResult) {
    Write-Error 'Move-FileToRecycleBin did not report success.'
    exit 1
}
if (Test-Path -LiteralPath $recycleSource) {
    Write-Error 'Recycle source still exists after a successful recycle action.'
    exit 1
}
if (-not (Test-Path -LiteralPath (Join-Path $fakeRecycle 'recycle-source.ico'))) {
    Write-Error 'Recycle action did not receive and move the source file.'
    exit 1
}

$failedSource = Join-Path $fixtureRoot 'failed-recycle.ico'
Set-Content -LiteralPath $failedSource -Encoding ASCII -Value 'keep me'
$previousWarningPreference = $WarningPreference
$WarningPreference = 'SilentlyContinue'
try {
    $failedResult = Move-FileToRecycleBin `
        -LiteralPath $failedSource `
        -RecycleAction { throw 'Recycle unavailable' }
}
finally {
    $WarningPreference = $previousWarningPreference
}
if ($failedResult) {
    Write-Error 'Move-FileToRecycleBin reported success after recycle failure.'
    exit 1
}
if (-not (Test-Path -LiteralPath $failedSource)) {
    Write-Error 'Recycle failure removed the source file.'
    exit 1
}

$permanentSource = Join-Path $fixtureRoot 'permanent.ico'
Set-Content -LiteralPath $permanentSource -Encoding ASCII -Value 'delete me'
Remove-FilePermanently -LiteralPath $permanentSource
if (Test-Path -LiteralPath $permanentSource) {
    Write-Error 'Remove-FilePermanently kept the source file.'
    exit 1
}

$projectFixture = Join-Path $fixtureRoot 'project'
$scriptsFixture = Join-Path $projectFixture 'scripts'
$folderFixture = Join-Path $projectFixture 'album'
$restoreFixture = Join-Path $fixtureRoot 'restore-bin'
[void](New-Item -ItemType Directory -Path $scriptsFixture -Force)
[void](New-Item -ItemType Directory -Path $folderFixture -Force)
[void](New-Item -ItemType Directory -Path $restoreFixture -Force)
Copy-Item -LiteralPath $mainScript -Destination $scriptsFixture
Copy-Item -LiteralPath $cleanupScript -Destination $scriptsFixture
foreach ($name in @('desktop.ini', 'fi_old.ico', 'folder.ico')) {
    Set-Content -LiteralPath (Join-Path $folderFixture $name) -Encoding ASCII -Value $name
}

$harnessPath = Join-Path $fixtureRoot 'Invoke-RestoreFixture.ps1'
Set-Content -LiteralPath $harnessPath -Encoding UTF8 -Value @'
$ErrorActionPreference = 'Stop'
$answers = [Collections.Generic.Queue[string]]::new()
foreach ($answer in @('2', '1', '0')) { $answers.Enqueue($answer) }
function Read-Host { param([string]$Prompt); return $answers.Dequeue() }
function attrib { param([Parameter(ValueFromRemainingArguments = $true)]$Arguments) }
function Stop-Process { param([Parameter(ValueFromRemainingArguments = $true)]$Arguments) }
function Start-Process { param([Parameter(ValueFromRemainingArguments = $true)]$Arguments) }
function Start-Sleep { param([Parameter(ValueFromRemainingArguments = $true)]$Arguments) }
& $env:FOLDER_THUMB_MAIN_SCRIPT -RecycleAction {
    param([string]$Path)
    Move-Item -LiteralPath $Path -Destination $env:FOLDER_THUMB_RECYCLE_FIXTURE
}
'@

$previousMainScript = $env:FOLDER_THUMB_MAIN_SCRIPT
$previousRecycleFixture = $env:FOLDER_THUMB_RECYCLE_FIXTURE
$env:FOLDER_THUMB_MAIN_SCRIPT = Join-Path $scriptsFixture 'set_folder_thumb.ps1'
$env:FOLDER_THUMB_RECYCLE_FIXTURE = $restoreFixture
try {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $harnessPath | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Error ('Restore fixture exited with code {0}.' -f $LASTEXITCODE)
        exit 1
    }
}
finally {
    $env:FOLDER_THUMB_MAIN_SCRIPT = $previousMainScript
    $env:FOLDER_THUMB_RECYCLE_FIXTURE = $previousRecycleFixture
}

foreach ($name in @('desktop.ini', 'fi_old.ico', 'folder.ico')) {
    if (Test-Path -LiteralPath (Join-Path $folderFixture $name)) {
        Write-Error ('Restore mode kept source file: {0}' -f $name)
        exit 1
    }
    if (-not (Test-Path -LiteralPath (Join-Path $restoreFixture $name))) {
        Write-Error ('Restore mode did not route file through recycle action: {0}' -f $name)
        exit 1
    }
}

Write-Host 'File cleanup checks passed.'
