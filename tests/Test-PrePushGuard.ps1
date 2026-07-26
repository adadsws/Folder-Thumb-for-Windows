[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$hookPath = Join-Path $repositoryRoot '.githooks\pre-push'
$gitPath = (Get-Command git -ErrorAction Stop).Source
$gitRoot = Split-Path -Parent (Split-Path -Parent $gitPath)
$shellPath = Join-Path $gitRoot 'bin\sh.exe'

if (-not (Test-Path -LiteralPath $hookPath -PathType Leaf)) {
    Write-Error 'Pre-push hook is missing.'
    exit 1
}

if (-not (Test-Path -LiteralPath $shellPath -PathType Leaf)) {
    Write-Error 'Git for Windows shell is missing.'
    exit 1
}

function Invoke-PathCheck {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $output = & $shellPath $hookPath -CheckPath $Path 2>&1
    $exitCode = $LASTEXITCODE
    $ErrorActionPreference = $previousErrorActionPreference

    [pscustomobject]@{
        ExitCode = $exitCode
        Output = ($output -join [Environment]::NewLine)
    }
}

$sanitizedExampleResult = Invoke-PathCheck -Path 'secrets/service.env.example'
if ($sanitizedExampleResult.ExitCode -ne 0) {
    Write-Error 'Pre-push hook rejected a sanitized secrets .example file.'
    exit 1
}

$privateSecretResult = Invoke-PathCheck -Path 'secrets/service.env'
if ($privateSecretResult.ExitCode -eq 0) {
    Write-Error 'Pre-push hook allowed a private secrets file.'
    exit 1
}

if ($privateSecretResult.Output -notmatch 'secrets/service.env') {
    Write-Error 'Pre-push hook did not identify the blocked private secrets path.'
    exit 1
}

$archivedResult = Invoke-PathCheck -Path '~archived/old-file.txt'
if ($archivedResult.ExitCode -eq 0) {
    Write-Error 'Pre-push hook allowed an archived path.'
    exit 1
}

if ($archivedResult.Output -notmatch '~archived/old-file.txt') {
    Write-Error 'Pre-push hook did not identify the blocked archived path.'
    exit 1
}

$referenceResult = Invoke-PathCheck -Path '~ref/reference/file.txt'
if ($referenceResult.ExitCode -eq 0) {
    Write-Error 'Pre-push hook allowed a reference repository path.'
    exit 1
}

if ($referenceResult.Output -notmatch '~ref/reference/file.txt') {
    Write-Error 'Pre-push hook did not identify the blocked reference repository path.'
    exit 1
}

$activePlanResult = Invoke-PathCheck -Path 'docs/superpowers/plans/active.md'
if ($activePlanResult.ExitCode -eq 0) {
    Write-Error 'Pre-push hook allowed an active Superpowers plan.'
    exit 1
}

if ($activePlanResult.Output -notmatch 'docs/superpowers/plans/active.md') {
    Write-Error 'Pre-push hook did not identify the blocked active Superpowers plan.'
    exit 1
}

$hookScriptPath = Join-Path $repositoryRoot '.githooks\pre-push.ps1'
foreach ($blockedUnicodePath in @('~ref/中文.txt', 'secrets/密钥.env')) {
    $unicodeFixtureRoot = Join-Path (
        Join-Path $repositoryRoot '~temp'
    ) ('pre-push-unicode-{0}' -f [guid]::NewGuid().ToString('N'))
    $unicodeFilePath = Join-Path $unicodeFixtureRoot $blockedUnicodePath

    [void](New-Item -ItemType Directory -Path (Split-Path -Parent $unicodeFilePath) -Force)
    Set-Content -LiteralPath $unicodeFilePath -Encoding UTF8 -Value 'local-only value'

    & git init --quiet $unicodeFixtureRoot
    & git -C $unicodeFixtureRoot add $blockedUnicodePath
    & git -C $unicodeFixtureRoot -c user.name=Test -c user.email=test@example.invalid `
        commit --quiet -m 'Add Unicode local-only path'
    if ($LASTEXITCODE -ne 0) {
        Write-Error 'Unable to create the Unicode pre-push fixture commit.'
        exit 1
    }

    $unicodeCommit = git -C $unicodeFixtureRoot rev-parse HEAD
    $unicodeInput = 'refs/heads/main {0} refs/heads/main {1}' -f `
        $unicodeCommit, ('0' * 40)
    $previousLocation = Get-Location
    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    Set-Location -LiteralPath $unicodeFixtureRoot
    try {
        $unicodeOutput = $unicodeInput |
            & powershell -NoProfile -ExecutionPolicy Bypass -File $hookScriptPath `
                origin 'https://example.invalid/repository.git' 2>&1
        $unicodeExitCode = $LASTEXITCODE
    }
    finally {
        Set-Location -LiteralPath $previousLocation
        $ErrorActionPreference = $previousErrorActionPreference
    }

    if ($unicodeExitCode -eq 0) {
        Write-Error (
            'Pre-push hook allowed Unicode local-only path from Git history: {0}' -f
            $blockedUnicodePath
        )
        exit 1
    }

    $unicodeOutputText = $unicodeOutput -join [Environment]::NewLine
    $blockedPathPrefix = '{0}/' -f ($blockedUnicodePath -split '/')[0]
    if (
        $unicodeOutputText -notmatch 'Push blocked' -or
        $unicodeOutputText -notmatch [regex]::Escape($blockedPathPrefix)
    ) {
        Write-Error (
            'Pre-push hook failed without classifying Unicode path {0}: {1}' -f
            $blockedUnicodePath, $unicodeOutputText
        )
        exit 1
    }
}

$safeDirectory = $repositoryRoot.Replace('\', '/')
$originMain = git -c "safe.directory=$safeDirectory" rev-parse origin/main

if ($LASTEXITCODE -ne 0) {
    Write-Error 'Unable to resolve HEAD and origin/main.'
    exit 1
}

$safeInput = 'refs/heads/main {0} refs/heads/main {0}' -f $originMain
$safeOutput = $safeInput |
    & $shellPath $hookPath origin 'https://example.invalid/repository.git' 2>&1
$safeExitCode = $LASTEXITCODE

if ($safeExitCode -ne 0) {
    Write-Error (
        'Pre-push hook rejected a range with no outgoing commits: {0}' -f
        ($safeOutput -join [Environment]::NewLine)
    )
    exit 1
}

Write-Host 'Pre-push guard checks passed.'
