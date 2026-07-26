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
