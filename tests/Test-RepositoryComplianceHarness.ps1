[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$complianceScript = Join-Path $PSScriptRoot 'Test-RepositoryCompliance.ps1'
$fixtureRoot = Join-Path (
    Join-Path $repositoryRoot '~temp'
) ('repository-compliance-{0}' -f [guid]::NewGuid().ToString('N'))

function Invoke-ComplianceCheck {
    param(
        [string]$Root
    )

    $previousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    if ([string]::IsNullOrWhiteSpace($Root)) {
        $output = & powershell -NoProfile -ExecutionPolicy Bypass `
            -File $complianceScript 2>&1
    }
    else {
        $output = & powershell -NoProfile -ExecutionPolicy Bypass `
            -File $complianceScript -RepositoryRoot $Root 2>&1
    }
    $exitCode = $LASTEXITCODE
    $ErrorActionPreference = $previousErrorActionPreference

    [pscustomobject]@{
        ExitCode = $exitCode
        Output = ($output -join [Environment]::NewLine)
    }
}

$directResult = Invoke-ComplianceCheck
if ($directResult.ExitCode -ne 0) {
    Write-Error (
        'Compliance check failed with its default repository root: {0}' -f
        $directResult.Output
    )
    exit 1
}

[void](New-Item -ItemType Directory -Path $fixtureRoot -Force)
[void](New-Item -ItemType Directory -Path (Join-Path $fixtureRoot 'output') -Force)
[void](New-Item -ItemType Directory -Path (Join-Path $fixtureRoot '~temp') -Force)
[void](New-Item -ItemType Directory -Path (Join-Path $fixtureRoot '~outputs') -Force)
[void](New-Item -ItemType Directory -Path (Join-Path $fixtureRoot '.worktrees') -Force)
[void](New-Item -ItemType Directory -Path (Join-Path $fixtureRoot '.venv') -Force)

Set-Content -LiteralPath (Join-Path $fixtureRoot '.gitignore') -Encoding UTF8 -Value @(
    '~outputs/'
    '~temp/'
    '.worktrees/'
    '.venv/'
)
Set-Content -LiteralPath (Join-Path $fixtureRoot 'README.md') -Encoding UTF8 -Value '# Test'
Set-Content -LiteralPath (Join-Path $fixtureRoot 'CHANGELOG.md') -Encoding UTF8 -Value '# Test'
Set-Content -LiteralPath (Join-Path $fixtureRoot 'AGENT_CONTEXT.md') -Encoding UTF8 -Value '# Test'

& git init --quiet $fixtureRoot
if ($LASTEXITCODE -ne 0) {
    Write-Error 'Unable to initialize the compliance test repository.'
    exit 1
}

& git -C $fixtureRoot add .gitignore README.md CHANGELOG.md AGENT_CONTEXT.md
if ($LASTEXITCODE -ne 0) {
    Write-Error 'Unable to stage the compliance test baseline.'
    exit 1
}

Set-Content -LiteralPath (Join-Path $fixtureRoot '~outputs\中文.txt') `
    -Encoding UTF8 -Value 'allowed output'
Set-Content -LiteralPath (Join-Path $fixtureRoot '~temp\中文.txt') `
    -Encoding UTF8 -Value 'allowed temporary file'
Set-Content -LiteralPath (Join-Path $fixtureRoot '.worktrees\中文.txt') `
    -Encoding UTF8 -Value 'allowed worktree file'
Set-Content -LiteralPath (Join-Path $fixtureRoot '.venv\中文.txt') `
    -Encoding UTF8 -Value 'allowed virtual environment file'

$baselineResult = Invoke-ComplianceCheck -Root $fixtureRoot
if ($baselineResult.ExitCode -ne 0) {
    Write-Error (
        'Compliance check rejected a tracked baseline with Unicode output paths: {0}' -f
        $baselineResult.Output
    )
    exit 1
}

Set-Content -LiteralPath (Join-Path $fixtureRoot 'private-local.txt') `
    -Encoding UTF8 -Value 'must be tracked'
Set-Content -LiteralPath (Join-Path $fixtureRoot '.git\info\exclude') -Encoding UTF8 -Value @(
    'private-local.txt'
)

$infoExcludeResult = Invoke-ComplianceCheck -Root $fixtureRoot
if ($infoExcludeResult.ExitCode -eq 0) {
    Write-Error 'Compliance check allowed a non-output file hidden by .git/info/exclude.'
    exit 1
}

$compactInfoExcludeOutput = $infoExcludeResult.Output -replace '\s', ''
if (
    $compactInfoExcludeOutput -notmatch 'hiddenbyaGitignore' -or
    $compactInfoExcludeOutput -notmatch 'private-local\.txt'
) {
    Write-Error (
        'Compliance check did not identify the .git/info/exclude file: {0}' -f
        $infoExcludeResult.Output
    )
    exit 1
}

& git -C $fixtureRoot add --force private-local.txt
if ($LASTEXITCODE -ne 0) {
    Write-Error 'Unable to stage the .git/info/exclude fixture.'
    exit 1
}

$globalIgnorePath = Join-Path $fixtureRoot '~temp\global-ignore'
Set-Content -LiteralPath $globalIgnorePath -Encoding UTF8 -Value 'global-hidden.txt'
& git -C $fixtureRoot config core.excludesFile $globalIgnorePath
if ($LASTEXITCODE -ne 0) {
    Write-Error 'Unable to configure the core.excludesFile fixture.'
    exit 1
}

Set-Content -LiteralPath (Join-Path $fixtureRoot 'global-hidden.txt') `
    -Encoding UTF8 -Value 'must be tracked'

$globalExcludeResult = Invoke-ComplianceCheck -Root $fixtureRoot
if ($globalExcludeResult.ExitCode -eq 0) {
    Write-Error 'Compliance check allowed a non-output file hidden by core.excludesFile.'
    exit 1
}

$compactGlobalExcludeOutput = $globalExcludeResult.Output -replace '\s', ''
if (
    $compactGlobalExcludeOutput -notmatch 'hiddenbyaGitignore' -or
    $compactGlobalExcludeOutput -notmatch 'global-hidden\.txt'
) {
    Write-Error (
        'Compliance check did not identify the core.excludesFile file: {0}' -f
        $globalExcludeResult.Output
    )
    exit 1
}

& git -C $fixtureRoot add --force global-hidden.txt
if ($LASTEXITCODE -ne 0) {
    Write-Error 'Unable to stage the core.excludesFile fixture.'
    exit 1
}

Set-Content -LiteralPath (Join-Path $fixtureRoot 'ordinary-untracked.txt') `
    -Encoding UTF8 -Value 'must be tracked'
$untrackedResult = Invoke-ComplianceCheck -Root $fixtureRoot
if ($untrackedResult.ExitCode -eq 0) {
    Write-Error 'Compliance check allowed an ordinary untracked file.'
    exit 1
}

$compactUntrackedOutput = $untrackedResult.Output -replace '\s', ''
if (
    $compactUntrackedOutput -notmatch 'mustbetrackedbylocalGit' -or
    $compactUntrackedOutput -notmatch 'ordinary-untracked\.txt'
) {
    Write-Error (
        'Compliance check did not identify the ordinary untracked file: {0}' -f
        $untrackedResult.Output
    )
    exit 1
}

Write-Host 'Repository compliance harness checks passed.'
