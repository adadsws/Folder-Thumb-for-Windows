[CmdletBinding()]
param(
    [string]$RepositoryRoot
)

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($RepositoryRoot)) {
    $RepositoryRoot = Split-Path -Parent $PSScriptRoot
}

$repositoryRoot = (Resolve-Path -LiteralPath $RepositoryRoot).Path
$projectRoot = (Resolve-Path -LiteralPath (Split-Path -Parent $PSScriptRoot)).Path
$isProjectRepository = $repositoryRoot -eq $projectRoot
$safeDirectory = $repositoryRoot.Replace('\', '/')
$failures = [System.Collections.Generic.List[string]]::new()

function Add-ComplianceFailure {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $failures.Add($Message)
}

$gitIgnorePath = Join-Path $repositoryRoot '.gitignore'
$requiredIgnoreRules = @('~outputs/', '~temp/', '.worktrees/', '.venv/')

if (-not (Test-Path -LiteralPath $gitIgnorePath -PathType Leaf)) {
    Add-ComplianceFailure '.gitignore does not exist.'
}
else {
    $actualIgnoreRules = @(
        Get-Content -LiteralPath $gitIgnorePath -Encoding UTF8 |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -and -not $_.StartsWith('#') }
    )

    $missingIgnoreRules = @($requiredIgnoreRules | Where-Object { $_ -notin $actualIgnoreRules })
    $unexpectedIgnoreRules = @($actualIgnoreRules | Where-Object { $_ -notin $requiredIgnoreRules })

    if ($missingIgnoreRules.Count -gt 0) {
        Add-ComplianceFailure (
            '.gitignore is missing required rules: {0}' -f ($missingIgnoreRules -join ', ')
        )
    }

    if ($unexpectedIgnoreRules.Count -gt 0) {
        Add-ComplianceFailure (
            '.gitignore contains unexpected rules: {0}' -f ($unexpectedIgnoreRules -join ', ')
        )
    }
}

foreach ($requiredDocument in @('README.md', 'CHANGELOG.md', 'AGENT_CONTEXT.md')) {
    $requiredDocumentPath = Join-Path $repositoryRoot $requiredDocument
    if (-not (Test-Path -LiteralPath $requiredDocumentPath -PathType Leaf)) {
        Add-ComplianceFailure ("Required document is missing: {0}" -f $requiredDocument)
    }
}

$ignoredPaths = @(
    git -c "safe.directory=$safeDirectory" -c core.quotepath=false -C $repositoryRoot `
        ls-files --others --ignored --exclude-standard
)
if ($LASTEXITCODE -ne 0) {
    Add-ComplianceFailure 'Unable to inspect ignored repository files.'
}
else {
    foreach ($ignoredPath in $ignoredPaths) {
        $normalizedPath = $ignoredPath.Replace('\', '/')
        if ($normalizedPath -notmatch '^(~outputs|~temp|\.worktrees|\.venv)(/|$)') {
            Add-ComplianceFailure (
                'Repository file is hidden by a Git ignore source: {0}' -f $normalizedPath
            )
        }
    }
}

if ($isProjectRepository) {
    foreach ($requiredProjectDocument in @('AGENTS.md', 'reference/README.md', '.gitmodules')) {
        $requiredProjectDocumentPath = Join-Path $repositoryRoot $requiredProjectDocument
        if (-not (Test-Path -LiteralPath $requiredProjectDocumentPath -PathType Leaf)) {
            Add-ComplianceFailure (
                'Required project document is missing: {0}' -f $requiredProjectDocument
            )
        }
    }

    foreach ($legacyPath in @('~archived', '~ref', 'output', 'set_folder_thumb.ps1')) {
        if (Test-Path -LiteralPath (Join-Path $repositoryRoot $legacyPath)) {
            Add-ComplianceFailure ('Legacy project path must be migrated: {0}' -f $legacyPath)
        }
    }

    foreach ($requiredProjectPath in @(
        'scripts/set_folder_thumb.ps1',
        'docs/finished_plans',
        'reference'
    )) {
        if (-not (Test-Path -LiteralPath (Join-Path $repositoryRoot $requiredProjectPath))) {
            Add-ComplianceFailure ('Required project path is missing: {0}' -f $requiredProjectPath)
        }
    }

    $prePushStage = git -c "safe.directory=$safeDirectory" -C $repositoryRoot `
        ls-files --stage -- .githooks/pre-push
    if ($prePushStage -notlike '100755 *') {
        Add-ComplianceFailure '.githooks/pre-push must be tracked as executable mode 100755.'
    }

    $expectedReferences = @(
        [pscustomobject]@{
            Path = 'reference/PowerToys'
            Url = 'https://github.com/microsoft/PowerToys.git'
            Commit = 'fc680d350f74f1f4eec8a64296420e614724e74d'
        },
        [pscustomobject]@{
            Path = 'reference/mattpocock-skills'
            Url = 'https://github.com/mattpocock/skills.git'
            Commit = 'ed37663cc5fbef691ddfecd080dff42f7e7e350d'
        },
        [pscustomobject]@{
            Path = 'reference/winutil'
            Url = 'https://github.com/ChrisTitusTech/winutil.git'
            Commit = '50be7390e586f664df76d7fed41fc3c39252288c'
        }
    )

    $gitModulesPath = Join-Path $repositoryRoot '.gitmodules'
    if (Test-Path -LiteralPath $gitModulesPath -PathType Leaf) {
        foreach ($reference in $expectedReferences) {
            $moduleName = $reference.Path.Substring('reference/'.Length)
            $configuredPath = git -c "safe.directory=$safeDirectory" -C $repositoryRoot `
                config -f .gitmodules --get "submodule.$moduleName.path"
            $configuredUrl = git -c "safe.directory=$safeDirectory" -C $repositoryRoot `
                config -f .gitmodules --get "submodule.$moduleName.url"

            if ($configuredPath -ne $reference.Path -or $configuredUrl -ne $reference.Url) {
                Add-ComplianceFailure (
                    'Submodule metadata is invalid for {0}.' -f $reference.Path
                )
            }

            $stageEntry = git -c "safe.directory=$safeDirectory" -C $repositoryRoot `
                ls-files --stage -- $reference.Path
            $expectedStagePrefix = '160000 {0} 0' -f $reference.Commit
            if ($stageEntry -notlike "$expectedStagePrefix*") {
                Add-ComplianceFailure (
                    'Submodule lock is invalid for {0}; expected {1}.' -f
                    $reference.Path, $reference.Commit
                )
            }
        }
    }
}

$untrackedPaths = @(
    git -c "safe.directory=$safeDirectory" -c core.quotepath=false -C $repositoryRoot `
        ls-files --others --exclude-standard
)
if ($LASTEXITCODE -ne 0) {
    Add-ComplianceFailure 'Unable to inspect untracked repository files.'
}
else {
    foreach ($untrackedPath in $untrackedPaths) {
        Add-ComplianceFailure (
            'Repository file must be tracked by local Git: {0}' -f
            $untrackedPath.Replace('\', '/')
        )
    }
}

$secretsPath = Join-Path $repositoryRoot 'secrets'
if (Test-Path -LiteralPath $secretsPath -PathType Container) {
    $privateFiles = Get-ChildItem -LiteralPath $secretsPath -File -Recurse |
        Where-Object { -not $_.Name.EndsWith('.example', [System.StringComparison]::OrdinalIgnoreCase) }

    foreach ($privateFile in $privateFiles) {
        $examplePath = '{0}.example' -f $privateFile.FullName
        if (-not (Test-Path -LiteralPath $examplePath -PathType Leaf)) {
            $relativePrivatePath = $privateFile.FullName.Substring($repositoryRoot.Length + 1)
            Add-ComplianceFailure (
                'Private file is missing a sanitized example: {0}.example' -f $relativePrivatePath
            )
        }
    }
}

if ($failures.Count -gt 0) {
    foreach ($failure in $failures) {
        Write-Error $failure -ErrorAction Continue
    }

    exit 1
}

Write-Host 'Repository compliance checks passed.'
