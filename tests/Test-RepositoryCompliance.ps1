[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$failures = [System.Collections.Generic.List[string]]::new()

function Add-ComplianceFailure {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message
    )

    $failures.Add($Message)
}

$gitIgnorePath = Join-Path $repositoryRoot '.gitignore'
$requiredIgnoreRules = @('output/', '~temp/')

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

foreach ($requiredDocument in @('README.md', 'CHANGELOG.md')) {
    $requiredDocumentPath = Join-Path $repositoryRoot $requiredDocument
    if (-not (Test-Path -LiteralPath $requiredDocumentPath -PathType Leaf)) {
        Add-ComplianceFailure ("Required document is missing: {0}" -f $requiredDocument)
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
