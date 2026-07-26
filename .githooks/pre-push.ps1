[CmdletBinding()]
param(
    [string]$RemoteName,
    [string]$RemoteUrl,
    [string[]]$CheckPath
)

$ErrorActionPreference = 'Stop'
$repositoryRoot = Split-Path -Parent $PSScriptRoot
$safeDirectory = $repositoryRoot.Replace('\', '/')
$zeroObjectId = '0000000000000000000000000000000000000000'
$blockedPaths = [System.Collections.Generic.HashSet[string]]::new(
    [System.StringComparer]::OrdinalIgnoreCase
)

function Test-IsLocalOnlyPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $normalizedPath = $Path.Replace('\', '/')

    if ($normalizedPath -match '^(~archived|~ref|docs/superpowers/plans)(/|$)') {
        return $true
    }

    if ($normalizedPath -match '^secrets(/|$)') {
        return -not $normalizedPath.EndsWith(
            '.example',
            [System.StringComparison]::OrdinalIgnoreCase
        )
    }

    return $false
}

foreach ($path in @($CheckPath)) {
    if ([string]::IsNullOrWhiteSpace($path)) {
        continue
    }

    if (Test-IsLocalOnlyPath -Path $path) {
        [void]$blockedPaths.Add($path.Replace('\', '/'))
    }
}

$stdinText = [Console]::In.ReadToEnd()
$updates = @($stdinText -split "\r?\n" | Where-Object { $_.Trim() })

foreach ($update in $updates) {
    $fields = @($update -split '\s+')
    if ($fields.Count -ne 4) {
        Write-Error ('Invalid pre-push input: {0}' -f $update)
        exit 1
    }

    $localObjectId = $fields[1]
    $remoteObjectId = $fields[3]

    if ($localObjectId -eq $zeroObjectId) {
        continue
    }

    if ($remoteObjectId -eq $zeroObjectId) {
        $commits = @(
            git -c "safe.directory=$safeDirectory" rev-list $localObjectId --not "--remotes=$RemoteName"
        )
    }
    else {
        $commits = @(
            git -c "safe.directory=$safeDirectory" rev-list "$remoteObjectId..$localObjectId"
        )
    }

    if ($LASTEXITCODE -ne 0) {
        Write-Error 'Unable to inspect outgoing commit history.'
        exit 1
    }

    foreach ($commit in $commits) {
        $paths = @(
            git -c "safe.directory=$safeDirectory" -c core.quotepath=false `
                diff-tree --root --no-commit-id --name-only -r $commit
        )

        if ($LASTEXITCODE -ne 0) {
            Write-Error ('Unable to inspect outgoing commit: {0}' -f $commit)
            exit 1
        }

        foreach ($path in $paths) {
            $normalizedPath = $path.Replace('\', '/')
            if (Test-IsLocalOnlyPath -Path $normalizedPath) {
                [void]$blockedPaths.Add($normalizedPath)
            }
        }
    }
}

if ($blockedPaths.Count -gt 0) {
    [Console]::Error.WriteLine('Push blocked: outgoing history contains local-only paths:')
    foreach ($blockedPath in @($blockedPaths | Sort-Object)) {
        [Console]::Error.WriteLine('  - {0}' -f $blockedPath)
    }
    [Console]::Error.WriteLine(
        'Publish from output/github-export/ using the sanitized main-only export workflow.'
    )
    exit 1
}

exit 0
