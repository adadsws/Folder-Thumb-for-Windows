function Move-FileToRecycleBin {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath,

        [scriptblock]$RecycleAction
    )

    if (-not (Test-Path -LiteralPath $LiteralPath -PathType Leaf)) {
        return $true
    }

    if ($null -eq $RecycleAction) {
        Add-Type -AssemblyName Microsoft.VisualBasic
        $RecycleAction = {
            param([string]$Path)

            [Microsoft.VisualBasic.FileIO.FileSystem]::DeleteFile(
                $Path,
                [Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs,
                [Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin
            )
        }
    }

    try {
        $null = & $RecycleAction $LiteralPath
        return $true
    }
    catch {
        Write-Warning (
            'Unable to move file to the Recycle Bin; file was kept: {0}. {1}' -f
            $LiteralPath, $_.Exception.Message
        )
        return $false
    }
}

function Remove-FilePermanently {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$LiteralPath
    )

    if (Test-Path -LiteralPath $LiteralPath -PathType Leaf) {
        Remove-Item -LiteralPath $LiteralPath -Force
    }
}
