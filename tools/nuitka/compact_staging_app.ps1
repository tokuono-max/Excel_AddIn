# -----------------------------------------------------------------------------
# compact_staging_app.ps1
# Apply NTFS compression to staging app\ (reduces "size on disk" for DLL/PYD-heavy trees).
# Safe for junctions: compresses underlying clusters once.
# Skip with: set SKIP_COMPACT_STAGING=1 before build_nuitka_all.bat
# -----------------------------------------------------------------------------

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$StagingRoot = "dist\CSV_Tool"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-RepoAbsolutePath {
    param([Parameter(Mandatory = $true)][string]$Path)
    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }
    $scriptDir = Split-Path -Parent $PSCommandPath
    $repoRoot = [System.IO.Path]::GetFullPath((Join-Path $scriptDir "..\.."))
    return [System.IO.Path]::GetFullPath((Join-Path $repoRoot $Path))
}

$stagingAbs = Resolve-RepoAbsolutePath -Path $StagingRoot
$appRoot = Join-Path $stagingAbs "app"
if (-not (Test-Path -LiteralPath $appRoot -PathType Container)) {
    throw "compact_staging: app not found: $appRoot"
}

$compact = Join-Path $env:SystemRoot "System32\compact.exe"
if (-not (Test-Path -LiteralPath $compact -PathType Leaf)) {
    throw "compact.exe not found: $compact"
}

Write-Host "[compact_staging] NTFS compress (recursive): $appRoot"
$p = Start-Process -FilePath $compact -ArgumentList @("/C", "/S", $appRoot) -Wait -NoNewWindow -PassThru
if ($p.ExitCode -ne 0) {
    throw "compact.exe failed with exit $($p.ExitCode)"
}
Write-Host "[compact_staging] Done."
exit 0
