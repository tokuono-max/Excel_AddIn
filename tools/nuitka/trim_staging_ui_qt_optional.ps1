# -----------------------------------------------------------------------------
# trim_staging_ui_qt_optional.ps1
# After Nuitka merge into app\bin, optional Qt image codec DLLs (Windows).
# Safe for a widgets-only CSV tool: no WebP/TIFF/ICNS/TGA/WBMP in typical UI.
# Also removes Qt SQL driver plugins (ui_qt does not use QSql).
# Restore files by rebuilding; if a rare format is needed, remove from $TrimImageFormats.
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

$TrimImageFormats = @(
    "qwebp.dll",
    "qtiff.dll",
    "qicns.dll",
    "qtga.dll",
    "qwbmp.dll"
)

$stagingAbs = Resolve-RepoAbsolutePath -Path $StagingRoot
$imgDir = Join-Path $stagingAbs "app\bin\PySide6\qt-plugins\imageformats"

if (-not (Test-Path -LiteralPath $imgDir -PathType Container)) {
    Write-Host "[trim_ui_qt] skip: not found: $imgDir"
    exit 0
}

$removed = 0L
foreach ($name in $TrimImageFormats) {
    $p = Join-Path $imgDir $name
    if (Test-Path -LiteralPath $p -PathType Leaf) {
        $len = (Get-Item -LiteralPath $p).Length
        Remove-Item -LiteralPath $p -Force
        $removed += $len
        Write-Host "[trim_ui_qt] removed: $p ($([math]::Round($len / 1KB, 1)) KB)"
    }
}

if ($removed -gt 0) {
    Write-Host "[trim_ui_qt] total removed: $([math]::Round($removed / 1MB, 3)) MB"
} else {
    Write-Host "[trim_ui_qt] no matching optional image plugins under $imgDir"
}

# Qt SQL drivers (QOCI/QPSQL/...) — not used by this UI stack
$sqlDir = Join-Path $stagingAbs "app\bin\PySide6\qt-plugins\sqldrivers"
if (Test-Path -LiteralPath $sqlDir -PathType Container) {
    $sqlFiles = @(Get-ChildItem -LiteralPath $sqlDir -File -ErrorAction SilentlyContinue)
    $sqlBytes = 0L
    foreach ($f in $sqlFiles) {
        $sqlBytes += $f.Length
        Remove-Item -LiteralPath $f.FullName -Force
    }
    if ($sqlBytes -gt 0) {
        Write-Host "[trim_ui_qt] removed sqldrivers: $([math]::Round($sqlBytes / 1KB, 1)) KB ($($sqlFiles.Count) files)"
    }
}

exit 0
