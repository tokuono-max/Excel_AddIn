# -----------------------------------------------------------------------------
# measure_csv_tool_size.ps1
# Estimates disk footprint of dist\CSV_Tool (or install root) for the 500 MB goal.
# Walks app\ once: does not recurse into directory junctions/symlinks (if any).
# Skips symlink files at reparse points. Layout: app\bin holds merged Nuitka outputs.
# Does NOT measure Inno Setup .exe size (that is a separate metric).
# Usage: powershell -File tools\nuitka\measure_csv_tool_size.ps1 [-StagingRoot dist\CSV_Tool]
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

function Get-TreeBytesSkipRecurseIntoReparseDirs {
    param([Parameter(Mandatory = $true)][string]$RootDir)
    if (-not (Test-Path -LiteralPath $RootDir -PathType Container)) {
        return [int64]0
    }
    [int64]$sum = 0
    $stack = New-Object System.Collections.Stack
    $stack.Push($RootDir)
    while ($stack.Count -gt 0) {
        $d = $stack.Pop()
        Get-ChildItem -LiteralPath $d -Force -ErrorAction SilentlyContinue | ForEach-Object {
            if ($_.PSIsContainer) {
                if ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
                    return
                }
                $stack.Push($_.FullName)
            }
            else {
                if (-not ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint)) {
                    $sum += [int64]$_.Length
                }
            }
        }
    }
    return $sum
}

$stagingAbs = Resolve-RepoAbsolutePath -Path $StagingRoot
$appRoot = Join-Path $stagingAbs "app"
$configDir = Join-Path $stagingAbs "config"
$xlwingsConf = Join-Path $stagingAbs "xlwings.conf"

if (-not (Test-Path -LiteralPath $stagingAbs -PathType Container)) {
    throw "Staging not found: $stagingAbs"
}

[int64]$appBytes = 0
if (Test-Path -LiteralPath $appRoot -PathType Container) {
    $appBytes = Get-TreeBytesSkipRecurseIntoReparseDirs -RootDir $appRoot
}

[int64]$configFiles = 0
if (Test-Path -LiteralPath $configDir -PathType Container) {
    Get-ChildItem -LiteralPath $configDir -File -ErrorAction SilentlyContinue | ForEach-Object {
        $configFiles += [int64]$_.Length
    }
}
[int64]$xw = 0
if (Test-Path -LiteralPath $xlwingsConf -PathType Leaf) {
    $xw = [int64](Get-Item -LiteralPath $xlwingsConf).Length
}

$total = $appBytes + $configFiles + $xw

Write-Host "[measure_csv_tool] Goal metric: CSV_Tool folder on disk (not the Inno installer .exe size)."
Write-Host "[measure_csv_tool] Staging: $stagingAbs"
Write-Host ("  app\\ (single walk, junction dirs not traversed twice)  {0,10:N1} MB" -f ($appBytes / 1MB))
Write-Host ("  config\\ + xlwings.conf                              {0,10:N1} MB" -f (($configFiles + $xw) / 1MB))
Write-Host ("  --- estimated total                                   {0,10:N1} MB" -f ($total / 1MB))
Write-Host "[measure_csv_tool] Cross-check: Explorer -> folder properties -> Size on disk for this path."
exit 0
