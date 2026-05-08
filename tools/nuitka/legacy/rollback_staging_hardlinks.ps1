[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$ManifestPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) {
    throw "Manifest not found: $ManifestPath"
}

$rows = Import-Csv -LiteralPath $ManifestPath
if ($rows.Count -eq 0) {
    Write-Host "[rollback] Manifest is empty. Nothing to do."
    exit 0
}

$restored = 0
$skipped = 0

foreach ($row in $rows) {
    $path = $row.Linked
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
        $skipped++
        continue
    }

    $tempPath = "$path.__rollback_tmp__"
    if (Test-Path -LiteralPath $tempPath) {
        Remove-Item -LiteralPath $tempPath -Force
    }

    Copy-Item -LiteralPath $path -Destination $tempPath -Force
    Remove-Item -LiteralPath $path -Force
    Move-Item -LiteralPath $tempPath -Destination $path
    $restored++
}

Write-Host ("[rollback] Restored files: {0}" -f $restored)
if ($skipped -gt 0) {
    Write-Host ("[rollback] Skipped missing paths: {0}" -f $skipped)
}

exit 0
