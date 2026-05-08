[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$StagingRoot = "dist\CSV_Tool",

    [Parameter(Mandatory = $false)]
    [string]$ManifestName = "hardlink_manifest.csv"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Resolve-AbsolutePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }
    $scriptDir = Split-Path -Parent $PSCommandPath
    $repoRoot = [System.IO.Path]::GetFullPath((Join-Path $scriptDir "..\.."))
    return [System.IO.Path]::GetFullPath((Join-Path $repoRoot $Path))
}

function Invoke-MklinkHardlink {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LinkPath,
        [Parameter(Mandatory = $true)]
        [string]$TargetPath
    )
    $quotedLink = '"' + $LinkPath + '"'
    $quotedTarget = '"' + $TargetPath + '"'
    cmd /c "mklink /H $quotedLink $quotedTarget" > $null
    if ($LASTEXITCODE -ne 0) {
        throw "mklink failed: '$LinkPath' -> '$TargetPath'"
    }
}

function Get-Sha256HexOfFile {
    param([Parameter(Mandatory = $true)][string]$LiteralPath)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    $fs = $null
    try {
        $fs = [System.IO.File]::OpenRead($LiteralPath)
        $bytes = $sha.ComputeHash($fs)
    }
    finally {
        if ($null -ne $fs) { $fs.Dispose() }
        $sha.Dispose()
    }
    return ([System.BitConverter]::ToString($bytes)).Replace("-", "")
}

function Get-FileId {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )
    $quotedPath = '"' + $Path + '"'
    $result = cmd /c "fsutil file queryfileid $quotedPath" 2>&1
    if ($LASTEXITCODE -ne 0) {
        throw "fsutil file queryfileid failed for '$Path': $result"
    }
    if ($result -match "0x[0-9A-Fa-f]+") {
        return $Matches[0].ToUpperInvariant()
    }
    throw "Could not parse file id for '$Path': $result"
}

$stagingAbs = Resolve-AbsolutePath -Path $StagingRoot
$appRoot = Join-Path $stagingAbs "app"
$manifestPath = Join-Path $stagingAbs $ManifestName

if (-not (Test-Path -LiteralPath $stagingAbs -PathType Container)) {
    throw "Staging root not found: $stagingAbs"
}
if (-not (Test-Path -LiteralPath $appRoot -PathType Container)) {
    throw "Target app directory not found: $appRoot"
}

Write-Host "[dedupe] Staging root: $stagingAbs"
Write-Host "[dedupe] Target root : $appRoot"

$sharedDir = Join-Path $appRoot "shared"
$sharedPrefix = if ($sharedDir.EndsWith("\")) { $sharedDir } else { $sharedDir + "\" }
$files = Get-ChildItem -LiteralPath $appRoot -Recurse -File | Where-Object {
    $full = $_.FullName
    if ($full.StartsWith($sharedPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $false
    }
    if ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
        return $false
    }
    return $true
}
if ($files.Count -eq 0) {
    Write-Host "[dedupe] No files found under app/. Nothing to do."
    exit 0
}

Write-Host "[dedupe] Hashing $($files.Count) files..."
$rows = foreach ($file in $files) {
    $hash = Get-Sha256HexOfFile -LiteralPath $file.FullName
    [PSCustomObject]@{
        Path   = $file.FullName
        Length = [Int64]$file.Length
        Hash   = $hash
    }
}

$dupGroups = $rows | Group-Object -Property Hash | Where-Object { $_.Count -gt 1 }
if ($dupGroups.Count -eq 0) {
    Write-Host "[dedupe] No duplicate payloads found."
    exit 0
}

$theoreticalSavings = ($dupGroups | ForEach-Object {
    ($_.Count - 1) * [Int64]$_.Group[0].Length
} | Measure-Object -Sum).Sum
if ($null -eq $theoreticalSavings) {
    $theoreticalSavings = 0
}

$actions = New-Object System.Collections.Generic.List[object]
$appliedBytes = 0L
$alreadyLinkedCount = 0

Write-Host ("[dedupe] Duplicate groups: {0}, theoretical savings: {1:N1} MB" -f $dupGroups.Count, ($theoreticalSavings / 1MB))

foreach ($group in $dupGroups) {
    $sorted = $group.Group | Sort-Object -Property Path
    $canonical = $sorted[0]
    $canonicalPath = $canonical.Path
    $canonicalId = Get-FileId -Path $canonicalPath

    for ($i = 1; $i -lt $sorted.Count; $i++) {
        $item = $sorted[$i]
        $linkPath = $item.Path

        if (-not (Test-Path -LiteralPath $linkPath -PathType Leaf)) {
            continue
        }

        $linkId = Get-FileId -Path $linkPath
        if ($linkId -eq $canonicalId) {
            $alreadyLinkedCount++
            continue
        }

        $tempPath = "$linkPath.__dedupe_tmp__"
        if (Test-Path -LiteralPath $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force
        }

        Move-Item -LiteralPath $linkPath -Destination $tempPath
        try {
            Invoke-MklinkHardlink -LinkPath $linkPath -TargetPath $canonicalPath
            Remove-Item -LiteralPath $tempPath -Force

            $actions.Add([PSCustomObject]@{
                Timestamp = (Get-Date).ToString("s")
                Canonical = $canonicalPath
                Linked    = $linkPath
                Length    = [Int64]$item.Length
                Hash      = $item.Hash
            }) | Out-Null
            $appliedBytes += [Int64]$item.Length
        }
        catch {
            if (Test-Path -LiteralPath $linkPath) {
                Remove-Item -LiteralPath $linkPath -Force
            }
            Move-Item -LiteralPath $tempPath -Destination $linkPath
            throw
        }
    }
}

if ($actions.Count -gt 0) {
    $actions | Export-Csv -LiteralPath $manifestPath -NoTypeInformation -Encoding UTF8
    Write-Host ("[dedupe] Linked files : {0}" -f $actions.Count)
    Write-Host ("[dedupe] Saved bytes  : {0:N1} MB" -f ($appliedBytes / 1MB))
    Write-Host ("[dedupe] Manifest     : {0}" -f $manifestPath)
    Write-Host ("[dedupe] Rollback cmd : powershell -NoProfile -ExecutionPolicy Bypass -File ""{0}"" -ManifestPath ""{1}""" -f (Join-Path (Split-Path -Parent $PSCommandPath) "rollback_staging_hardlinks.ps1"), $manifestPath)
}
else {
    Write-Host "[dedupe] Nothing newly linked."
}

if ($alreadyLinkedCount -gt 0) {
    Write-Host ("[dedupe] Already linked duplicates skipped: {0}" -f $alreadyLinkedCount)
}

exit 0
