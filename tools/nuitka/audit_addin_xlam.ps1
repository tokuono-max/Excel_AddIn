# Log SHA256 / size / mtime for *.xlam under repo addin vs staging addin (UTF-8 append).
param(
    [Parameter(Mandatory = $true)]
    [string]$LogPath,
    [Parameter(Mandatory = $true)]
    [string]$RepoAddin,
    [Parameter(Mandatory = $true)]
    [string]$StagingAddin
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Add-L([string]$s) {
    Add-Content -LiteralPath $LogPath -Encoding utf8 -Value $s
}

Add-L "---- addin *.xlam audit ----"
if (-not (Test-Path -LiteralPath $RepoAddin -PathType Container)) {
    Add-L "[staging] audit: repo addin dir missing: $RepoAddin"
    Add-L "---- end addin *.xlam audit ----"
    exit 0
}

$repoFiles = @(Get-ChildItem -LiteralPath $RepoAddin -Filter *.xlam -File -ErrorAction SilentlyContinue)
if ($repoFiles.Count -eq 0) {
    Add-L "[staging] audit: no *.xlam under repo addin: $RepoAddin"
} else {
    foreach ($f in $repoFiles) {
        $h = Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256
        Add-L ("repo    {0} sha256={1} length={2} mtime_utc={3:o}" -f $f.Name, $h.Hash, $f.Length, $f.LastWriteTimeUtc)
    }
}

if (-not (Test-Path -LiteralPath $StagingAddin -PathType Container)) {
    Add-L "[staging] audit: staging addin dir missing: $StagingAddin"
    Add-L "---- end addin *.xlam audit ----"
    exit 0
}

$stFiles = @(Get-ChildItem -LiteralPath $StagingAddin -Filter *.xlam -File -ErrorAction SilentlyContinue)
if ($stFiles.Count -eq 0) {
    Add-L "[staging] audit: no *.xlam under staging addin: $StagingAddin"
} else {
    foreach ($f in $stFiles) {
        $h = Get-FileHash -LiteralPath $f.FullName -Algorithm SHA256
        Add-L ("staging {0} sha256={1} length={2} mtime_utc={3:o}" -f $f.Name, $h.Hash, $f.Length, $f.LastWriteTimeUtc)
    }
}

Add-L "---- end addin *.xlam audit ----"
