# Append one UTF-8 line to a staging audit log (called from assemble_csv_tool_staging.bat).
param(
    [Parameter(Mandatory = $true)]
    [string]$LogPath,
    [Parameter(Mandatory = $true)]
    [string]$Line
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$dir = Split-Path -Parent -LiteralPath $LogPath
if ($dir -and -not (Test-Path -LiteralPath $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}
Add-Content -LiteralPath $LogPath -Encoding utf8 -Value $Line
