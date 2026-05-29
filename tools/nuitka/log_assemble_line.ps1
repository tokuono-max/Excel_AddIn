# Append one UTF-8 line to a staging audit log (called from assemble_csv_tool_staging.bat).
# Args: <LogPath> <Line>  (positional; avoids ParameterBinding issues on some hosts.)
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ($args.Count -lt 2) {
    throw "usage: log_assemble_line.ps1 <LogPath> <Line>"
}
$LogPath = [string]$args[0]
$Line = [string]$args[1]

# Avoid Split-Path -LiteralPath (AmbiguousParameterSet on some Windows PowerShell builds).
$dir = [System.IO.Path]::GetDirectoryName($LogPath)
if ($dir -and -not (Test-Path -LiteralPath $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}
Add-Content -LiteralPath $LogPath -Encoding utf8 -Value $Line
