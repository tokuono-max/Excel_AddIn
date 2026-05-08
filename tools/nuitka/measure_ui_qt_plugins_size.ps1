# One-off helper: report PySide6 plugins\ subfolder sizes under ui_server staging (MB).
param([string]$UiServerRoot = "dist\CSV_Tool\app\bin")
$repo = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$plugins = Join-Path (Join-Path (Join-Path $repo $UiServerRoot) "PySide6") "qt-plugins"
if (-not (Test-Path -LiteralPath $plugins)) {
    Write-Host "[measure_ui_qt_plugins] not found: $plugins"
    exit 1
}
Get-ChildItem -LiteralPath $plugins -Directory | ForEach-Object {
    $sum = (Get-ChildItem -LiteralPath $_.FullName -Recurse -File -ErrorAction SilentlyContinue |
        Measure-Object -Property Length -Sum).Sum
    [PSCustomObject]@{ Name = $_.Name; MB = [math]::Round($sum / 1MB, 2) }
} | Sort-Object MB -Descending | Format-Table -AutoSize
