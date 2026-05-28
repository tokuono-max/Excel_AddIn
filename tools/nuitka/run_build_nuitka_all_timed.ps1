# Run build_nuitka_all_core.bat with [et H:MM:SS.fff]  prefix on every line (merged stdout+stderr via cmd 2>&1).
# "et" = elapsed time; two spaces after ] before body. Console + UTF-8 log. Line-by-line read (live).
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-EtPrefix {
    param([System.Diagnostics.Stopwatch]$Stopwatch)
    $tm = [int64][math]::Floor($Stopwatch.Elapsed.TotalMilliseconds)
    $h = [int][math]::Floor($tm / [int64]3600000)
    $m = [int][math]::Floor(($tm % [int64]3600000) / [int64]60000)
    $s = [int][math]::Floor(($tm % [int64]60000) / [int64]1000)
    $f = [int]($tm % [int64]1000)
    return "[et {0}:{1:00}:{2:00}.{3:000}]  " -f $h, $m, $s, $f
}

$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot "..\.."))
$bat = Join-Path $PSScriptRoot "build_nuitka_all_core.bat"
if (-not (Test-Path -LiteralPath $bat)) {
    throw "Missing: $bat"
}

$logDir = Join-Path $repoRoot "logs\nuitka"
if (-not (Test-Path -LiteralPath $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}
$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$logFile = Join-Path $logDir "build_console_$ts.log"
$logLatest = Join-Path $logDir "build_console_latest.log"

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$logLock = New-Object System.Object

function Write-PrefixedLine {
    param([string]$Text)
    $line = (Get-EtPrefix -Stopwatch $sw) + $Text
    Write-Host $line
    [void][System.Threading.Monitor]::Enter($logLock)
    try {
        Add-Content -LiteralPath $logFile -Encoding utf8 -Value $line
    } finally {
        [void][System.Threading.Monitor]::Exit($logLock)
    }
}

$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = "cmd.exe"
$batQ = '"' + $bat + '"'
$psi.Arguments = "/d /s /c call $batQ 2>&1 & exit /b %ERRORLEVEL%"
$psi.WorkingDirectory = $PSScriptRoot
$psi.UseShellExecute = $false
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError = $false
$psi.CreateNoWindow = $true

$cp = [System.Globalization.CultureInfo]::CurrentCulture.TextInfo.OEMCodePage
try {
    $enc = [System.Text.Encoding]::GetEncoding($cp)
} catch {
    $enc = [Console]::OutputEncoding
}
$psi.StandardOutputEncoding = $enc

$proc = New-Object System.Diagnostics.Process
$proc.StartInfo = $psi
[void]$proc.Start()
$reader = $proc.StandardOutput
try {
    while (($line = $reader.ReadLine()) -ne $null) {
        Write-PrefixedLine -Text $line
    }
} finally {
    if (-not $proc.HasExited) {
        $proc.WaitForExit()
    }
}
if (-not $proc.HasExited) {
    $proc.WaitForExit()
}
$exitCode = $proc.ExitCode

$tail = (Get-EtPrefix -Stopwatch $sw) + "[Nuitka] build_console log: $logLatest"
Write-Host $tail
Add-Content -LiteralPath $logFile -Encoding utf8 -Value $tail
$fin = (Get-EtPrefix -Stopwatch $sw) + "===== build finished exit_code=$exitCode ====="
Add-Content -LiteralPath $logFile -Encoding utf8 -Value $fin
Copy-Item -LiteralPath $logFile -Destination $logLatest -Force
exit $exitCode
