# -----------------------------------------------------------------------------
# promote_staging_shared_dlls.ps1  (tools\nuitka\legacy — 既定ビルドでは未使用)
# 目的: ステージング `app\bridge` 等に重複配置されたランタイムを `app\shared\` に集約する。
#       - 第1段: EXE 直下の同一 DLL（ファイル単位のシンボリックリンク）
#       - 第2段: 同一ツリーのディレクトリ束（例: numpy.libs）を shared に1本化し、各モジュールへ NTFS ジャンクション(mklink /J)を置く
#         （ディレクトリのシンボリックリンクは管理者権限を要することが多いため）
# 前提: assemble_csv_tool_staging.bat 完了後、dedupe_staging_hardlinks.ps1 の前に実行する想定。
# 注意: シンボリックリンク作成には Windows の「開発者モード」または管理者権限が要る場合がある。
#       失敗したパスはコピーに戻し警告を出す（ビルドは続行）。
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

function Get-Sha256Hex {
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

function New-ModuleRootSymlink {
    param(
        [Parameter(Mandatory = $true)][string]$LinkPath,
        [Parameter(Mandatory = $true)][string]$TargetPath
    )
    # -Target の相対パスはプロセスのカレントに依存し得るため、呼び出し側でリンク親基準の絶対パスを渡す。
    New-Item -ItemType SymbolicLink -Path $LinkPath -Target $TargetPath -ErrorAction Stop | Out-Null
}

function Resolve-AbsoluteChildPath {
    param(
        [Parameter(Mandatory = $true)][string]$BaseDir,
        [Parameter(Mandatory = $true)][string]$RelativeOrAbsolute
    )
    return [System.IO.Path]::GetFullPath((Join-Path $BaseDir $RelativeOrAbsolute))
}

function Get-DirectoryContentManifest {
    param([Parameter(Mandatory = $true)][string]$DirPath)
    if (-not (Test-Path -LiteralPath $DirPath -PathType Container)) {
        return $null
    }
    $rootItem = Get-Item -LiteralPath $DirPath
    if ($rootItem.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
        return $null
    }
    $rootFull = $rootItem.FullName.TrimEnd("\")
    $files = @(Get-ChildItem -LiteralPath $DirPath -Recurse -File | Sort-Object { $_.FullName })
    $lines = foreach ($f in $files) {
        $rel = $f.FullName.Substring($rootFull.Length).TrimStart("\")
        $h = Get-Sha256Hex -LiteralPath $f.FullName
        "$rel|$h"
    }
    return ($lines -join "`n")
}

function Invoke-MklinkJunction {
    param(
        [Parameter(Mandatory = $true)][string]$LinkPath,
        [Parameter(Mandatory = $true)][string]$AbsoluteTarget
    )
    $quotedLink = '"' + $LinkPath + '"'
    $quotedTarget = '"' + $AbsoluteTarget + '"'
    cmd /c "mklink /J $quotedLink $quotedTarget" | Out-Null
    if ($LASTEXITCODE -ne 0) {
        throw "mklink /J failed: $LinkPath -> $AbsoluteTarget (exit $LASTEXITCODE)"
    }
}

function Invoke-RobocopyMirrorCopy {
    param(
        [Parameter(Mandatory = $true)][string]$SourceDir,
        [Parameter(Mandatory = $true)][string]$DestDir
    )
    if (-not (Test-Path -LiteralPath $DestDir -PathType Container)) {
        New-Item -ItemType Directory -Path $DestDir -Force | Out-Null
    }
    $quotedSrc = '"' + $SourceDir + '"'
    $quotedDst = '"' + $DestDir + '"'
    cmd /c "robocopy $quotedSrc $quotedDst /E /COPY:DAT /DCOPY:DAT /R:2 /W:2 /NFL /NDL /NJH /NJS /nc /ns /np" | Out-Null
    $rc = $LASTEXITCODE
    if ($rc -gt 7) {
        throw "robocopy failed: $SourceDir -> $DestDir (exit $rc)"
    }
}

$stagingAbs = Resolve-RepoAbsolutePath -Path $StagingRoot
$appRoot = Join-Path $stagingAbs "app"
$sharedRoot = Join-Path $appRoot "shared"

if (-not (Test-Path -LiteralPath $appRoot -PathType Container)) {
    throw "promote_shared: app directory not found: $appRoot"
}

$moduleDirs = @("bridge", "svc_server", "ui_server", "xlwings_short_runner")

# 第1段階: 各 EXE 直下に同名で並ぶランタイム DLL のみ
$whitelist = @(
    "python312.dll",
    "python3.dll",
    "libcrypto-3.dll",
    "libssl-3.dll",
    "sqlite3.dll"
)

Write-Host "[promote_shared] Staging: $stagingAbs"
Write-Host "[promote_shared] Shared root will be: $sharedRoot"

$promoted = 0
$skipped = 0

foreach ($name in $whitelist) {
    $paths = New-Object System.Collections.Generic.List[string]

    foreach ($d in $moduleDirs) {
        $p = Join-Path (Join-Path $appRoot $d) $name
        if (-not (Test-Path -LiteralPath $p -PathType Leaf)) {
            continue
        }
        $item = Get-Item -LiteralPath $p
        if ($item.LinkType) {
            Write-Host "[promote_shared] skip (already reparse point): $p"
            continue
        }
        $paths.Add($p) | Out-Null
    }

    if ($paths.Count -lt 2) {
        $skipped++
        continue
    }

    $hashes = @(foreach ($p in $paths) { Get-Sha256Hex -LiteralPath $p })
    $distinct = @($hashes | Select-Object -Unique)
    if ($distinct.Count -ne 1) {
        Write-Warning "[promote_shared] hash mismatch for '$name' — skip group"
        $skipped++
        continue
    }

    if (-not (Test-Path -LiteralPath $sharedRoot -PathType Container)) {
        New-Item -ItemType Directory -Path $sharedRoot -Force | Out-Null
    }

    $sharedFile = Join-Path $sharedRoot $name
    if (-not (Test-Path -LiteralPath $sharedFile -PathType Leaf)) {
        Copy-Item -LiteralPath $paths[0] -Destination $sharedFile -Force
    }
    else {
        $hShared = Get-Sha256Hex -LiteralPath $sharedFile
        if ($hShared -ne @($distinct)[0]) {
            throw "[promote_shared] existing shared\$name has different hash than app copies"
        }
    }

    $relativeTarget = Join-Path (Join-Path ".." "shared") $name

    foreach ($p in $paths) {
        Remove-Item -Path $p -Force
        $linkParent = Split-Path -Parent $p
        $absFileTarget = Resolve-AbsoluteChildPath -BaseDir $linkParent -RelativeOrAbsolute $relativeTarget
        try {
            New-ModuleRootSymlink -LinkPath $p -TargetPath $absFileTarget
            Write-Host "[promote_shared] symlink: $p -> $absFileTarget"
        }
        catch {
            Copy-Item -Path $sharedFile -Destination $p -Force
            Write-Warning "[promote_shared] symlink failed for '$p' ($($_.Exception.Message)); restored file copy from shared"
        }
    }
    $promoted++
}

# 第1段の補完: app\shared に正本があるが、某モジュール直下にホワイトリスト DLL が無い場合にリンクまたはコピーで埋める。
# （Nuitka 出力差で第1段の $paths に当該モジュールが含まれなかった場合の取りこぼし防止。svc_host は子 PATH に app\shared も付与するが、ローダ互換のため各所にも置く。）
foreach ($name in $whitelist) {
    $sharedFile = Join-Path $sharedRoot $name
    if (-not (Test-Path -LiteralPath $sharedFile -PathType Leaf)) {
        continue
    }
    foreach ($d in $moduleDirs) {
        $mp = Join-Path (Join-Path $appRoot $d) $name
        if (Test-Path -LiteralPath $mp -PathType Leaf) {
            continue
        }
        $relativeTarget = Join-Path (Join-Path ".." "shared") $name
        $linkParent = Join-Path $appRoot $d
        $absFileTarget = Resolve-AbsoluteChildPath -BaseDir $linkParent -RelativeOrAbsolute $relativeTarget
        try {
            New-ModuleRootSymlink -LinkPath $mp -TargetPath $absFileTarget
            Write-Host "[promote_shared] whitelist gap-fill symlink: $mp -> $absFileTarget"
        }
        catch {
            Copy-Item -LiteralPath $sharedFile -Destination $mp -Force
            Write-Warning "[promote_shared] whitelist gap-fill copy: $mp ($($_.Exception.Message))"
        }
    }
}

# -----------------------------------------------------------------------------
# 第2段階（段階的）: 同一ツリーが複数モジュールに重複する「ディレクトリ束」を app\shared に1本化し、
#                 各モジュール直下にはディレクトリのシンボリックリンクを置く。
# 条件: 対象ディレクトリが実ディレクトリ（再解析ポイントでない）かつ、少なくとも2モジュールに存在し、
#       相対パス＋SHA256 のマニフェストがすべて一致するときのみ実行する（同一ボリューム上のジャンクション前提）。
# -----------------------------------------------------------------------------
# Directory trees that are byte-identical between svc_server and ui_server (verified same venv / Nuitka graph).
# Skipped automatically if manifest differs or fewer than two modules contain the bundle.
# Order: smaller / safer bundles first; large packages last.
$directoryBundles = @(
    "jaraco",
    "shiboken6",
    "xlwings",
    "pytz",
    "tzdata",
    "numpy",
    "numpy.libs",
    "pandas",
    "pandas.libs",
    "_polars_runtime_32"
)

$promotedDirs = 0
$skippedDirs = 0

foreach ($bundleName in $directoryBundles) {
    $dirPaths = New-Object System.Collections.Generic.List[string]
    $abortDirBundle = $false

    foreach ($d in $moduleDirs) {
        $p = Join-Path (Join-Path $appRoot $d) $bundleName
        if (-not (Test-Path -LiteralPath $p -PathType Container)) {
            continue
        }
        $rootItem = Get-Item -LiteralPath $p
        if ($rootItem.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
            Write-Host "[promote_shared] skip dir bundle '$bundleName' (already reparse point): $p"
            $skippedDirs++
            $abortDirBundle = $true
            break
        }
        $dirPaths.Add($p) | Out-Null
    }

    if ($abortDirBundle) {
        continue
    }

    if ($dirPaths.Count -lt 2) {
        Write-Host "[promote_shared] skip dir bundle '$bundleName' (present in fewer than 2 modules)"
        $skippedDirs++
        continue
    }

    $refManifest = Get-DirectoryContentManifest -DirPath $dirPaths[0]
    if ([string]::IsNullOrEmpty($refManifest)) {
        Write-Warning "[promote_shared] empty or unreadable dir bundle '$bundleName' at $($dirPaths[0])"
        $skippedDirs++
        continue
    }

    $allMatch = $true
    for ($i = 1; $i -lt $dirPaths.Count; $i++) {
        $m = Get-DirectoryContentManifest -DirPath $dirPaths[$i]
        if ($m -ne $refManifest) {
            Write-Warning "[promote_shared] dir bundle '$bundleName' manifest mismatch:`n  ref=$($dirPaths[0])`n  other=$($dirPaths[$i])"
            $allMatch = $false
            break
        }
    }
    if (-not $allMatch) {
        $skippedDirs++
        continue
    }

    $canonicalDir = $dirPaths[0]
    if (-not (Test-Path -LiteralPath $sharedRoot -PathType Container)) {
        New-Item -ItemType Directory -Path $sharedRoot -Force | Out-Null
    }

    $sharedBundle = Join-Path $sharedRoot $bundleName
    if (Test-Path -LiteralPath $sharedBundle -PathType Container) {
        $mExist = Get-DirectoryContentManifest -DirPath $sharedBundle
        if ($mExist -ne $refManifest) {
            throw "[promote_shared] existing shared\$bundleName differs from staged app copies"
        }
    }
    else {
        Invoke-RobocopyMirrorCopy -SourceDir $canonicalDir -DestDir $sharedBundle
        $mAfter = Get-DirectoryContentManifest -DirPath $sharedBundle
        if ($mAfter -ne $refManifest) {
            throw "[promote_shared] robocopy to shared\$bundleName did not reproduce source tree"
        }
    }

    $relativeDirTarget = Join-Path (Join-Path ".." "shared") $bundleName
    $dirLinkAnyFailure = $false

    foreach ($p in $dirPaths) {
        if (Test-Path -LiteralPath $p) {
            $it = Get-Item -LiteralPath $p
            if ($it.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
                Write-Host "[promote_shared] skip remove (already reparse point): $p"
                continue
            }
        }
        Remove-Item -LiteralPath $p -Recurse -Force
        $parentDir = Split-Path -Parent $p
        $absDirTarget = Resolve-AbsoluteChildPath -BaseDir $parentDir -RelativeOrAbsolute $relativeDirTarget
        try {
            Invoke-MklinkJunction -LinkPath $p -AbsoluteTarget $absDirTarget
            Write-Host "[promote_shared] dir junction: $p -> $absDirTarget"
        }
        catch {
            $dirLinkAnyFailure = $true
            Invoke-RobocopyMirrorCopy -SourceDir $sharedBundle -DestDir $p
            Write-Warning "[promote_shared] dir junction failed for '$p' ($($_.Exception.Message)); restored directory copy from shared"
        }
    }

    if ($dirLinkAnyFailure) {
        Write-Warning "[promote_shared] dir bundle '$bundleName': junction step failed for one or more modules; removing shared\$bundleName to avoid duplicate full trees"
        if (Test-Path -LiteralPath $sharedBundle -PathType Container) {
            Remove-Item -LiteralPath $sharedBundle -Recurse -Force
        }
        $skippedDirs++
        continue
    }

    $promotedDirs++
}

Write-Host "[promote_shared] Done. Promoted file kinds: $promoted, skipped (single or missing): $skipped | Dir bundles promoted: $promotedDirs, dir bundles skipped: $skippedDirs"
exit 0
