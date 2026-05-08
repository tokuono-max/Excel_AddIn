# -*- coding: utf-8 -*-
"""
Packaged deployment updater:
- Read shared catalog.json
- Apply config payload silently (failure only notifies)
- Prompt user for bin update (Now/Later); 「今すぐ」で **bin.patch（差分）** または **bin.full（フル）**
  zip を検証し、Excel 終了後にバックグラウンドで反映（差分はファイルマージ、フルは従来どおり MIR）。

See docs\\インストールと運用（利用者・運用向け）.md §3.3, §4.2.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import uuid
import zipfile
from pathlib import Path
from typing import Any

from packaging.version import InvalidVersion, Version

from core import core_env, runtime_layout
from core.core_log import append_text_with_cap, get_logger
from core.runtime_layout import packaged_spawn_requested
from core.patch_manifest import materialize_manifest_patch_zip as _materialize_patch_zip_for_worker
from core.update_state import build_paths, read_pending, write_pending
try:
    from core import core_cst as cst
except Exception:
    cst = None  # type: ignore

logger = get_logger(__name__)

_UPDATE_CHECK_UI_LOCK = threading.Lock()

# True when apply_pending_update returned deferred (No on pending-apply confirm).
# Suppresses duplicate bin prompt in maybe_check_updates_on_startup same process launch.
_SUPPRESS_STARTUP_BIN_PROMPT_AFTER_PENDING_DEFER = False

ENV_DEPLOY_ROOT = "HC_DEPLOY_ROOT"
ENV_CATALOG_PATH = "HC_CATALOG_PATH"
ENV_UPDATE_CHECK_AT_STARTUP = "HC_UPDATE_CHECK_AT_STARTUP"
ENV_USE_UPDATER_EXE = "HC_USE_UPDATER_EXE"
# Dev-only: allow python.exe + temp helper when hc_updater.exe is missing (not for production packages).
ENV_ALLOW_PYTHON_ELEVATE = "HC_ALLOW_PYTHON_ELEVATE"

UPDATE_LOG_REL = Path("logs") / "hc_update.log"
BIN_APPLY_NOTIFY_MARKER_REL = Path("logs") / "bin_apply_success_marker.json"
ADMIN_CATALOG_REL = Path("config") / "catalog_path.txt"

MB_OK = 0x0
MB_YESNO = 0x4
MB_ICONINFORMATION = 0x40
MB_ICONWARNING = 0x30
MB_SETFOREGROUND = 0x10000
IDYES = 6
IDNO = 7


def _get_update_check_msg_cfg() -> dict[str, Any]:
    try:
        if cst is not None:
            raw = cst.get_ui_config_from_file_required("update_check")
            if isinstance(raw, dict):
                m = raw.get("MESSAGES")
                if isinstance(m, dict):
                    return m
    except Exception:
        pass
    return {}


def _um(key: str, default: str, **fmt: Any) -> str:
    cfg = _get_update_check_msg_cfg()
    base = cfg.get(key, default) if isinstance(cfg, dict) else default
    try:
        return str(base).format(**fmt)
    except Exception:
        return str(base)


def _update_check_title() -> str:
    return _um("UPDATE_CHECK_DIALOG_TITLE", "CSV Tool 更新")


# Excel 終了待ち後に app\\bin を差し替えるワーカー（-File 用）。パラメータは JSON 1 ファイル経由。
_APPLY_BIN_WORKER_PS1 = r"""
param(
    [Parameter(Mandatory = $true)]
    [string] $ParamPath
)
$ErrorActionPreference = 'Stop'
$j = Get-Content -LiteralPath $ParamPath -Raw -Encoding UTF8 | ConvertFrom-Json
$InstallRoot = [string]$j.InstallRoot
$ZipPath = [string]$j.ZipPath
$ExpectedSha = [string]$j.ExpectedSha
$LogPath = [string]$j.LogPath
$DisplayVersion = [string]$j.DisplayVersion
$ApplyMode = 'full'
if ($j.PSObject.Properties.Name -contains 'ApplyMode' -and $j.ApplyMode) {
    $ApplyMode = [string]$j.ApplyMode
}
$TargetBinVersion = ''
if ($j.PSObject.Properties.Name -contains 'TargetBinVersion' -and $j.TargetBinVersion) {
    $TargetBinVersion = [string]$j.TargetBinVersion
}
$CleanupPath = ''
if ($j.PSObject.Properties.Name -contains 'CleanupPath' -and $j.CleanupPath) {
    $CleanupPath = [string]$j.CleanupPath
}
$NotifyMarkerPath = ''
if ($j.PSObject.Properties.Name -contains 'NotifyMarkerPath' -and $j.NotifyMarkerPath) {
    $NotifyMarkerPath = [string]$j.NotifyMarkerPath
}

function Write-ApplyLog([string]$m) {
    $maxBytes = 1048576
    $ts = Get-Date -Format 'yyyy-MM-ddTHH:mm:ss'
    $dir = Split-Path -Parent $LogPath
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    Add-Content -LiteralPath $LogPath -Value ($ts + ' ' + $m) -Encoding UTF8
    try {
        $fi = Get-Item -LiteralPath $LogPath -ErrorAction Stop
        if ($fi.Length -gt $maxBytes) {
            $keep = New-Object byte[] $maxBytes
            $fs = [System.IO.File]::Open($LogPath, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::ReadWrite)
            try {
                [void]$fs.Seek(-$maxBytes, [System.IO.SeekOrigin]::End)
                [void]$fs.Read($keep, 0, $maxBytes)
                $fs.SetLength(0)
                [void]$fs.Seek(0, [System.IO.SeekOrigin]::Begin)
                $fs.Write($keep, 0, $maxBytes)
            } finally {
                $fs.Dispose()
            }
        }
    } catch {
    }
}

function Init-PhaseUi() {
    try {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing
        $global:PhaseForm = New-Object System.Windows.Forms.Form
        $global:PhaseForm.Text = 'CSV Tool 更新'
        $global:PhaseForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
        $global:PhaseForm.TopMost = $true
        $global:PhaseForm.Width = 520
        $global:PhaseForm.Height = 190
        $global:PhaseForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $global:PhaseForm.MaximizeBox = $false
        $global:PhaseForm.MinimizeBox = $false

        $global:PhaseTitle = New-Object System.Windows.Forms.Label
        $global:PhaseTitle.Left = 16
        $global:PhaseTitle.Top = 16
        $global:PhaseTitle.Width = 480
        $global:PhaseTitle.Height = 28
        $global:PhaseTitle.Text = '状態: 待機中'
        $global:PhaseTitle.Font = New-Object System.Drawing.Font('Yu Gothic UI', 10, [System.Drawing.FontStyle]::Bold)
        $global:PhaseForm.Controls.Add($global:PhaseTitle)

        $global:PhaseMessage = New-Object System.Windows.Forms.Label
        $global:PhaseMessage.Left = 16
        $global:PhaseMessage.Top = 52
        $global:PhaseMessage.Width = 480
        $global:PhaseMessage.Height = 36
        $global:PhaseMessage.Text = 'Excel をすべて閉じると更新を開始します。'
        $global:PhaseForm.Controls.Add($global:PhaseMessage)

        $global:PhaseProgress = New-Object System.Windows.Forms.ProgressBar
        $global:PhaseProgress.Left = 16
        $global:PhaseProgress.Top = 100
        $global:PhaseProgress.Width = 480
        $global:PhaseProgress.Height = 20
        $global:PhaseProgress.Minimum = 0
        $global:PhaseProgress.Maximum = 100
        $global:PhaseProgress.Value = 0
        $global:PhaseForm.Controls.Add($global:PhaseProgress)

        $global:PhaseForm.Show()
        [System.Windows.Forms.Application]::DoEvents()
    } catch {
        Write-ApplyLog ('apply_bin: phase_ui init failed err=' + $_.Exception.Message)
    }
}

function Set-Phase([string]$PhaseKey, [string]$TitleText, [string]$MessageText, [int]$Progress) {
    Write-ApplyLog ("apply_bin: phase=" + $PhaseKey + " message=" + $MessageText)
    try {
        if ($global:PhaseTitle -ne $null) {
            $global:PhaseTitle.Text = ('状態: ' + $TitleText)
        }
        if ($global:PhaseMessage -ne $null) {
            $global:PhaseMessage.Text = $MessageText
        }
        if ($global:PhaseProgress -ne $null) {
            if ($Progress -lt 0) { $Progress = 0 }
            if ($Progress -gt 100) { $Progress = 100 }
            $global:PhaseProgress.Value = $Progress
        }
        [System.Windows.Forms.Application]::DoEvents()
    } catch {
    }
}

function Close-PhaseUi() {
    try {
        if ($global:PhaseForm -ne $null) {
            $global:PhaseForm.Close()
            $global:PhaseForm.Dispose()
        }
    } catch {
    } finally {
        $global:PhaseForm = $null
        $global:PhaseTitle = $null
        $global:PhaseMessage = $null
        $global:PhaseProgress = $null
    }
}

function Try-SetDisplayValues([string]$DisplayVersion) {
    if (-not $DisplayVersion) { return }
    $displayName = 'CSV Tool'
    $sub = 'Software\Microsoft\Windows\CurrentVersion\Uninstall\{B5E8C4A2-1D3F-4E6B-9C0D-1A2B3C4D5E6F}_is1'
    $paths = @(
        ('Registry::HKEY_LOCAL_MACHINE\' + $sub),
        ('Registry::HKEY_LOCAL_MACHINE\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{B5E8C4A2-1D3F-4E6B-9C0D-1A2B3C4D5E6F}_is1'),
        ('Registry::HKEY_CURRENT_USER\' + $sub),
        ('Registry::HKEY_CURRENT_USER\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{B5E8C4A2-1D3F-4E6B-9C0D-1A2B3C4D5E6F}_is1')
    )
    $target = $null
    foreach ($p in $paths) {
        if (Test-Path -LiteralPath $p) {
            $target = $p
            break
        }
    }
    if (-not $target) {
        Write-ApplyLog ('apply_bin: DisplayName/DisplayVersion key not found or not writable value=' + $DisplayVersion)
        return
    }
    try {
        Set-ItemProperty -LiteralPath $target -Name 'DisplayName' -Value $displayName -Type String
        Set-ItemProperty -LiteralPath $target -Name 'DisplayVersion' -Value $DisplayVersion -Type String
        Write-ApplyLog ('apply_bin: DisplayName/DisplayVersion updated to ' + $displayName + ' / ' + $DisplayVersion + ' at ' + $target)
    } catch {
        Write-ApplyLog ('apply_bin: DisplayName/DisplayVersion write failed at ' + $target + ' err=' + $_.Exception.Message)
    }
}

try {
    Init-PhaseUi
    Set-Phase 'wait_excel_exit' '待機中' 'Excel をすべて閉じると更新を開始します。' 5
    Write-ApplyLog 'apply_bin: waiting for Excel (EXCEL) to exit'
    while ($true) {
        $ex = Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue
        if (-not $ex) { break }
        Start-Sleep -Seconds 2
    }
    Start-Sleep -Seconds 1
    Set-Phase 'start' '開始' '更新処理を開始しています。' 15

    $downloadTmp = Join-Path $env:TEMP ('csv_tool_bin_download_' + [Guid]::NewGuid().ToString())
    New-Item -ItemType Directory -Path $downloadTmp | Out-Null
    $zipLocal = Join-Path $downloadTmp ([System.IO.Path]::GetFileName($ZipPath))
    Set-Phase 'download_copy' 'ダウンロード中（コピー中）' '更新ファイルを取得しています。' 30
    Copy-Item -LiteralPath $ZipPath -Destination $zipLocal -Force

    Write-ApplyLog ('apply_bin: verifying zip mode=' + $ApplyMode)
    if ($ExpectedSha) {
        $h = (Get-FileHash -LiteralPath $zipLocal -Algorithm SHA256).Hash.ToLowerInvariant()
        if ($h -ne $ExpectedSha.ToLowerInvariant()) {
            throw ('sha256 mismatch got=' + $h)
        }
    }
    $tmp = Join-Path $env:TEMP ('csv_tool_bin_extract_' + [Guid]::NewGuid().ToString())
    New-Item -ItemType Directory -Path $tmp | Out-Null
    try {
        Set-Phase 'extract' '展開中' '更新ファイルを展開しています。' 50
        Expand-Archive -LiteralPath $zipLocal -DestinationPath $tmp -Force
        Set-Phase 'apply' '適用中' '更新を適用しています。' 75
        if ($ApplyMode -eq 'patch') {
            $patchApp = Join-Path $tmp 'app\bin'
            $patchAddin = Join-Path $tmp 'addin'
            if (-not (Test-Path -LiteralPath $patchApp) -and -not (Test-Path -LiteralPath $patchAddin)) {
                throw 'invalid patch zip: need app\bin and/or addin'
            }
            if (-not $TargetBinVersion) {
                throw 'patch apply requires TargetBinVersion'
            }
            function Copy-MergeTree([string]$SrcRoot, [string]$DstRoot) {
                if (-not (Test-Path -LiteralPath $SrcRoot)) { return }
                Get-ChildItem -LiteralPath $SrcRoot -Recurse -File | ForEach-Object {
                    $rel = $_.FullName.Substring($SrcRoot.Length).TrimStart([char]'\', [char]'/')
                    $dest = Join-Path $DstRoot $rel
                    $destDir = Split-Path -Parent $dest
                    if (-not (Test-Path -LiteralPath $destDir)) {
                        New-Item -ItemType Directory -Path $destDir -Force | Out-Null
                    }
                    Copy-Item -LiteralPath $_.FullName -Destination $dest -Force
                }
            }
            if (Test-Path -LiteralPath $patchApp) {
                $destBin = Join-Path $InstallRoot 'app\bin'
                if (-not (Test-Path -LiteralPath $destBin)) {
                    New-Item -ItemType Directory -Path $destBin -Force | Out-Null
                }
                Copy-MergeTree $patchApp $destBin
            }
            if (Test-Path -LiteralPath $patchAddin) {
                $addinDst = Join-Path $InstallRoot 'addin'
                if (-not (Test-Path -LiteralPath $addinDst)) {
                    New-Item -ItemType Directory -Path $addinDst -Force | Out-Null
                }
                Copy-MergeTree $patchAddin $addinDst
            }
            $utf8 = New-Object System.Text.UTF8Encoding($false)
            $vpath = Join-Path $InstallRoot 'VERSION.txt'
            [System.IO.File]::WriteAllText($vpath, ($TargetBinVersion + "`r`n"), $utf8)
            Write-ApplyLog ('apply_bin: patch merged TargetBinVersion=' + $TargetBinVersion)
        } else {
            $appBin = Join-Path $tmp 'app\bin'
            if (-not (Test-Path -LiteralPath $appBin)) {
                throw 'invalid zip: missing app\bin (expect make_release_payloads bin full layout)'
            }
            $dest = Join-Path $InstallRoot 'app\bin'
            if (-not (Test-Path -LiteralPath $dest)) {
                New-Item -ItemType Directory -Path $dest -Force | Out-Null
            }
            & robocopy.exe $appBin $dest /MIR /NFL /NDL /NJH /NJS /nc /ns /np | Out-Null
            if ($LASTEXITCODE -ge 8) {
                throw ('robocopy app\bin failed exit ' + $LASTEXITCODE)
            }
            $vsrc = Join-Path $tmp 'VERSION.txt'
            if ($TargetBinVersion) {
                $utf8 = New-Object System.Text.UTF8Encoding($false)
                [System.IO.File]::WriteAllText((Join-Path $InstallRoot 'VERSION.txt'), ($TargetBinVersion + "`r`n"), $utf8)
            } elseif (Test-Path -LiteralPath $vsrc) {
                Copy-Item -LiteralPath $vsrc -Destination (Join-Path $InstallRoot 'VERSION.txt') -Force
            }
            $addinSrc = Join-Path $tmp 'addin'
            $addinDst = Join-Path $InstallRoot 'addin'
            if (Test-Path -LiteralPath $addinSrc) {
                if (-not (Test-Path -LiteralPath $addinDst)) {
                    New-Item -ItemType Directory -Path $addinDst -Force | Out-Null
                }
                & robocopy.exe $addinSrc $addinDst /MIR /NFL /NDL /NJH /NJS /nc /ns /np | Out-Null
                if ($LASTEXITCODE -ge 8) {
                    throw ('robocopy addin failed exit ' + $LASTEXITCODE)
                }
            }
        }
        Try-SetDisplayValues $DisplayVersion
        if ($NotifyMarkerPath) {
            try {
                $mDir = Split-Path -Parent $NotifyMarkerPath
                if ($mDir -and -not (Test-Path -LiteralPath $mDir)) {
                    New-Item -ItemType Directory -Path $mDir -Force | Out-Null
                }
                $meta = @{
                    ts = (Get-Date -Format 'yyyy-MM-ddTHH:mm:ss')
                    target_bin_version = $TargetBinVersion
                    display_version = $DisplayVersion
                    log_path = $LogPath
                } | ConvertTo-Json -Compress
                [System.IO.File]::WriteAllText($NotifyMarkerPath, $meta, (New-Object System.Text.UTF8Encoding($false)))
                Write-ApplyLog ('apply_bin: success_notify marker_written path=' + $NotifyMarkerPath)
            } catch {
                Write-ApplyLog ('apply_bin: success_notify marker_write_failed err=' + $_.Exception.Message)
            }
        }
        Write-ApplyLog 'apply_bin: success'
        Set-Phase 'done' '完了' '更新が完了しました。Excel を再起動してください。' 100
        try {
            Add-Type -AssemblyName System.Windows.Forms
            [void][System.Windows.Forms.MessageBox]::Show(
                ('bin 更新が完了しました。' + [Environment]::NewLine + [Environment]::NewLine +
                 'Excel を再起動してください。' + [Environment]::NewLine + [Environment]::NewLine + 'ログ: ' + $LogPath),
                'CSV Tool 更新',
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            )
            Write-ApplyLog 'apply_bin: success_notify shown'
        } catch {
            Write-ApplyLog ('apply_bin: success_notify failed err=' + $_.Exception.Message)
        }
    } finally {
        Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $downloadTmp -Recurse -Force -ErrorAction SilentlyContinue
    }
} catch {
    Set-Phase 'error' '失敗' '更新処理中にエラーが発生しました。' 100
    Write-ApplyLog ('apply_bin: ERROR ' + $_.Exception.Message)
    try {
        Add-Type -AssemblyName System.Windows.Forms
        [void][System.Windows.Forms.MessageBox]::Show(
            ('bin の自動更新に失敗しました。' + [Environment]::NewLine + [Environment]::NewLine +
             $_.Exception.Message + [Environment]::NewLine + [Environment]::NewLine + 'ログ: ' + $LogPath),
            'CSV Tool 更新',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
    } catch { }
} finally {
    Close-PhaseUi
    if ($CleanupPath) {
        try {
            if (Test-Path -LiteralPath $CleanupPath -PathType Leaf) {
                Remove-Item -LiteralPath $CleanupPath -Force -ErrorAction SilentlyContinue
            } elseif (Test-Path -LiteralPath $CleanupPath -PathType Container) {
                Remove-Item -LiteralPath $CleanupPath -Recurse -Force -ErrorAction SilentlyContinue
            }
        } catch { }
    }
    Remove-Item -LiteralPath $ParamPath -Force -ErrorAction SilentlyContinue
    if ($PSCommandPath) {
        Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction SilentlyContinue
    }
}
"""


def _install_root() -> Path | None:
    raw = (os.environ.get("HC_INSTALL_ROOT") or "").strip()
    if not raw:
        return None
    p = Path(raw)
    return p if p.is_dir() else None


def _update_log_path(install_root: Path | None) -> Path:
    if install_root is not None:
        return install_root / UPDATE_LOG_REL
    temp = os.environ.get("TEMP", "C:\\Temp")
    return Path(temp) / "csv_tool" / "hc_update.log"


def _bin_apply_success_marker_path(install_root: Path | None) -> Path:
    if install_root is not None:
        return install_root / BIN_APPLY_NOTIFY_MARKER_REL
    temp = os.environ.get("TEMP", "C:\\Temp")
    return Path(temp) / "csv_tool" / "bin_apply_success_marker.json"


def _append_update_log(install_root: Path | None, line: str) -> None:
    log_path = _update_log_path(install_root)
    try:
        ts = time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime())
        append_text_with_cap(log_path, f"{ts} {line}\n")
    except OSError as e:
        logger.warning("packaged_update log failed: %s", e)


def _is_process_elevated() -> bool:
    if os.name != "nt":
        return False
    try:
        import ctypes

        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def _resolve_python_for_elevation(install_root: Path) -> tuple[str, list[tuple[str, bool]]]:
    candidates: list[Path] = []
    candidates.append(install_root / "app" / "bin" / "python.exe")
    try:
        ir = runtime_layout.install_root()
        if ir is not None:
            candidates.append(ir / "app" / "bin" / "python.exe")
    except Exception:
        pass
    try:
        base_exe = Path(str(getattr(sys, "_base_executable", "") or "")).resolve()
        if str(base_exe):
            candidates.append(base_exe)
    except Exception:
        pass
    try:
        base_prefix_exe = Path(str(getattr(sys, "base_prefix", "") or "")).resolve() / "python.exe"
        candidates.append(base_prefix_exe)
    except Exception:
        pass
    candidates.append(Path(str(sys.executable or "")).resolve())
    checked: list[tuple[str, bool]] = []
    for c in candidates:
        ok = False
        try:
            ok = c.is_file()
        except Exception:
            ok = False
        checked.append((str(c), ok))
        if ok:
            return str(c), checked
    return str(candidates[-1]), checked


def _run_pending_apply_elevated(install_root: Path) -> tuple[bool, str]:
    """
    apply_pending_update を昇格プロセスで実行する。
    既定: app/bin/hc_updater.exe --job（配布標準。python.exe 不在でも可）
    開発のみ: HC_ALLOW_PYTHON_ELEVATE=1 かつ python 解決時、一時 .py 経路
    戻り値: (ok, reason_or_detail)
    """
    if os.name != "nt":
        _append_update_log(install_root, "elevate: skip reason=elevation_not_supported")
        return False, "elevation_not_supported"

    tmpdir = install_root / "update" / "tmp"
    tmpdir.mkdir(parents=True, exist_ok=True)
    uid = uuid.uuid4().hex[:10]
    result_path = tmpdir / f"csv_tool_apply_pending_admin_result_{uid}.json"
    job_path = tmpdir / f"csv_tool_apply_pending_job_{uid}.json"
    ps_out_path = tmpdir / f"csv_tool_apply_pending_admin_ps_stdout_{uid}.log"
    ps_err_path = tmpdir / f"csv_tool_apply_pending_admin_ps_stderr_{uid}.log"
    helper_trace_path = tmpdir / f"csv_tool_apply_pending_admin_trace_{uid}.log"
    log_file = _update_log_path(install_root)
    job_payload = {
        "JobType": "apply_pending",
        "InstallRoot": str(install_root.resolve()),
        "ResultPath": str(result_path.resolve()),
        "LogPath": str(log_file.resolve()),
    }
    job_path.write_text(json.dumps(job_payload, ensure_ascii=False), encoding="utf-8")

    work_dir = install_root / "app" / "bin"
    if not work_dir.is_dir():
        work_dir = install_root

    updater_exe = install_root / "app" / "bin" / "hc_updater.exe"
    use_updater_exe = updater_exe.is_file()
    allow_python = core_env.truthy(os.environ.get(ENV_ALLOW_PYTHON_ELEVATE) or "", empty_means_false=True)

    helper: Path | None = None
    py_cmd = ""
    ps = ""

    if use_updater_exe:
        py_cmd = str(updater_exe)
        _append_update_log(
            install_root,
            "elevate: launcher_kind=hc_updater_exe path={p} job={j} result={res} wd={wd}".format(
                p=updater_exe,
                j=job_path,
                res=result_path,
                wd=work_dir,
            ),
        )
        ps = (
            "try {{ "
            "$p = Start-Process -FilePath '{exe}' -ArgumentList @('--job','{job}') -WorkingDirectory '{wd}' -Verb RunAs "
            "-PassThru -Wait; "
            "Write-Output ('__ELEVATE_EXIT:' + $p.ExitCode); "
            "exit $p.ExitCode "
            "}} catch {{ "
            "$msg = $_.Exception.Message; "
            "$fqid = $_.FullyQualifiedErrorId; "
            "$cat = $_.CategoryInfo; "
            "$pos = ''; if ($_.InvocationInfo -and $_.InvocationInfo.PositionMessage) {{ $pos = $_.InvocationInfo.PositionMessage }}; "
            "$stk = $_.ScriptStackTrace; "
            "Write-Output ('__ELEVATE_PSERR:' + $msg); "
            "Write-Output ('__ELEVATE_PSERR_DETAIL:msg=' + $msg + '|fqid=' + $fqid + '|category=' + $cat + '|pos=' + $pos + '|stack=' + $stk); "
            "exit 199 "
            "}}"
        ).format(
            exe=str(updater_exe).replace("'", "''"),
            job=str(job_path).replace("'", "''"),
            wd=str(work_dir).replace("'", "''"),
        )
    elif allow_python:
        py_cmd, py_candidates = _resolve_python_for_elevation(install_root)
        helper = tmpdir / f"csv_tool_apply_pending_admin_{uid}.py"
        helper.write_text(
            "\n".join(
                [
                    "from __future__ import annotations",
                    "import argparse, json, os, sys, traceback",
                    "from pathlib import Path",
                    "def main() -> int:",
                    "    ap = argparse.ArgumentParser()",
                    "    ap.add_argument('--install-root', required=True)",
                    "    ap.add_argument('--result-path', required=True)",
                    "    ap.add_argument('--trace-path', required=False, default='')",
                    "    args = ap.parse_args()",
                    "    out = {'ok': False, 'applied': False}",
                    "    trace_path = Path(args.trace_path) if args.trace_path else None",
                    "    def _trace(msg: str) -> None:",
                    "        if not trace_path:",
                    "            return",
                    "        try:",
                    "            trace_path.parent.mkdir(parents=True, exist_ok=True)",
                    "            with trace_path.open('a', encoding='utf-8') as f:",
                    "                f.write(msg + '\\n')",
                    "        except Exception:",
                    "            pass",
                    "    _trace('helper:start executable={e} cwd={cwd} argv={argv}'.format(e=sys.executable, cwd=os.getcwd(), argv=sys.argv))",
                    "    try:",
                    "        _trace('helper:before_import apply_pending_update')",
                    "        from bootstrap.update_bootstrap import apply_pending_update",
                    "        _trace('helper:after_import apply_pending_update')",
                    "        _trace('helper:before_apply_pending_update')",
                    "        res = apply_pending_update(Path(args.install_root))",
                    "        _trace('helper:after_apply_pending_update type={t}'.format(t=type(res).__name__))",
                    "        if isinstance(res, dict):",
                    "            out.update(res)",
                    "            out['ok'] = bool(res.get('ok', True))",
                    "        else:",
                    "            out = {'ok': True, 'applied': False}",
                    "    except Exception as e:",
                    "        out = {'ok': False, 'applied': False, 'error': f'{type(e).__name__}: {e}', 'traceback': traceback.format_exc()}",
                    "        _trace('helper:exception {t}: {m}'.format(t=type(e).__name__, m=e))",
                    "        _trace(out.get('traceback', ''))",
                    "    Path(args.result_path).write_text(json.dumps(out, ensure_ascii=False), encoding='utf-8')",
                    "    _trace('helper:result_written path={p} ok={ok}'.format(p=args.result_path, ok=bool(out.get('ok'))))",
                    "    return 0 if bool(out.get('ok')) else 1",
                    "if __name__ == '__main__':",
                    "    raise SystemExit(main())",
                    "",
                ]
            ),
            encoding="utf-8",
        )
        if not Path(py_cmd).is_file():
            _append_update_log(install_root, f"elevate: launcher_missing path={py_cmd}")
            try:
                job_path.unlink(missing_ok=True)
            except OSError:
                pass
            return False, "uac_start_failed_path"
        _append_update_log(
            install_root,
            "elevate: launcher_kind=python_helper py={py} helper={helper} result={res} wd={wd} candidates={cands}".format(
                py=py_cmd,
                helper=helper,
                res=result_path,
                wd=work_dir,
                cands=[{"path": p, "exists": ok} for p, ok in py_candidates],
            ),
        )
        ps = (
            "try {{ "
            "$p = Start-Process -FilePath '{py}' -ArgumentList @('-u','{helper}','--install-root','{root}',"
            "'--result-path','{res}','--trace-path','{trace}') -WorkingDirectory '{wd}' -Verb RunAs "
            "-PassThru -Wait; "
            "Write-Output ('__ELEVATE_EXIT:' + $p.ExitCode); "
            "exit $p.ExitCode "
            "}} catch {{ "
            "$msg = $_.Exception.Message; "
            "$fqid = $_.FullyQualifiedErrorId; "
            "$cat = $_.CategoryInfo; "
            "$pos = ''; if ($_.InvocationInfo -and $_.InvocationInfo.PositionMessage) {{ $pos = $_.InvocationInfo.PositionMessage }}; "
            "$stk = $_.ScriptStackTrace; "
            "Write-Output ('__ELEVATE_PSERR:' + $msg); "
            "Write-Output ('__ELEVATE_PSERR_DETAIL:msg=' + $msg + '|fqid=' + $fqid + '|category=' + $cat + '|pos=' + $pos + '|stack=' + $stk); "
            "exit 199 "
            "}}"
        ).format(
            py=py_cmd.replace("'", "''"),
            root=str(install_root).replace("'", "''"),
            res=str(result_path).replace("'", "''"),
            helper=str(helper).replace("'", "''"),
            wd=str(work_dir).replace("'", "''"),
            trace=str(helper_trace_path).replace("'", "''"),
        )
    else:
        _append_update_log(
            install_root,
            "elevate: updater_exe_missing path={p} env={e}=1 allows dev python fallback only".format(
                p=updater_exe,
                e=ENV_ALLOW_PYTHON_ELEVATE,
            ),
        )
        try:
            job_path.unlink(missing_ok=True)
        except OSError:
            pass
        return False, "updater_exe_missing"

    _append_update_log(install_root, f"elevate: ps_command={ps}")
    try:
        cp = subprocess.run(
            ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", ps],
            capture_output=True,
            text=True,
            check=False,
        )
        try:
            ps_out_path.write_text(cp.stdout or "", encoding="utf-8")
            ps_err_path.write_text(cp.stderr or "", encoding="utf-8")
        except Exception:
            pass
        out = (cp.stdout or "").strip().replace("\r", "\\r").replace("\n", "\\n")
        err = (cp.stderr or "").strip().replace("\r", "\\r").replace("\n", "\\n")
        _append_update_log(
            install_root,
            "elevate: powershell returncode={rc} stdout={out} stderr={err} result_file={res} ps_stdout_file={pso} ps_stderr_file={pse} uid={uid}".format(
                rc=cp.returncode,
                out=out[:400] if out else "-",
                err=err[:400] if err else "-",
                res=result_path,
                pso=ps_out_path,
                pse=ps_err_path,
                uid=uid,
            ),
        )
        if cp.stdout:
            _append_update_log(
                install_root,
                "elevate: powershell_stdout_full uid={uid} data={data}".format(
                    uid=uid,
                    data=cp.stdout.strip().replace("\r", "\\r").replace("\n", "\\n"),
                ),
            )
        if cp.stderr:
            _append_update_log(
                install_root,
                "elevate: powershell_stderr_full uid={uid} data={data}".format(
                    uid=uid,
                    data=cp.stderr.strip().replace("\r", "\\r").replace("\n", "\\n"),
                ),
            )
        if cp.returncode != 0 and not result_path.is_file():
            _append_update_log(
                install_root,
                "elevate: diagnostics uid={uid} returncode={rc} result_exists={re} helper_exists={he} "
                "launcher_exists={pe} wd_exists={we} trace_exists={te} job_exists={je}".format(
                    uid=uid,
                    rc=cp.returncode,
                    re=result_path.is_file(),
                    he=bool(helper and helper.is_file()),
                    pe=Path(py_cmd).is_file(),
                    we=work_dir.is_dir(),
                    te=helper_trace_path.is_file(),
                    je=job_path.is_file(),
                ),
            )
            if helper_trace_path.is_file():
                try:
                    trace_text = helper_trace_path.read_text(encoding="utf-8", errors="replace")
                    _append_update_log(
                        install_root,
                        "elevate: helper_trace_full uid={uid} data={data}".format(
                            uid=uid,
                            data=trace_text.strip().replace("\r", "\\r").replace("\n", "\\n"),
                        ),
                    )
                except Exception:
                    pass
            low = ((cp.stdout or "") + "\n" + (cp.stderr or "")).lower()
            if "operation was canceled" in low or "cancel" in low or "中止" in low:
                return False, "uac_cancelled"
            if "__elevate_pserr" in low:
                return False, "uac_ps_error"
            if "__elevate_exit:2" in low:
                return False, "uac_child_exit_2"
            if "parameter set" in low or "パラメーター セットを解決できません" in low:
                return False, "ps_param_set_error"
            if "can't open file" in low or "no such file" in low or "not recognized" in low:
                return False, "uac_start_failed_path"
            if "blocked" in low or "group policy" in low or "disabled" in low:
                return False, "uac_start_failed_policy"
            return False, "uac_child_result_missing"
        if result_path.is_file():
            try:
                obj = json.loads(result_path.read_text(encoding="utf-8-sig"))
                if isinstance(obj, dict) and bool(obj.get("ok", False)):
                    return True, "ok"
                echild = str((obj or {}).get("error") or "elevated_apply_failed")
                _append_update_log(
                    install_root,
                    "elevate: child_result ok=false error={e}".format(e=echild),
                )
                return False, "elevated_child_failed"
            except Exception as e:
                _append_update_log(
                    install_root,
                    "elevate: result_parse_failed err={t}: {m}".format(
                        t=type(e).__name__,
                        m=e,
                    ),
                )
                return False, "result_parse_failed"
        if cp.returncode == 0:
            return True, "ok_no_result"
        if cp.returncode == 2:
            return False, "uac_child_exit_2"
        return False, "uac_start_failed"
    except Exception as e:
        _append_update_log(
            install_root,
            "elevate: subprocess_failed err={t}: {m}".format(
                t=type(e).__name__,
                m=e,
            ),
        )
        return False, "uac_start_failed"
    finally:
        if helper is not None:
            try:
                helper.unlink(missing_ok=True)
            except OSError:
                pass
        try:
            result_path.unlink(missing_ok=True)
        except OSError:
            pass
        try:
            helper_trace_path.unlink(missing_ok=True)
        except OSError:
            pass
        try:
            job_path.unlink(missing_ok=True)
        except OSError:
            pass


def _read_and_clear_bin_apply_success_marker(install_root: Path | None) -> dict[str, Any] | None:
    path = _bin_apply_success_marker_path(install_root)
    try:
        if not path.is_file():
            return None
        raw = path.read_text(encoding="utf-8-sig")
        obj = json.loads(raw)
        got = obj if isinstance(obj, dict) else {}
    except Exception:
        got = {}
    try:
        path.unlink(missing_ok=True)
    except Exception:
        pass
    return got


def _notify_pending_bin_apply_success(*, owner_hwnd: int | None = None, sheet_id: str = "") -> None:
    root = _install_root()
    meta = _read_and_clear_bin_apply_success_marker(root)
    if not isinstance(meta, dict):
        return
    target_bin = str(meta.get("target_bin_version") or "").strip()
    display_version = str(meta.get("display_version") or "").strip()
    log_path = str(meta.get("log_path") or _update_log_path(root)).strip()
    target_hint = ""
    if target_bin:
        target_hint = f"\n\n適用後 bin 版: {target_bin}"
    elif display_version:
        target_hint = f"\n\n表示版: {display_version}"
    _message_box(
        _um(
            "BIN_APPLY_SUCCESS_NOTIFY_TEMPLATE",
            "前回予約した bin 更新が完了しています。{target_hint}\n\n"
            "Excel を再起動して利用してください。\n\nログ: {log_path}",
            target_hint=target_hint,
            log_path=log_path,
        ),
        title=_update_check_title(),
        owner_hwnd=owner_hwnd,
        sheet_id=sheet_id,
    )


def _wait_ui_dispatch_result(result_path: Path, timeout_sec: float = 120.0) -> dict[str, Any] | None:
    deadline = time.time() + max(0.1, float(timeout_sec))
    while time.time() < deadline:
        try:
            if result_path.exists() and result_path.stat().st_size > 0:
                from ui_qt.ipc_file import read_pickle

                got = read_pickle(result_path)
                return got if isinstance(got, dict) else None
        except Exception:
            pass
        time.sleep(0.05)
    return None


def _show_update_dialog_via_ui_server(
    *,
    req_dict: dict[str, Any],
    owner_hwnd: int,
    sheet_id: str = "",
    timeout_sec: float = 120.0,
) -> dict[str, Any] | None:
    try:
        from svc.svc_host import ensure_ui_server
        from ui_qt.ipc_file import get_ipc_root, get_request_dir, write_pickle
    except Exception:
        return None

    try:
        ensure_ui_server()
    except Exception as e:
        logger.warning("packaged_update ensure_ui_server failed: %s", e)
        return None

    ts_ms = int(time.time() * 1000)
    pid = os.getpid()
    res_dir = Path(get_ipc_root()) / "result"
    res_dir.mkdir(parents=True, exist_ok=True)
    result_path = res_dir / f"res_update_check_{ts_ms}_{pid}.pkl"
    payload = {
        "parent_hwnd": int(owner_hwnd or 0),
        "result_path": str(result_path),
        "ready_path": "",
        "sheet_id": str(sheet_id or "_"),
        "log_path": "",
        "action": "update_check",
        "module": "ui_qt.ui_update_check",
        "req_dict": req_dict,
    }
    req_path = get_request_dir() / f"req_update_check_{ts_ms}_{pid}.pkl"
    try:
        write_pickle(req_path, payload)
        return _wait_ui_dispatch_result(result_path, timeout_sec=timeout_sec)
    except Exception as e:
        logger.warning("packaged_update UI request failed: %s", e)
        return None
    finally:
        try:
            result_path.unlink(missing_ok=True)
        except Exception:
            pass


def _message_box(
    text: str,
    *,
    title: str | None = None,
    style: int = MB_OK | MB_ICONINFORMATION,
    owner_hwnd: int | None = None,
    sheet_id: str = "",
) -> int:
    if os.name != "nt":
        logger.info("packaged_update msg (no GUI): %s", text.replace("\n", " "))
        return 0
    eff_title = str(title) if title is not None else _update_check_title()
    hwnd = int(owner_hwnd or 0)
    if hwnd < 0:
        hwnd = 0
    # 優先経路: ui_server + Qt ダイアログ（Excel 前面・中央の安定表示）
    if hwnd:
        is_confirm = bool(int(style) & MB_YESNO)
        is_warning = bool(int(style) & MB_ICONWARNING)
        req_action = "update_check_confirm" if is_confirm else ("update_check_warning" if is_warning else "update_check_done")
        got = _show_update_dialog_via_ui_server(
            req_dict={
                "action": req_action,
                "title": eff_title,
                "message": str(text or ""),
            },
            owner_hwnd=hwnd,
            sheet_id=sheet_id,
        )
        if isinstance(got, dict):
            if is_confirm:
                button = str(got.get("button") or "").strip().lower()
                rc = int(got.get("rc", 0) or 0)
                if button == "yes" or rc == 1:
                    return IDYES
                if button == "no":
                    return IDNO
                return 0
            return int(got.get("rc", 1) or 1)
        _append_update_log(_install_root(), "ui_update_dialog=fallback_to_win32")

    # フォールバック: Win32 MessageBox
    try:
        import ctypes

        if hwnd:
            # Excel を先に前面化して owner 中央表示の成功率を上げる（失敗時も続行）。
            try:
                from core import core_w32

                root = int(core_w32.get_root_window(hwnd) or hwnd)
                if root > 0:
                    try:
                        core_w32.nudge_top_level_to_foreground(root)
                    except Exception:
                        pass
                    try:
                        core_w32.set_foreground_window_attach_input(root)
                    except Exception:
                        pass
                    try:
                        core_w32.bring_to_front(root)
                    except Exception:
                        pass
                    hwnd = root
            except Exception:
                pass
        msg_style = int(style) | MB_SETFOREGROUND
        return int(ctypes.windll.user32.MessageBoxW(hwnd, text, eff_title, msg_style))
    except Exception:
        logger.warning("packaged_update MessageBox failed; text=%s", text[:200])
        return 0


def _read_version_file(path: Path) -> str | None:
    try:
        if not path.is_file():
            return None
        lines = path.read_text(encoding="utf-8-sig").splitlines()
        if not lines:
            return None
        v = lines[0].strip()
        return v or None
    except OSError:
        return None


def read_installed_bin_version(install_root: Path) -> str | None:
    return _read_version_file(install_root / "VERSION.txt")


def read_installed_config_version(install_root: Path) -> str | None:
    return _read_version_file(install_root / "config" / "VERSION.txt")


def read_installed_bootstrap_version(install_root: Path) -> str | None:
    return _read_version_file(install_root / "bootstrap" / "VERSION.txt")


def _resolve_catalog_file(install_root: Path) -> Path | None:
    override = (os.environ.get(ENV_CATALOG_PATH) or "").strip()
    if override:
        p = Path(override)
        return p if p.is_file() else None

    admin = install_root / ADMIN_CATALOG_REL
    try:
        if admin.is_file():
            first = admin.read_text(encoding="utf-8-sig").splitlines()
            if first:
                p = Path(first[0].strip())
                if p.is_file():
                    return p
    except OSError:
        pass

    root = (os.environ.get(ENV_DEPLOY_ROOT) or "").strip()
    if not root:
        return None
    cand = Path(root) / "catalog.json"
    return cand if cand.is_file() else None


def load_catalog(path: Path) -> dict[str, Any] | None:
    try:
        raw = path.read_text(encoding="utf-8-sig")
        obj = json.loads(raw)
        return obj if isinstance(obj, dict) else None
    except (OSError, json.JSONDecodeError) as e:
        logger.info("packaged_update catalog read failed: %s", e)
        return None


def _version_is_newer(installed: str, latest: str) -> bool:
    try:
        return Version(latest) > Version(installed)
    except InvalidVersion:
        logger.info("packaged_update non-semver compare skipped: %r %r", installed, latest)
        return False


def _parse_bin_triplet(raw: str | None) -> tuple[int, int, int] | None:
    txt = str(raw or "").strip()
    m = re.fullmatch(r"(\d+)\.(\d+)\.(\d+)", txt)
    if not m:
        return None
    return int(m.group(1)), int(m.group(2)), int(m.group(3))


def _parse_config_revision(raw: str | None) -> int | None:
    txt = str(raw or "").strip()
    if not re.fullmatch(r"\d+", txt):
        return None
    return int(txt)


def _parse_bootstrap_triplet(raw: str | None) -> tuple[int, int, int] | None:
    txt = str(raw or "").strip()
    m = re.fullmatch(r"(\d+)\.(\d+)\.(\d+)", txt)
    if not m:
        return None
    return int(m.group(1)), int(m.group(2)), int(m.group(3))


def needs_bootstrap_update_bool(installed_bootstrap: str | None, latest_bootstrap: str | None) -> bool:
    """catalog.bootstrap.latest_version とインストール済み bootstrap 版の比較（check_for_updates と同じ判定）。"""
    if not latest_bootstrap or not str(latest_bootstrap).strip():
        return False
    lb = str(latest_bootstrap).strip()
    i_bt = _parse_bootstrap_triplet(installed_bootstrap)
    l_bt = _parse_bootstrap_triplet(lb)
    if i_bt is not None and l_bt is not None:
        return l_bt > i_bt
    if not installed_bootstrap or not str(installed_bootstrap).strip():
        return True
    return False


def _bin_is_newer(installed: str, latest: str) -> bool:
    a = _parse_bin_triplet(installed)
    b = _parse_bin_triplet(latest)
    if a is None or b is None:
        return _version_is_newer(installed, latest)
    return b > a


def _bin_gte(a_raw: str, b_raw: str) -> bool:
    a = _parse_bin_triplet(a_raw)
    b = _parse_bin_triplet(b_raw)
    if a is None or b is None:
        return _version_gte(a_raw, b_raw)
    return a >= b


def _config_is_newer(installed: str | None, latest: str | None) -> bool:
    li = _parse_config_revision(installed)
    ll = _parse_config_revision(latest)
    if ll is None and "." in str(latest or ""):
        return _version_is_newer(str(installed or ""), str(latest or ""))
    if ll is None:
        return False
    if li is None:
        return True
    return ll > li


def _version_gte(a: str, b: str) -> bool:
    try:
        return Version(a) >= Version(b)
    except InvalidVersion:
        return False


def _catalog_bin_latest(data: dict[str, Any]) -> str | None:
    b = data.get("bin")
    if isinstance(b, dict):
        v = str((b.get("latest_version") or "").strip())
        return v or None
    v = str((data.get("latest_version") or data.get("latest") or "").strip())
    return v or None


def _catalog_config_latest(data: dict[str, Any]) -> str | None:
    c = data.get("config")
    if not isinstance(c, dict):
        return None
    v = str((c.get("latest_version") or "").strip())
    return v or None


def _catalog_set_version(data: dict[str, Any]) -> str | None:
    v = str((data.get("set_version") or "").strip())
    return v or None


def _catalog_bootstrap_latest(data: dict[str, Any]) -> str | None:
    b = data.get("bootstrap")
    if not isinstance(b, dict):
        return None
    v = str((b.get("latest_version") or "").strip())
    return v or None


def _normalize_set_version_text(raw: str | None) -> str | None:
    txt = str(raw or "").strip()
    m = re.fullmatch(r"(\d+)\.(\d+)\.(\d+)\.(\d+)", txt)
    if not m:
        return None
    return f"{int(m.group(1))}.{int(m.group(2))}.{int(m.group(3))}.{int(m.group(4))}"


def _compose_set_version(bin_version: str | None, config_version: str | None) -> str | None:
    b = _parse_bin_triplet(bin_version)
    c = _parse_config_revision(config_version)
    if b is None or c is None:
        return None
    return f"{b[0]}.{b[1]}.{b[2]}.{c}"


def _normalize_display_version(version_text: str | None) -> str | None:
    """DisplayVersion uses major.minor.patch (ignore 4th+ segment)."""
    if not version_text:
        return None
    try:
        release = list(Version(version_text).release)
    except InvalidVersion:
        parts = [p.strip() for p in str(version_text).split(".") if p.strip()]
        return ".".join(parts[:3]) if parts else None
    if not release:
        return None
    while len(release) < 3:
        release.append(0)
    return ".".join(str(p) for p in release[:3])


def _try_sync_display_version(display_version: str | None, install_root: Path | None) -> bool:
    val = str(display_version or "").strip()
    if not val:
        if install_root is not None:
            _append_update_log(install_root, "display_version_sync=skip reason=empty")
        return False
    if os.name != "nt":
        return False
    try:
        import winreg
    except Exception:
        return False

    sub = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\{B5E8C4A2-1D3F-4E6B-9C0D-1A2B3C4D5E6F}_is1"
    candidates = [
        (winreg.HKEY_LOCAL_MACHINE, sub),
        (winreg.HKEY_LOCAL_MACHINE, r"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{B5E8C4A2-1D3F-4E6B-9C0D-1A2B3C4D5E6F}_is1"),
        (winreg.HKEY_CURRENT_USER, sub),
        (winreg.HKEY_CURRENT_USER, r"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{B5E8C4A2-1D3F-4E6B-9C0D-1A2B3C4D5E6F}_is1"),
    ]
    fail_counts: dict[str, int] = {
        "hklm_access_denied": 0,
        "hklm_not_found": 0,
        "hklm_other_oserror": 0,
        "hkcu_access_denied": 0,
        "hkcu_not_found": 0,
        "hkcu_other_oserror": 0,
    }
    for root, subkey in candidates:
        root_name = "HKLM" if root == winreg.HKEY_LOCAL_MACHINE else "HKCU"
        if install_root is not None:
            _append_update_log(
                install_root,
                f"display_version_sync=try root={root_name} key={subkey} value={val}",
            )
        try:
            with winreg.OpenKey(root, subkey, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, "CSV Tool")
                winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, val)
            if install_root is not None:
                _append_update_log(install_root, f"display_version_sync=ok value={val} root={root_name} key={subkey}")
            return True
        except OSError as e:
            winerr = int(getattr(e, "winerror", 0) or 0)
            err_no = int(getattr(e, "errno", 0) or 0)
            reason = "other_oserror"
            if winerr in (5,) or err_no in (13,):
                reason = "access_denied"
            elif winerr in (2, 3) or err_no in (2,):
                reason = "key_not_found"
            bucket = f"{root_name.lower()}_{'not_found' if reason == 'key_not_found' else reason}"
            if bucket in fail_counts:
                fail_counts[bucket] += 1
            if install_root is not None:
                _append_update_log(
                    install_root,
                    "display_version_sync=try_failed root={r} key={k} reason={rsn} "
                    "winerror={we} errno={en} err_type={et} err={em}".format(
                        r=root_name,
                        k=subkey,
                        rsn=reason,
                        we=winerr or "-",
                        en=err_no or "-",
                        et=type(e).__name__,
                        em=e,
                    ),
                )
            continue
    for subkey in (
        sub,
        r"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{B5E8C4A2-1D3F-4E6B-9C0D-1A2B3C4D5E6F}_is1",
    ):
        try:
            with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, subkey, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, "CSV Tool")
                winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, val)
            if install_root is not None:
                _append_update_log(
                    install_root,
                    f"display_version_sync=ok value={val} root=HKCU_created key={subkey}",
                )
            return True
        except OSError as e:
            if install_root is not None:
                _append_update_log(
                    install_root,
                    "display_version_sync=hkcu_create_failed key={k} err_type={et} err={em}".format(
                        k=subkey,
                        et=type(e).__name__,
                        em=e,
                    ),
                )
            continue
    if install_root is not None:
        _append_update_log(
            install_root,
            "display_version_sync=skip reason=key_not_found_or_not_writable value={v} "
            "hklm_access_denied={ha} hklm_not_found={hn} hklm_other_oserror={ho} "
            "hkcu_access_denied={ca} hkcu_not_found={cn} hkcu_other_oserror={co}".format(
                v=val,
                ha=fail_counts["hklm_access_denied"],
                hn=fail_counts["hklm_not_found"],
                ho=fail_counts["hklm_other_oserror"],
                ca=fail_counts["hkcu_access_denied"],
                cn=fail_counts["hkcu_not_found"],
                co=fail_counts["hkcu_other_oserror"],
            ),
        )
    return False


def sync_uninstall_display_version_from_catalog(
    catalog_path: str | Path,
    install_root: Path,
    *,
    log_prefix: str = "post_bin_apply",
) -> bool:
    """
    bin 更新（patch/full）適用後に、catalog から 4 桁セット版を求め Windows の
    Uninstall DisplayVersion をベストエフォートで同期する。
    """
    p = Path(str(catalog_path).strip())
    if not p.is_file():
        _append_update_log(
            install_root,
            f"{log_prefix} display_version_sync=skip reason=no_catalog path={p}",
        )
        return False
    data = load_catalog(p)
    if not isinstance(data, dict):
        _append_update_log(
            install_root,
            f"{log_prefix} display_version_sync=skip reason=catalog_unreadable path={p}",
        )
        return False
    set_v = _catalog_set_version(data)
    latest_bin = _catalog_bin_latest(data)
    latest_cfg = _catalog_config_latest(data)
    display_version = _normalize_set_version_text(set_v) or _compose_set_version(latest_bin, latest_cfg)
    if not display_version:
        _append_update_log(
            install_root,
            f"{log_prefix} display_version_sync=skip reason=empty_computed "
            f"set={set_v!r} bin={latest_bin!r} cfg={latest_cfg!r}",
        )
        return False
    return _try_sync_display_version(display_version, install_root)


def _catalog_resolve_payload(catalog_path: Path, relative_path: str) -> Path:
    p = Path(relative_path)
    if p.is_absolute():
        return p
    return (catalog_path.parent / p).resolve()


def _copy_payload_to_local(local_path: Path, source_path: Path) -> None:
    local_path.parent.mkdir(parents=True, exist_ok=True)
    tmp = local_path.with_suffix(local_path.suffix + ".new")
    shutil.copy2(source_path, tmp)
    os.replace(tmp, local_path)


def _queue_pending_bin_update(st: dict[str, Any], *, source: str, require_admin: bool = False) -> tuple[bool, str]:
    root = _install_root()
    if root is None:
        return False, "HC_INSTALL_ROOT not set or not a directory"
    if not root.is_dir():
        return False, "install root not found"
    target = str(st.get("latest_bin_version") or "").strip()
    if not target:
        return False, "latest bin version missing"
    catalog_path = str(st.get("catalog_path") or "").strip()
    patch_zip = str(st.get("bin_zip_path") or "").strip()
    if not patch_zip:
        return False, "bin zip path missing"
    patch_mode = str(st.get("bin_apply_mode") or "patch").strip().lower()
    if patch_mode not in ("patch", "full"):
        patch_mode = "patch"
    patch_src = Path(patch_zip)
    if not patch_src.is_file():
        return False, f"payload not found: {patch_src}"
    require_admin_final = bool(require_admin)
    root_text = str(root).lower()
    if (
        os.name == "nt"
        and ("\\program files\\" in root_text or root_text.endswith("\\program files"))
        and not require_admin_final
    ):
        require_admin_final = True
        _append_update_log(
            root,
            "apply_bin: require_admin_forced reason=program_files_install root={r} source={s}".format(
                r=root,
                s=source,
            ),
        )
    else:
        _append_update_log(
            root,
            "apply_bin: require_admin_effective selected={sel} effective={eff} source={s}".format(
                sel=bool(require_admin),
                eff=require_admin_final,
                s=source,
            ),
        )

    paths = build_paths(root)
    paths.update_root.mkdir(parents=True, exist_ok=True)
    paths.payload_root.mkdir(parents=True, exist_ok=True)
    patch_local = paths.payload_root / "patch.zip"
    _copy_payload_to_local(patch_local, patch_src)

    full_local = paths.payload_root / "full.zip"
    full_src_s = str(st.get("bin_full_zip_path") or "").strip()
    full_sha = str(st.get("bin_full_zip_sha256_expected") or "").strip().lower()
    full_obj: dict[str, Any] = {
        "relative_path": full_src_s,
        "sha256": full_sha,
        "local_path": str(full_local),
        "downloaded": False,
    }
    if full_src_s and Path(full_src_s).is_file() and patch_mode == "full":
        _copy_payload_to_local(full_local, Path(full_src_s))
        full_obj["downloaded"] = True

    bootstrap_obj: dict[str, Any] = {
        "target_version": "",
        "local_new_path": "",
        "pending_swap": False,
    }
    apply_scope_bin = "bin_only"
    cat_path = str(st.get("catalog_path") or "").strip()
    if cat_path and Path(cat_path).is_file():
        data = load_catalog(Path(cat_path))
        if isinstance(data, dict):
            bzp, _bsha, berr = _prepare_bootstrap_full_apply(data, Path(cat_path))
            if bzp is not None and bzp.is_file():
                latest_bt = str(_catalog_bootstrap_latest(data) or "").strip()
                installed_bt = read_installed_bootstrap_version(root)
                if not needs_bootstrap_update_bool(installed_bt, latest_bt or None):
                    _append_update_log(
                        root,
                        "bootstrap queue skipped reason=already_latest "
                        f"installed={installed_bt or '-'} latest={latest_bt or '-'}",
                    )
                else:
                    local_new = paths.payload_root / "bootstrap.new"
                    try:
                        _copy_payload_to_local(local_new, bzp)
                        bootstrap_obj["target_version"] = str(
                            _catalog_bootstrap_latest(data) or data.get("set_version") or ""
                        ).strip()
                        bootstrap_obj["local_new_path"] = str(local_new)
                        bootstrap_obj["pending_swap"] = True
                        apply_scope_bin = "bin+bootstrap"
                    except OSError as e:
                        _append_update_log(root, f"bootstrap queue copy failed err={type(e).__name__}: {e}")
            elif berr:
                _append_update_log(root, f"bootstrap queue skipped err={berr}")

    pending = {
        "schema_version": 2,
        "apply_scope": apply_scope_bin,
        "require_admin": require_admin_final,
        "state": "downloaded",
        "target_bin_version": target,
        "mode": patch_mode,
        "catalog_path": catalog_path,
        "catalog_checked_at": time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime()),
        "patch": {
            "relative_path": patch_zip,
            "sha256": str(st.get("bin_zip_sha256_expected") or "").strip().lower(),
            "local_path": str(patch_local),
        },
        "full": full_obj,
        "bootstrap": bootstrap_obj,
        "retry": {
            "patch_retry_in_run": 0,
            "bootstrap_retry_in_run": 0,
            "patch_fail_total": 0,
            "full_fail_total": 0,
            "last_error_code": "",
            "last_error_message": "",
            "last_failed_at": "",
        },
        "timestamps": {
            "created_at": time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime()),
            "updated_at": time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime()),
        },
        "source": source,
    }
    write_pending(paths, pending)
    return True, str(paths.pending_path)


def _queue_pending_bootstrap_only(st: dict[str, Any], *, source: str, require_admin: bool = False) -> tuple[bool, str]:
    """bootstrap のみ更新（bin は対象外）。catalog.bootstrap.full を payload にコピーする。"""
    root = _install_root()
    if root is None:
        return False, "HC_INSTALL_ROOT not set or not a directory"
    if not root.is_dir():
        return False, "install root not found"
    catalog_path = str(st.get("catalog_path") or "").strip()
    if not catalog_path or not Path(catalog_path).is_file():
        return False, "catalog path missing"
    data = load_catalog(Path(catalog_path))
    if not isinstance(data, dict):
        return False, "catalog not readable"
    bzp, _bsha, berr = _prepare_bootstrap_full_apply(data, Path(catalog_path))
    if bzp is None or not bzp.is_file():
        return False, berr or "bootstrap full payload missing"

    paths = build_paths(root)
    paths.update_root.mkdir(parents=True, exist_ok=True)
    paths.payload_root.mkdir(parents=True, exist_ok=True)
    placeholder_patch = paths.payload_root / "patch.zip"
    try:
        if not placeholder_patch.is_file():
            placeholder_patch.write_bytes(b"")
    except OSError:
        pass

    local_new = paths.payload_root / "bootstrap.new"
    try:
        _copy_payload_to_local(local_new, bzp)
    except OSError as e:
        return False, f"bootstrap copy failed: {e}"

    bootstrap_obj: dict[str, Any] = {
        "target_version": str(_catalog_bootstrap_latest(data) or data.get("set_version") or "").strip(),
        "local_new_path": str(local_new),
        "pending_swap": True,
    }
    pending: dict[str, Any] = {
        "schema_version": 2,
        "apply_scope": "bootstrap_only",
        "require_admin": bool(require_admin),
        "state": "downloaded",
        "target_bin_version": "",
        "mode": "patch",
        "catalog_path": catalog_path,
        "catalog_checked_at": time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime()),
        "patch": {
            "relative_path": "",
            "sha256": "",
            "local_path": str(placeholder_patch),
        },
        "full": {
            "relative_path": "",
            "sha256": "",
            "local_path": "",
            "downloaded": False,
        },
        "bootstrap": bootstrap_obj,
        "retry": {
            "patch_retry_in_run": 0,
            "bootstrap_retry_in_run": 0,
            "patch_fail_total": 0,
            "full_fail_total": 0,
            "last_error_code": "",
            "last_error_message": "",
            "last_failed_at": "",
        },
        "timestamps": {
            "created_at": time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime()),
            "updated_at": time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime()),
        },
        "source": source,
    }
    write_pending(paths, pending)
    return True, str(paths.pending_path)


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().lower()


def _copy_tree_atomic(src_root: Path, dst_root: Path) -> None:
    dst_root.mkdir(parents=True, exist_ok=True)
    for s in src_root.rglob("*"):
        rel = s.relative_to(src_root)
        d = dst_root / rel
        if s.is_dir():
            d.mkdir(parents=True, exist_ok=True)
            continue
        d.parent.mkdir(parents=True, exist_ok=True)
        tmp = d.with_name(d.name + ".new")
        shutil.copy2(s, tmp)
        os.replace(tmp, d)


def _try_apply_config_update(
    *,
    install_root: Path,
    catalog: dict[str, Any],
    catalog_path: Path,
    installed_bin: str,
    installed_cfg: str | None,
) -> tuple[bool, str | None, str | None]:
    """
    Returns: (applied, new_config_version, error_message)
    """
    cfg = catalog.get("config")
    if not isinstance(cfg, dict):
        return False, installed_cfg, None

    latest_cfg = str((cfg.get("latest_version") or "").strip())
    if not latest_cfg:
        return False, installed_cfg, None

    if installed_cfg and not _version_is_newer(installed_cfg, latest_cfg):
        return False, installed_cfg, None
    if not installed_cfg and not latest_cfg:
        return False, installed_cfg, None

    min_bin = str((cfg.get("min_bin_version") or "").strip())
    if min_bin and not _bin_gte(installed_bin, min_bin):
        return False, installed_cfg, None

    payload = cfg.get("payload")
    if not isinstance(payload, dict):
        return False, installed_cfg, "config.payload missing"

    rel = str((payload.get("relative_path") or "").strip())
    if not rel:
        return False, installed_cfg, "config.payload.relative_path missing"
    zip_path = _catalog_resolve_payload(catalog_path, rel)
    if not zip_path.is_file():
        return False, installed_cfg, f"config payload not found: {zip_path}"

    expected_sha = str((payload.get("sha256") or "").strip()).lower()
    if expected_sha:
        actual_sha = _sha256_file(zip_path)
        if actual_sha != expected_sha:
            return False, installed_cfg, "config payload sha256 mismatch"

    target_cfg = install_root / "config"
    with tempfile.TemporaryDirectory(prefix="csv_tool_cfg_", dir=str(install_root)) as td:
        temp_root = Path(td)
        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                zf.extractall(temp_root)
        except Exception as e:
            return False, installed_cfg, f"config zip extract failed: {type(e).__name__}: {e}"

        src_cfg = temp_root / "config"
        if not src_cfg.is_dir():
            # Accept zips where the root itself is the config payload
            src_cfg = temp_root
        try:
            _copy_tree_atomic(src_cfg, target_cfg)
            (target_cfg / "VERSION.txt").write_text(latest_cfg + "\n", encoding="utf-8")
        except Exception as e:
            return False, installed_cfg, f"config apply failed: {type(e).__name__}: {e}"

    return True, latest_cfg, None


def _prepare_bin_full_apply(
    catalog: dict[str, Any],
    catalog_path: Path,
) -> tuple[Path | None, str | None, str | None]:
    """
    catalog.bin.full の zip を解決し、sha256 が catalog と一致するか検証する。
    戻り値: (zip_path, expected_sha_or_none, error_message_or_none)
    """
    b = catalog.get("bin")
    if not isinstance(b, dict):
        return None, None, "catalog.bin がありません"
    full = b.get("full")
    if not isinstance(full, dict):
        return None, None, "catalog.bin.full がありません（自動適用はフル zip のみ）"
    rel = str((full.get("relative_path") or "").strip())
    if not rel:
        return None, None, "catalog.bin.full.relative_path がありません"
    zp = _catalog_resolve_payload(catalog_path, rel)
    if not zp.is_file():
        return None, None, f"bin フル zip が見つかりません: {zp}"
    expected_sha = str((full.get("sha256") or "").strip()).lower()
    if expected_sha:
        actual = _sha256_file(zp)
        if actual != expected_sha:
            return None, None, "bin フル zip の sha256 が catalog と一致しません"
    return zp, expected_sha or None, None


def _patch_meta_eligible(installed_bin: str, patch: dict[str, Any]) -> bool:
    """True if installed version is within catalog.bin.patch from_* range (SemVer)."""
    lo = str((patch.get("from_min_version") or "")).strip()
    hi = str((patch.get("from_max_version") or "")).strip()
    if not lo and not hi:
        return False
    if lo and not _bin_gte(installed_bin, lo):
        return False
    if hi and not _bin_gte(hi, installed_bin):
        return False
    return True


def _prepare_bin_apply(
    catalog: dict[str, Any],
    catalog_path: Path,
    installed_bin: str,
    log_root: Path | None,
) -> tuple[str, Path | None, str | None, str | None]:
    """
    Resolve bin zip for user-confirmed apply: prefer catalog.bin.patch when eligible, else bin.full.

    Returns: (mode, zip_path, expected_sha_or_none, error_or_none)
    mode is 'patch' or 'full'; on total failure mode is '' and zip None.
    """
    b = catalog.get("bin")
    if not isinstance(b, dict):
        return "", None, None, "catalog.bin がありません"

    def log(line: str) -> None:
        if log_root is not None:
            _append_update_log(log_root, line)

    patch = b.get("patch")
    if isinstance(patch, dict):
        rel_p = str((patch.get("relative_path") or "").strip())
        if rel_p and _patch_meta_eligible(installed_bin, patch):
            zp = _catalog_resolve_payload(catalog_path, rel_p)
            if zp.is_file():
                expected_sha = str((patch.get("sha256") or "").strip()).lower()
                if expected_sha:
                    actual = _sha256_file(zp)
                    if actual != expected_sha:
                        log(
                            "bin_apply: patch sha256 mismatch catalog!=zip try_full "
                            f"installed={installed_bin} zip={zp}"
                        )
                    else:
                        log(f"bin_apply: prepared patch zip={zp} installed={installed_bin}")
                        return "patch", zp, expected_sha or None, None
                else:
                    log(f"bin_apply: prepared patch (no sha in catalog) zip={zp} installed={installed_bin}")
                    return "patch", zp, None, None
            else:
                log(f"bin_apply: patch zip missing try_full path={zp}")
        elif rel_p:
            log(
                "bin_apply: patch not eligible "
                f"installed={installed_bin} from_min={patch.get('from_min_version')!s} "
                f"from_max={patch.get('from_max_version')!s}"
            )

    zp2, zsha2, err2 = _prepare_bin_full_apply(catalog, catalog_path)
    if err2:
        return "", None, None, err2
    log(f"bin_apply: prepared full zip={zp2}")
    return "full", zp2, zsha2, None


def _prepare_bootstrap_full_apply(
    catalog: dict[str, Any],
    catalog_path: Path,
) -> tuple[Path | None, str | None, str | None]:
    b = catalog.get("bootstrap")
    if not isinstance(b, dict):
        return None, None, "catalog.bootstrap がありません"
    full = b.get("full")
    if not isinstance(full, dict):
        return None, None, "catalog.bootstrap.full がありません"
    rel = str((full.get("relative_path") or "").strip())
    if not rel:
        return None, None, "catalog.bootstrap.full.relative_path がありません"
    zp = _catalog_resolve_payload(catalog_path, rel)
    if not zp.is_file():
        return None, None, f"bootstrap フル zip が見つかりません: {zp}"
    expected_sha = str((full.get("sha256") or "").strip()).lower()
    if expected_sha:
        actual = _sha256_file(zp)
        if actual != expected_sha:
            return None, None, "bootstrap フル zip の sha256 が catalog と一致しません"
    return zp, expected_sha or None, None


def _spawn_bin_apply_worker(
    install_root: Path,
    zip_path: Path,
    expected_sha: str | None,
    display_version: str | None,
    *,
    apply_mode: str = "full",
    target_bin_version: str | None = None,
    cleanup_path: str | None = None,
) -> None:
    log_path = _update_log_path(install_root)
    params = {
        "InstallRoot": str(install_root.resolve()),
        "ZipPath": str(zip_path.resolve()),
        "ExpectedSha": (expected_sha or "").lower(),
        "LogPath": str(log_path.resolve()),
        "DisplayVersion": (display_version or "").strip(),
        "ApplyMode": (apply_mode or "full").strip().lower(),
        "TargetBinVersion": (target_bin_version or "").strip(),
        "CleanupPath": (cleanup_path or "").strip(),
        "NotifyMarkerPath": str(_bin_apply_success_marker_path(install_root).resolve()),
    }
    tmpdir = Path(tempfile.gettempdir())
    uid = uuid.uuid4().hex[:12]
    param_file = tmpdir / f"csv_tool_bin_apply_{os.getpid()}_{uid}.json"
    param_file.write_text(json.dumps(params, ensure_ascii=False), encoding="utf-8")

    def _resolve_updater_launcher() -> list[str] | None:
        # Prefer packaged updater binary under install root.
        ir = runtime_layout.install_root()
        if ir is not None:
            updater_exe = ir / "app" / "bin" / "hc_updater.exe"
            if updater_exe.is_file():
                return [str(updater_exe), "--job", str(param_file)]

        # Dev fallback: run updater python directly with current interpreter.
        try:
            updater_py = Path(__file__).resolve().parents[1] / "hc_updater.py"
            if updater_py.is_file():
                return [sys.executable, str(updater_py), "--job", str(param_file)]
        except Exception:
            pass
        return None

    use_updater_v = os.environ.get(ENV_USE_UPDATER_EXE)
    use_updater = True if use_updater_v is None else core_env.truthy(use_updater_v, empty_means_false=False)
    updater_cmd = _resolve_updater_launcher() if use_updater else None
    script_file = tmpdir / f"csv_tool_bin_worker_{os.getpid()}_{uid}.ps1"
    if updater_cmd:
        cmd = updater_cmd
    else:
        if use_updater:
            raise OSError(
                "hc_updater launcher not found. "
                "Expected {root}\\app\\bin\\hc_updater.exe or core/hc_updater.py. "
                "Set HC_USE_UPDATER_EXE=0 only for controlled legacy fallback."
            )
        # Legacy fallback for environments where updater.exe is not built yet.
        script_file.write_text("\ufeff" + _APPLY_BIN_WORKER_PS1.strip() + "\r\n", encoding="utf-8")
        cmd = [
            "powershell.exe",
            "-NoProfile",
            "-Sta",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            str(script_file),
            "-ParamPath",
            str(param_file),
        ]

    creationflags = 0
    if os.name == "nt":
        creationflags = (
            getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0)
            | getattr(subprocess, "DETACHED_PROCESS", 0)
            | getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
        )
    launcher_kind = "updater_exe_or_py" if updater_cmd else "legacy_ps1"
    _append_update_log(
        install_root,
        "apply_bin: worker launching launcher={lk} mode={m} job={j} cmd0={c0}".format(
            lk=launcher_kind,
            m=(apply_mode or "full").strip().lower(),
            j=param_file,
            c0=(cmd[0] if cmd else "-"),
        ),
    )
    proc = subprocess.Popen(cmd, close_fds=True, creationflags=creationflags)
    _append_update_log(install_root, f"apply_bin: worker launched pid={proc.pid}")


def check_for_updates(
    *,
    source: str,
    notify_offline: bool,
    owner_hwnd: int | None = None,
    sheet_id: str = "",
) -> dict[str, Any]:
    """
    Return status dict.
    Backward-compatible keys:
      ok, needs_update, installed, latest_version, error, catalog_path
    New keys:
      installed_bin, installed_config, latest_bin_version, latest_config_version,
      installed_bootstrap_version, latest_bootstrap_version, needs_bootstrap_update,
      needs_bin_update, needs_config_update, config_update_applied, config_update_error
      bin_zip_path, bin_zip_sha256_expected, bin_update_prepare_error, bin_apply_mode
      bin_full_zip_path, bin_full_zip_sha256_expected
    """
    out: dict[str, Any] = {
        "ok": False,
        "needs_update": False,  # alias of needs_bin_update
        "installed": None,  # alias of installed_bin
        "latest_version": None,  # alias of latest_bin_version
        "error": None,
        "catalog_path": None,
        "installed_bin": None,
        "installed_config": None,
        "installed_bootstrap_version": None,
        "latest_bin_version": None,
        "latest_config_version": None,
        "latest_bootstrap_version": None,
        "display_version": None,
        "needs_bin_update": False,
        "needs_config_update": False,
        "needs_bootstrap_update": False,
        "config_update_applied": False,
        "config_update_error": None,
        "bin_zip_path": None,
        "bin_zip_sha256_expected": None,
        "bin_update_prepare_error": None,
        "bin_apply_mode": None,
        "bin_full_zip_path": None,
        "bin_full_zip_sha256_expected": None,
    }

    root = _install_root()
    if not root:
        out["error"] = "HC_INSTALL_ROOT not set or not a directory"
        _append_update_log(None, f"source={source} error=no_hc_install_root path={_update_log_path(None)}")
        logger.info("packaged_update source=%s error=no_hc_install_root", source)
        return out

    installed_bin = read_installed_bin_version(root)
    installed_cfg = read_installed_config_version(root)
    installed_bootstrap = read_installed_bootstrap_version(root)
    out["installed_bin"] = installed_bin
    out["installed_config"] = installed_cfg
    out["installed_bootstrap_version"] = installed_bootstrap
    out["installed"] = installed_bin

    if not installed_bin:
        out["error"] = "VERSION.txt missing or empty under HC_INSTALL_ROOT"
        _append_update_log(root, f"source={source} error=no_installed_bin_version")
        if notify_offline:
            _message_box(
                "更新確認できませんでした。\nインストール先に VERSION.txt がありません。\n"
                + str(root / "VERSION.txt"),
                owner_hwnd=owner_hwnd,
                sheet_id=sheet_id,
            )
        return out

    cat_path = _resolve_catalog_file(root)
    out["catalog_path"] = str(cat_path) if cat_path else None
    if not cat_path:
        configured = bool((os.environ.get(ENV_DEPLOY_ROOT) or "").strip() or (os.environ.get(ENV_CATALOG_PATH) or "").strip())
        try:
            configured = configured or (root / ADMIN_CATALOG_REL).is_file()
        except OSError:
            pass
        if not configured:
            out["error"] = "skipped_no_catalog_config"
            _append_update_log(root, f"source={source} error=skipped_no_catalog_config")
            return out
        out["error"] = "catalog not found"
        _append_update_log(root, f"source={source} error=no_catalog")
        if notify_offline:
            _message_box(
                "更新確認できませんでした。\n共有の catalog.json にアクセスできないか、"
                "環境変数 HC_DEPLOY_ROOT / HC_CATALOG_PATH が未設定です。\n\n詳細はログ "
                + str(root / UPDATE_LOG_REL)
                + " を参照してください。",
                owner_hwnd=owner_hwnd,
                sheet_id=sheet_id,
            )
        return out

    data = load_catalog(cat_path)
    if not data:
        out["error"] = "catalog.json invalid or unreadable"
        _append_update_log(root, f"source={source} error=bad_catalog path={cat_path}")
        if notify_offline:
            _message_box(
                "更新確認できませんでした。\ncatalog.json を読み取れませんでした。",
                owner_hwnd=owner_hwnd,
                sheet_id=sheet_id,
            )
        return out

    latest_bin = _catalog_bin_latest(data)
    if not latest_bin:
        out["error"] = "catalog has no bin.latest_version"
        _append_update_log(root, f"source={source} error=no_bin_latest path={cat_path}")
        if notify_offline:
            _message_box(
                "更新確認できませんでした。\ncatalog に bin.latest_version がありません。",
                owner_hwnd=owner_hwnd,
                sheet_id=sheet_id,
            )
        return out

    latest_cfg = _catalog_config_latest(data)
    latest_bootstrap = _catalog_bootstrap_latest(data)
    set_version = _catalog_set_version(data)
    out["latest_bin_version"] = latest_bin
    out["latest_config_version"] = latest_cfg
    out["latest_bootstrap_version"] = latest_bootstrap
    out["latest_version"] = latest_bin
    out["ok"] = True
    display_version = _normalize_set_version_text(set_version) or _compose_set_version(latest_bin, latest_cfg)
    out["display_version"] = display_version

    need_bin = _bin_is_newer(installed_bin, latest_bin)
    need_cfg = _config_is_newer(installed_cfg, latest_cfg)
    need_bootstrap = False
    if latest_bootstrap:
        i_bt = _parse_bootstrap_triplet(installed_bootstrap)
        l_bt = _parse_bootstrap_triplet(latest_bootstrap)
        if i_bt is not None and l_bt is not None:
            need_bootstrap = l_bt > i_bt
        elif not installed_bootstrap:
            need_bootstrap = True
    out["needs_bin_update"] = need_bin
    out["needs_config_update"] = need_cfg
    out["needs_bootstrap_update"] = need_bootstrap
    out["needs_update"] = need_bin

    if need_bin:
        full_zp, full_sha, full_err = _prepare_bin_full_apply(data, cat_path)
        out["bin_full_zip_path"] = str(full_zp) if full_zp else None
        out["bin_full_zip_sha256_expected"] = full_sha
        mode, zp, zsha, prep_err = _prepare_bin_apply(data, cat_path, installed_bin, root)
        out["bin_apply_mode"] = mode if zp else None
        out["bin_zip_path"] = str(zp) if zp else None
        out["bin_zip_sha256_expected"] = zsha
        out["bin_update_prepare_error"] = prep_err or full_err

    # config update is silent by policy (failure only notifies).
    if need_cfg:
        applied, new_cfg, cfg_err = _try_apply_config_update(
            install_root=root,
            catalog=data,
            catalog_path=cat_path,
            installed_bin=installed_bin,
            installed_cfg=installed_cfg,
        )
        out["config_update_applied"] = applied
        out["installed_config"] = new_cfg
        out["config_update_error"] = cfg_err
        if cfg_err:
            _append_update_log(root, f"source={source} config_update=failed error={cfg_err}")
        elif applied:
            _append_update_log(root, f"source={source} config_update=applied version={new_cfg}")
            if not _try_sync_display_version(display_version, root):
                _append_update_log(root, "source={s} display_version_sync=skipped value={v}".format(s=source, v=display_version or ""))
        else:
            cfg_meta = data.get("config")
            min_bin = ""
            if isinstance(cfg_meta, dict):
                min_bin = str((cfg_meta.get("min_bin_version") or "").strip())
            if min_bin and not _bin_gte(installed_bin, min_bin):
                _append_update_log(
                    root,
                    "source={s} config_update=skipped reason=min_bin_not_met installed_bin={ib} min_bin={mb}".format(
                        s=source,
                        ib=installed_bin,
                        mb=min_bin,
                    ),
                )

    _append_update_log(
        root,
        "source={s} installed_bin={ib} latest_bin={lb} needs_bin={nb} bin_mode={bm} "
        "installed_cfg={ic} latest_cfg={lc} needs_cfg={nc} "
        "installed_bootstrap={ibt} latest_bootstrap={lbt} needs_bootstrap={nbt} "
        "cfg_applied={ca} cfg_err={ce} set={sv} catalog={cp}".format(
            s=source,
            ib=installed_bin,
            lb=latest_bin,
            nb=need_bin,
            bm=out.get("bin_apply_mode") or "-",
            ic=installed_cfg,
            lc=latest_cfg,
            nc=need_cfg,
            ibt=installed_bootstrap or "-",
            lbt=latest_bootstrap or "-",
            nbt=need_bootstrap,
            ca=out["config_update_applied"],
            ce=out["config_update_error"],
            sv=display_version or "-",
            cp=cat_path,
        ),
    )
    return out


def _show_bin_update_prompt(
    st: dict[str, Any], *, owner_hwnd: int | None = None, sheet_id: str = ""
) -> None:
    prep_err = st.get("bin_update_prepare_error")
    if prep_err:
        lp = str(_update_log_path(_install_root()))
        _message_box(
            _um(
                "BIN_UPDATE_PREP_FAILED_PROMPT_TEMPLATE",
                "bin の新しい版がありますが、自動適用の準備に失敗しました。\n\n"
                + str(prep_err)
                + "\n\n手動でフル配布を適用するか、catalog と zip を確認してください。\nログ: "
                + lp,
                prep_err=str(prep_err),
                log_path=lp,
            ),
            style=MB_OK | MB_ICONWARNING,
            owner_hwnd=owner_hwnd,
            sheet_id=sheet_id,
        )
        return
    zip_s = st.get("bin_zip_path")
    if not zip_s:
        _message_box(
            _um("BIN_UPDATE_ZIP_PATH_MISSING", "bin の zip パスが取得できませんでした。"),
            style=MB_OK | MB_ICONWARNING,
            owner_hwnd=owner_hwnd,
            sheet_id=sheet_id,
        )
        return

    apply_hint = ""
    if (st.get("bin_apply_mode") or "").lower() == "patch":
        apply_hint = _um("BIN_UPDATE_APPLY_HINT_PATCH", "（差分 zip で更新します）\n")
    elif (st.get("bin_apply_mode") or "").lower() == "full":
        apply_hint = _um("BIN_UPDATE_APPLY_HINT_FULL", "")
    msg = _um(
        "BIN_UPDATE_CONFIRM_TEMPLATE",
        "bin の新しい版があります。\n\n"
        "現在 (bin): {installed_bin}\n"
        "現在 (config): {installed_config}\n"
        "現在 (bootstrap): {installed_bootstrap}\n"
        "配布 (bin): {latest_bin}\n"
        "配布 (config): {latest_config}\n"
        "配布 (bootstrap): {latest_bootstrap}\n"
        "表示版: {display_version}\n\n"
        "{apply_hint}"
        "「はい」を選ぶと、すべての Excel を終了したあとに bin（実行ファイル一式）を自動更新します。\n"
        "「いいえ」は後で更新できます。",
        installed_bin=str(st.get("installed_bin") or "-"),
        installed_config=str(st.get("installed_config") or "-"),
        installed_bootstrap=str(st.get("installed_bootstrap_version") or "-"),
        latest_bin=str(st.get("latest_bin_version") or "-"),
        latest_config=str(st.get("latest_config_version") or "-"),
        latest_bootstrap=str(st.get("latest_bootstrap_version") or "-"),
        display_version=str(st.get("display_version") or "-"),
        apply_hint=apply_hint,
    )
    rc = _message_box(
        msg,
        style=MB_YESNO | MB_ICONINFORMATION,
        owner_hwnd=owner_hwnd,
        sheet_id=sheet_id,
    )
    root = _install_root()
    if rc != IDYES:
        if root and root.is_dir():
            _append_update_log(
                root,
                "apply_bin: user_decision=start_no apply_mode_selected={m} target={t}".format(
                    m=str(st.get("bin_apply_mode") or "-"),
                    t=str(st.get("latest_bin_version") or "-"),
                ),
            )
        return

    zp = Path(zip_s)
    if not root or not root.is_dir() or not zp.is_file():
        _message_box(
            _um("BIN_UPDATE_INVALID_ROOT_OR_ZIP", "インストール先または zip が無効です。"),
            style=MB_OK | MB_ICONWARNING,
            owner_hwnd=owner_hwnd,
            sheet_id=sheet_id,
        )
        return
    _append_update_log(
        root,
        "apply_bin: user_decision=start_yes apply_mode_selected={m} target={t}".format(
            m=str(st.get("bin_apply_mode") or "-"),
            t=str(st.get("latest_bin_version") or "-"),
        ),
    )
    rc_admin = _message_box(
        _um(
            "UPDATE_RUN_AS_ADMIN_CONFIRM_TEMPLATE",
            "管理者権限で更新を適用しますか？\n\n"
            "「はい」: 次回起動時に UAC 確認後、管理者権限で更新を適用します。\n"
            "「いいえ」: 通常権限で更新します（Windows の表示版同期に失敗する場合があります）。",
        ),
        style=MB_YESNO | MB_ICONINFORMATION,
        owner_hwnd=owner_hwnd,
        sheet_id=sheet_id,
    )
    require_admin = rc_admin == IDYES
    _append_update_log(root, f"apply_bin: require_admin_selected={require_admin}")

    ok, msg = _queue_pending_bin_update(st, source="interactive_confirm", require_admin=require_admin)
    if not ok:
        _append_update_log(root, f"apply_bin: queue pending failed err={msg}")
        _message_box(
            _um(
                "BIN_WORKER_SPAWN_FAILED_TEMPLATE",
                "更新ワーカーの起動に失敗しました。\n{error}",
                error=str(msg),
            ),
            style=MB_OK | MB_ICONWARNING,
            owner_hwnd=owner_hwnd,
            sheet_id=sheet_id,
        )
        return
    _append_update_log(root, f"apply_bin: queued pending path={msg}")
    _message_box(
        _um(
            "BIN_SCHEDULED_BACKGROUND_TEMPLATE",
            "更新ファイルを取得し、次回起動時の適用を予約しました。\n\n"
            "次回 Excel 起動時に、進捗ウィンドウで更新処理を実行します。\n"
            "処理が完了したら通常起動を継続します。\n\n"
            "ログ: {log_path}",
            log_path=str(_update_log_path(root)),
        ),
        owner_hwnd=owner_hwnd,
        sheet_id=sheet_id,
    )


def _show_bootstrap_update_prompt(
    st: dict[str, Any], *, owner_hwnd: int | None = None, sheet_id: str = ""
) -> None:
    root = _install_root()
    msg = _um(
        "BOOTSTRAP_UPDATE_CONFIRM_TEMPLATE",
        "bootstrap の新しい版があります。\n\n"
        "現在 (bootstrap): {installed_bootstrap}\n"
        "配布 (bootstrap): {latest_bootstrap}\n"
        "表示版: {display_version}\n\n"
        "「はい」を選ぶと、次回 Excel 起動時に bootstrap（更新ランナー）を更新します。\n"
        "「いいえ」は後で更新できます。",
        installed_bootstrap=str(st.get("installed_bootstrap_version") or "-"),
        latest_bootstrap=str(st.get("latest_bootstrap_version") or "-"),
        display_version=str(st.get("display_version") or "-"),
    )
    rc = _message_box(
        msg,
        style=MB_YESNO | MB_ICONINFORMATION,
        owner_hwnd=owner_hwnd,
        sheet_id=sheet_id,
    )
    if rc != IDYES:
        if root and root.is_dir():
            _append_update_log(root, "apply_bootstrap: user_decision=start_no")
        return
    if not root or not root.is_dir():
        _message_box(
            _um("BIN_UPDATE_INVALID_ROOT_OR_ZIP", "インストール先が無効です。"),
            style=MB_OK | MB_ICONWARNING,
            owner_hwnd=owner_hwnd,
            sheet_id=sheet_id,
        )
        return
    _append_update_log(root, "apply_bootstrap: user_decision=start_yes")
    rc_admin = _message_box(
        _um(
            "UPDATE_RUN_AS_ADMIN_CONFIRM_TEMPLATE",
            "管理者権限で更新を適用しますか？\n\n"
            "「はい」: 次回起動時に UAC 確認後、管理者権限で更新を適用します。\n"
            "「いいえ」: 通常権限で更新します（Windows の表示版同期に失敗する場合があります）。",
        ),
        style=MB_YESNO | MB_ICONINFORMATION,
        owner_hwnd=owner_hwnd,
        sheet_id=sheet_id,
    )
    require_admin = rc_admin == IDYES
    _append_update_log(root, f"apply_bootstrap: require_admin_selected={require_admin}")
    ok, qmsg = _queue_pending_bootstrap_only(st, source="interactive_bootstrap", require_admin=require_admin)
    if not ok:
        _append_update_log(root, f"apply_bootstrap: queue pending failed err={qmsg}")
        _message_box(
            _um(
                "BIN_WORKER_SPAWN_FAILED_TEMPLATE",
                "bootstrap 更新の予約に失敗しました。\n{error}",
                error=str(qmsg),
            ),
            style=MB_OK | MB_ICONWARNING,
            owner_hwnd=owner_hwnd,
            sheet_id=sheet_id,
        )
        return
    _append_update_log(root, f"apply_bootstrap: queued pending path={qmsg}")
    _message_box(
        _um(
            "BOOTSTRAP_SCHEDULED_BACKGROUND_TEMPLATE",
            "bootstrap 更新を次回 Excel 起動時に予約しました。\n\n"
            "起動時に確認後、進捗ウィンドウで適用します。\n\n"
            "ログ: {log_path}",
            log_path=str(_update_log_path(root)),
        ),
        owner_hwnd=owner_hwnd,
        sheet_id=sheet_id,
    )


def check_for_updates_interactive(
    source: str, *, owner_hwnd: int | None = None, sheet_id: str = ""
) -> None:
    _notify_pending_bin_apply_success(owner_hwnd=owner_hwnd, sheet_id=sheet_id)
    notify_offline = True
    try:
        st = check_for_updates(
            source=source,
            notify_offline=notify_offline,
            owner_hwnd=owner_hwnd,
            sheet_id=sheet_id,
        )
    except Exception as e:
        logger.exception("packaged_update interactive failed source=%s", source)
        _append_update_log(_install_root(), f"source={source} exception={type(e).__name__}: {e}")
        if not _UPDATE_CHECK_UI_LOCK.acquire(blocking=False):
            return
        try:
            _message_box(
                _um(
                    "INTERACTIVE_ERROR_TEMPLATE",
                    "更新確認中にエラーが発生しました。\n"
                    "hc_csv.log および hc_update.log を参照してください。\n\n{error}",
                    error=str(e)[:500],
                ),
                owner_hwnd=owner_hwnd,
                sheet_id=sheet_id,
            )
        finally:
            _UPDATE_CHECK_UI_LOCK.release()
        return

    if not _UPDATE_CHECK_UI_LOCK.acquire(blocking=False):
        _message_box(
            _um(
                "UPDATE_CHECK_BUSY",
                "更新確認はすでに実行中です。\n\n完了してから再度お試しください。",
            ),
            owner_hwnd=owner_hwnd,
            sheet_id=sheet_id,
        )
        return
    try:
        if st.get("error") == "skipped_no_catalog_config":
            _message_box(
                _um(
                    "SKIPPED_NO_CATALOG_TEMPLATE",
                    "更新確認の対象が設定されていません。\n"
                    "環境変数 HC_DEPLOY_ROOT（共有の配布ルート）または HC_CATALOG_PATH（catalog.json のフルパス）、\n"
                    "またはインストール先の config\\catalog_path.txt（1 行目に catalog.json のパス）を設定してください。\n\n"
                    "ログ: {log_path}",
                    log_path=str(_update_log_path(_install_root())),
                ),
                owner_hwnd=owner_hwnd,
                sheet_id=sheet_id,
            )
            return
        if not st["ok"]:
            err = st.get("error") or "unknown"
            if err == "HC_INSTALL_ROOT not set or not a directory":
                _message_box(
                    _um(
                        "CHECK_FAILED_INSTALL_ROOT_TEMPLATE",
                        "更新確認できませんでした。\n"
                        "HC_INSTALL_ROOT が設定されていないか、フォルダではありません。\n"
                        "配布モードではインストール直後に Excel を開き直すか、"
                        "start_excel_packaged_test.bat 等で起動してください。\n\n"
                        "ログ: {log_path}",
                        log_path=str(_update_log_path(None)),
                    ),
                    owner_hwnd=owner_hwnd,
                    sheet_id=sheet_id,
                )
            else:
                _message_box(
                    _um(
                        "CHECK_FAILED_GENERIC_TEMPLATE",
                        "更新確認できませんでした。\n{error}\n\nログ: {log_path}",
                        error=str(err),
                        log_path=str(_update_log_path(_install_root())),
                    ),
                    owner_hwnd=owner_hwnd,
                    sheet_id=sheet_id,
                )
            return

        if st.get("config_update_error"):
            _message_box(
                _um(
                    "CONFIG_SILENT_UPDATE_FAILED_TEMPLATE",
                    "config のサイレント更新に失敗しました。\n{error}\n\nログ: {log_path}",
                    error=str(st["config_update_error"]),
                    log_path=str(_update_log_path(_install_root())),
                ),
                style=MB_OK | MB_ICONWARNING,
                owner_hwnd=owner_hwnd,
                sheet_id=sheet_id,
            )

        if st.get("needs_bin_update"):
            if st.get("bin_update_prepare_error"):
                _message_box(
                    _um(
                        "BIN_PREPARE_FAILED_INTERACTIVE_TEMPLATE",
                        "bin の新しい版があります（配布 {latest_bin}）。\n\n"
                        "自動適用の準備に失敗しました: {prepare_error}\n\n"
                        "catalog・共有上の zip・sha256 を確認するか、手動でフル適用してください。\n\n"
                        "ログ: {log_path}",
                        latest_bin=str(st.get("latest_bin_version") or "-"),
                        prepare_error=str(st.get("bin_update_prepare_error")),
                        log_path=str(_update_log_path(_install_root())),
                    ),
                    style=MB_OK | MB_ICONWARNING,
                    owner_hwnd=owner_hwnd,
                    sheet_id=sheet_id,
                )
                return
            _show_bin_update_prompt(st, owner_hwnd=owner_hwnd, sheet_id=sheet_id)
            return

        if st.get("needs_bootstrap_update") and not st.get("needs_bin_update"):
            cpath = str(st.get("catalog_path") or "").strip()
            if not cpath or not Path(cpath).is_file():
                _message_box(
                    _um(
                        "BOOTSTRAP_PREP_FAILED_INTERACTIVE_TEMPLATE",
                        "bootstrap の更新を確認できません。\n{prepare_error}\n\nログ: {log_path}",
                        prepare_error="catalog.json のパスが無効です。",
                        log_path=str(_update_log_path(_install_root())),
                    ),
                    style=MB_OK | MB_ICONWARNING,
                    owner_hwnd=owner_hwnd,
                    sheet_id=sheet_id,
                )
                return
            data = load_catalog(Path(cpath))
            bzp, _, berr = (
                _prepare_bootstrap_full_apply(data, Path(cpath))
                if isinstance(data, dict)
                else (None, None, "catalog invalid")
            )
            if bzp is None or not bzp.is_file():
                _message_box(
                    _um(
                        "BOOTSTRAP_PREP_FAILED_INTERACTIVE_TEMPLATE",
                        "bootstrap の新しい版がありますが、配布ファイルの準備に失敗しました。\n"
                        "{prepare_error}\n\nログ: {log_path}",
                        prepare_error=str(berr or "bootstrap.full not found"),
                        log_path=str(_update_log_path(_install_root())),
                    ),
                    style=MB_OK | MB_ICONWARNING,
                    owner_hwnd=owner_hwnd,
                    sheet_id=sheet_id,
                )
                return
            _show_bootstrap_update_prompt(st, owner_hwnd=owner_hwnd, sheet_id=sheet_id)
            return

        if st.get("config_update_applied"):
            _message_box(
                _um(
                    "CONFIG_UPDATED_TEMPLATE",
                    "config を更新しました。\n\n新しい config 版: {installed_config}",
                    installed_config=str(st.get("installed_config") or "(unknown)"),
                ),
                owner_hwnd=owner_hwnd,
                sheet_id=sheet_id,
            )
            return

        latest_msg = _um(
            "BIN_UPDATE_LATEST_TEMPLATE",
            "お使いの版は最新です。\n\nbin: {installed_bin}\nconfig: {installed_config}\nbootstrap: {installed_bootstrap}\n表示版: {display_version}",
            installed_bin=str(st.get("installed_bin") or "-"),
            installed_config=str(st.get("installed_config") or "-"),
            installed_bootstrap=str(st.get("installed_bootstrap_version") or "-"),
            display_version=str(st.get("display_version") or "-"),
        )
        _message_box(latest_msg, owner_hwnd=owner_hwnd, sheet_id=sheet_id)
    finally:
        _UPDATE_CHECK_UI_LOCK.release()


def maybe_apply_pending_bootstrap_update(*, owner_hwnd: int | None = None, sheet_id: str = "_") -> None:
    global _SUPPRESS_STARTUP_BIN_PROMPT_AFTER_PENDING_DEFER
    _ = owner_hwnd, sheet_id
    root = _install_root()
    if root is None or not root.is_dir():
        return
    pending = read_pending(build_paths(root))
    require_admin = bool((pending or {}).get("require_admin", False))
    _append_update_log(
        root,
        "bootstrap_apply: pending_detected require_admin={ra} elevated_now={el}".format(
            ra=require_admin,
            el=_is_process_elevated(),
        ),
    )
    if require_admin and not _is_process_elevated():
        rc = _message_box(
            _um(
                "PENDING_APPLY_RUN_AS_ADMIN_TEMPLATE",
                "予約された更新は管理者権限での適用が選択されています。\n\n"
                "「はい」: UAC 確認後、管理者権限で適用します。\n"
                "「いいえ」: 今回は通常権限で適用します（表示版同期に失敗する場合があります）。",
            ),
            style=MB_YESNO | MB_ICONINFORMATION,
            owner_hwnd=0,
            sheet_id=sheet_id,
        )
        if rc == IDYES:
            ok_elev, detail = _run_pending_apply_elevated(root)
            _append_update_log(root, f"bootstrap_apply elevated_attempt ok={ok_elev} detail={detail}")
            if ok_elev:
                _SUPPRESS_STARTUP_BIN_PROMPT_AFTER_PENDING_DEFER = False
                return
            _message_box(
                _um(
                    "PENDING_APPLY_RUN_AS_ADMIN_FAILED_TEMPLATE",
                    "管理者権限での更新適用に失敗またはキャンセルされました。\n"
                    "通常権限で続行します。\n\n詳細: {detail}",
                    detail=detail,
                ),
                style=MB_OK | MB_ICONWARNING,
                owner_hwnd=0,
                sheet_id=sheet_id,
            )
        else:
            _append_update_log(root, "bootstrap_apply: user_declined_elevation continue_non_admin=true")
    try:
        from bootstrap.update_bootstrap import apply_pending_update
    except Exception as e:
        _append_update_log(root, f"bootstrap_apply skipped import_failed err={type(e).__name__}: {e}")
        return
    try:
        res = apply_pending_update(root)
        if isinstance(res, dict):
            if res.get("deferred"):
                _SUPPRESS_STARTUP_BIN_PROMPT_AFTER_PENDING_DEFER = True
                _append_update_log(
                    root,
                    "startup: pending_apply deferred; skip duplicate bin prompt on same launch",
                )
            else:
                _SUPPRESS_STARTUP_BIN_PROMPT_AFTER_PENDING_DEFER = False
            if res.get("applied"):
                _append_update_log(root, "bootstrap_apply applied=true")
            elif not res.get("ok", True):
                _append_update_log(root, f"bootstrap_apply applied=false err={res.get('error')}")
    except Exception as e:
        _SUPPRESS_STARTUP_BIN_PROMPT_AFTER_PENDING_DEFER = False
        _append_update_log(root, f"bootstrap_apply failed err={type(e).__name__}: {e}")


def maybe_check_updates_on_startup(*, owner_hwnd: int | None = None, sheet_id: str = "_") -> None:
    """Called from svc_host after bridge registration when packaged.

    Pending bootstrap apply runs earlier in svc_host (before bridge/UI); do not
    call maybe_apply_pending_bootstrap_update here to avoid a second confirm dialog.
    """
    global _SUPPRESS_STARTUP_BIN_PROMPT_AFTER_PENDING_DEFER
    if not packaged_spawn_requested():
        return
    with _UPDATE_CHECK_UI_LOCK:
        _notify_pending_bin_apply_success(owner_hwnd=owner_hwnd, sheet_id=sheet_id)
    v = os.environ.get(ENV_UPDATE_CHECK_AT_STARTUP)
    if v is not None and not core_env.truthy(v, empty_means_false=True):
        _SUPPRESS_STARTUP_BIN_PROMPT_AFTER_PENDING_DEFER = False
        return

    st = check_for_updates(
        source="startup",
        notify_offline=False,
        owner_hwnd=owner_hwnd,
        sheet_id=sheet_id,
    )
    with _UPDATE_CHECK_UI_LOCK:
        if st.get("config_update_error"):
            _message_box(
                _um(
                    "CONFIG_AUTO_UPDATE_FAILED_TEMPLATE",
                    "config の自動更新に失敗しました。\n{error}\n\nログ: {log_path}",
                    error=str(st["config_update_error"]),
                    log_path=str(_update_log_path(_install_root())),
                ),
                style=MB_OK | MB_ICONWARNING,
                owner_hwnd=owner_hwnd,
                sheet_id=sheet_id,
            )

        if st.get("needs_bin_update"):
            if _SUPPRESS_STARTUP_BIN_PROMPT_AFTER_PENDING_DEFER:
                _SUPPRESS_STARTUP_BIN_PROMPT_AFTER_PENDING_DEFER = False
                lr = _install_root()
                if lr and lr.is_dir():
                    _append_update_log(
                        lr,
                        "startup: bin update prompt skipped (pending_apply was deferred this launch)",
                    )
            else:
                _show_bin_update_prompt(st, owner_hwnd=owner_hwnd, sheet_id=sheet_id)
