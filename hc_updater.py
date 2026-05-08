#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Standalone updater worker for packaged bin apply.

This removes dependency on PowerShell -File worker scripts so update apply can
run in environments where script execution is restricted.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import shutil
import subprocess
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any


def _append_with_cap(path: Path, line: str, max_bytes: int = 1024 * 1024) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8", errors="replace") as f:
        f.write(line)
    try:
        if path.stat().st_size <= max_bytes:
            return
        with path.open("rb") as rf:
            rf.seek(-max_bytes, os.SEEK_END)
            keep = rf.read(max_bytes)
        with path.open("wb") as wf:
            wf.write(keep)
    except OSError:
        pass


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().lower()


def _remove_path(path_text: str) -> None:
    if not path_text:
        return
    p = Path(path_text)
    try:
        if p.is_file():
            p.unlink(missing_ok=True)
        elif p.is_dir():
            shutil.rmtree(p, ignore_errors=True)
    except OSError:
        pass


def _is_excel_running() -> bool:
    try:
        cp = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True,
            text=True,
            check=False,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
        out = (cp.stdout or "").splitlines()
        for line in out:
            txt = line.strip().strip('"')
            if not txt:
                continue
            # Locale-agnostic detection: CSV first column is image name.
            first = txt.split(",", 1)[0].strip().strip('"').upper()
            if first == "EXCEL.EXE":
                return True
        return False
    except OSError:
        return False


def _load_update_messages_for_job(job_path: Path) -> dict[str, str]:
    msgs: dict[str, str] = {}
    try:
        raw = json.loads(job_path.read_text(encoding="utf-8-sig"))
        install_root = Path(str(raw.get("InstallRoot", "")).strip())
        if not install_root:
            return msgs
        cfg_path = install_root / "config" / "ui_update_check.json"
        if not cfg_path.is_file():
            return msgs
        cfg = json.loads(cfg_path.read_text(encoding="utf-8-sig"))
        maybe = cfg.get("MESSAGES") if isinstance(cfg, dict) else None
        if isinstance(maybe, dict):
            for k, v in maybe.items():
                if isinstance(k, str) and isinstance(v, str):
                    msgs[k] = v
    except Exception:
        pass
    return msgs


def _copy_merge_tree(src_root: Path, dst_root: Path) -> None:
    if not src_root.exists():
        return
    for fp in src_root.rglob("*"):
        if not fp.is_file():
            continue
        rel = fp.relative_to(src_root)
        dst = dst_root / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(fp, dst)


def _mirror_tree(src_root: Path, dst_root: Path) -> None:
    # Equivalent intent to robocopy /MIR used by old worker.
    if dst_root.exists():
        shutil.rmtree(dst_root, ignore_errors=True)
    shutil.copytree(src_root, dst_root)


def _apply_delete_list(install_root: Path, delete_list_path: Path) -> None:
    if not delete_list_path.is_file():
        return
    try:
        lines = delete_list_path.read_text(encoding="utf-8-sig").splitlines()
    except OSError:
        return
    for raw in lines:
        rel = raw.strip().replace("\\", "/").lstrip("/")
        if not rel:
            continue
        target = install_root / Path(rel)
        try:
            if target.is_file():
                target.unlink(missing_ok=True)
            elif target.is_dir():
                shutil.rmtree(target, ignore_errors=True)
        except OSError:
            pass


def _try_set_display_values(display_version: str, log_path: Path) -> None:
    if not display_version:
        return
    try:
        import winreg  # type: ignore
    except Exception:
        return
    display_name = "CSV Tool"
    sub = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\{B5E8C4A2-1D3F-4E6B-9C0D-1A2B3C4D5E6F}_is1"
    candidates = [
        (winreg.HKEY_LOCAL_MACHINE, sub),
        (winreg.HKEY_LOCAL_MACHINE, r"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{B5E8C4A2-1D3F-4E6B-9C0D-1A2B3C4D5E6F}_is1"),
        (winreg.HKEY_CURRENT_USER, sub),
        (winreg.HKEY_CURRENT_USER, r"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{B5E8C4A2-1D3F-4E6B-9C0D-1A2B3C4D5E6F}_is1"),
    ]
    for root, subkey in candidates:
        try:
            with winreg.OpenKey(root, subkey, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, display_name)
                winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, display_version)
            _append_with_cap(log_path, f"{_ts()} apply_bin: DisplayName/DisplayVersion updated value={display_version} key={subkey}\n")
            return
        except OSError:
            continue
    for subkey in (
        sub,
        r"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{B5E8C4A2-1D3F-4E6B-9C0D-1A2B3C4D5E6F}_is1",
    ):
        try:
            with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, subkey, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, display_name)
                winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, display_version)
            _append_with_cap(
                log_path,
                f"{_ts()} apply_bin: DisplayName/DisplayVersion updated (HKCU created) value={display_version} key={subkey}\n",
            )
            return
        except OSError:
            continue
    _append_with_cap(log_path, f"{_ts()} apply_bin: DisplayName/DisplayVersion key not found or not writable value={display_version}\n")


def _message_box(text: str, title: str = "CSV Tool update", icon: int = 0x40) -> None:
    try:
        import ctypes

        ctypes.windll.user32.MessageBoxW(0, text, title, 0x0 | icon)
    except Exception:
        pass


class _ProgressUi:
    def __init__(self, messages: dict[str, str] | None = None) -> None:
        self._ok = False
        self._root = None
        self._title = None
        self._msg = None
        self._bar = None
        self._messages = messages or {}
        try:
            import tkinter as tk
            from tkinter import ttk

            root = tk.Tk()
            root.title(self._messages.get("UPDATER_WINDOW_TITLE", "CSV Tool update"))
            root.attributes("-topmost", True)
            root.geometry("540x190")
            root.resizable(False, False)
            self._title = tk.StringVar(value=self._messages.get("UPDATER_INITIAL_STATUS", "Status: waiting"))
            self._msg = tk.StringVar(value=self._messages.get("UPDATER_INITIAL_MESSAGE", "Waiting for all Excel processes to exit."))
            ttk.Label(root, textvariable=self._title).pack(anchor="w", padx=16, pady=(16, 6))
            ttk.Label(root, textvariable=self._msg).pack(anchor="w", padx=16, pady=(0, 10))
            self._bar = ttk.Progressbar(root, orient="horizontal", mode="determinate", maximum=100, length=500)
            self._bar.pack(anchor="w", padx=16, pady=(0, 6))
            root.update_idletasks()
            root.update()
            self._root = root
            self._ok = True
        except Exception:
            self._ok = False

    def set(self, title: str, message: str, progress: int) -> None:
        if not self._ok:
            return
        try:
            self._title.set(f"Status: {title}")
            self._msg.set(message)
            self._bar["value"] = max(0, min(100, int(progress)))
            self._root.update_idletasks()
            self._root.update()
        except Exception:
            pass

    def close(self) -> None:
        if not self._ok:
            return
        try:
            self._root.destroy()
        except Exception:
            pass


def _ts() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime())


@dataclass
class Job:
    install_root: Path
    zip_path: Path
    expected_sha: str
    log_path: Path
    display_version: str
    apply_mode: str
    target_bin_version: str
    cleanup_path: str
    notify_marker_path: str


def _phase(log_path: Path, ui: _ProgressUi, key: str, title: str, msg: str, progress: int) -> None:
    _append_with_cap(log_path, f"{_ts()} apply_bin: phase={key} message={msg}\n")
    ui.set(title, msg, progress)


def _run_apply_pending_job(job_path: Path, raw: dict[str, Any]) -> int:
    """UAC 昇格で pending を適用（UI なし）。結果 JSON を ResultPath に書く。"""
    import traceback

    install_root = Path(str(raw.get("InstallRoot", "")).strip())
    result_s = str(raw.get("ResultPath", "")).strip()
    result_path = Path(result_s) if result_s else Path()
    log_s = str(raw.get("LogPath", "")).strip()
    log_line = Path(log_s) if log_s else (install_root / "logs" / "hc_update.log")
    try:
        if not install_root.is_dir():
            raise RuntimeError(f"InstallRoot not found: {install_root}")
        if not result_s:
            raise RuntimeError("ResultPath is empty")
        _append_with_cap(
            log_line,
            f"{_ts()} apply_pending: elevated worker start pid={os.getpid()} install_root={install_root}\n",
        )
        from bootstrap.update_bootstrap import apply_pending_update

        res = apply_pending_update(install_root)
        out: dict[str, Any] = {"ok": False, "applied": False}
        if isinstance(res, dict):
            out.update(res)
            out["ok"] = bool(res.get("ok", True))
        else:
            out = {"ok": True, "applied": False}
        result_path.parent.mkdir(parents=True, exist_ok=True)
        result_path.write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
        _append_with_cap(
            log_line,
            f"{_ts()} apply_pending: done ok={out.get('ok')} applied={out.get('applied')}\n",
        )
        return 0 if out.get("ok") else 1
    except Exception as e:
        tb = traceback.format_exc()
        try:
            err_obj: dict[str, Any] = {
                "ok": False,
                "applied": False,
                "error": f"{type(e).__name__}: {e}",
                "traceback": tb,
            }
            if result_s:
                result_path.parent.mkdir(parents=True, exist_ok=True)
                result_path.write_text(json.dumps(err_obj, ensure_ascii=False), encoding="utf-8")
        except OSError:
            pass
        _append_with_cap(log_line, f"{_ts()} apply_pending: ERROR {e}\n")
        return 1
    finally:
        try:
            job_path.unlink(missing_ok=True)
        except OSError:
            pass


def _load_job(job_path: Path) -> Job:
    raw = json.loads(job_path.read_text(encoding="utf-8-sig"))
    return Job(
        install_root=Path(str(raw.get("InstallRoot", "")).strip()),
        zip_path=Path(str(raw.get("ZipPath", "")).strip()),
        expected_sha=str(raw.get("ExpectedSha", "")).strip().lower(),
        log_path=Path(str(raw.get("LogPath", "")).strip()),
        display_version=str(raw.get("DisplayVersion", "")).strip(),
        apply_mode=str(raw.get("ApplyMode", "full")).strip().lower() or "full",
        target_bin_version=str(raw.get("TargetBinVersion", "")).strip(),
        cleanup_path=str(raw.get("CleanupPath", "")).strip(),
        notify_marker_path=str(raw.get("NotifyMarkerPath", "")).strip(),
    )


def _write_marker(path_text: str, target_bin_version: str, display_version: str, log_path: Path) -> None:
    if not path_text:
        return
    p = Path(path_text)
    p.parent.mkdir(parents=True, exist_ok=True)
    data = {
        "ts": _ts(),
        "target_bin_version": target_bin_version,
        "display_version": display_version,
        "log_path": str(log_path),
    }
    p.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    _append_with_cap(log_path, f"{_ts()} apply_bin: success_notify marker_written path={p}\n")


def run(job_path: Path) -> int:
    try:
        raw_head = json.loads(job_path.read_text(encoding="utf-8-sig"))
    except Exception as e:
        bootstrap_log_path = Path(tempfile.gettempdir()) / "csv_tool" / "hc_update.log"
        _append_with_cap(
            bootstrap_log_path,
            f"{_ts()} apply_bin: job read failed {type(e).__name__}: {e} path={job_path}\n",
        )
        return 1
    jt = str(raw_head.get("JobType", "bin_apply")).strip().lower()
    if jt == "apply_pending":
        return _run_apply_pending_job(job_path, raw_head)

    bootstrap_log_path = Path(tempfile.gettempdir()) / "csv_tool" / "hc_update.log"
    _append_with_cap(
        bootstrap_log_path,
        f"{_ts()} apply_bin: updater process start pid={os.getpid()} job={job_path}\n",
    )
    ui_msgs = _load_update_messages_for_job(job_path)
    ui = _ProgressUi(ui_msgs)
    try:
        job = _load_job(job_path)
        _append_with_cap(
            job.log_path,
            f"{_ts()} apply_bin: updater job loaded pid={os.getpid()} job={job_path}\n",
        )
        if not job.install_root.is_dir():
            raise RuntimeError(f"InstallRoot not found: {job.install_root}")
        if not job.zip_path.is_file():
            raise RuntimeError(f"ZipPath not found: {job.zip_path}")
        if not str(job.log_path):
            raise RuntimeError("LogPath is empty")

        _phase(
            job.log_path,
            ui,
            "wait_excel_exit",
            ui_msgs.get("UPDATER_PHASE_WAIT_TITLE", "waiting"),
            ui_msgs.get("UPDATER_PHASE_WAIT_MESSAGE", "Waiting for all Excel processes to exit."),
            5,
        )
        while _is_excel_running():
            time.sleep(2)
            ui.set(
                ui_msgs.get("UPDATER_PHASE_WAIT_TITLE", "waiting"),
                ui_msgs.get("UPDATER_PHASE_WAIT_MESSAGE", "Waiting for all Excel processes to exit."),
                5,
            )
        time.sleep(1)

        _phase(
            job.log_path,
            ui,
            "start",
            ui_msgs.get("UPDATER_PHASE_START_TITLE", "start"),
            ui_msgs.get("UPDATER_PHASE_START_MESSAGE", "Starting update process."),
            15,
        )
        dl_tmp = Path(tempfile.mkdtemp(prefix="csv_tool_bin_download_"))
        zip_local = dl_tmp / job.zip_path.name
        _phase(
            job.log_path,
            ui,
            "download_copy",
            ui_msgs.get("UPDATER_PHASE_DOWNLOAD_TITLE", "downloading"),
            ui_msgs.get("UPDATER_PHASE_DOWNLOAD_MESSAGE", "Copying update archive."),
            30,
        )
        shutil.copy2(job.zip_path, zip_local)

        if job.expected_sha:
            got = _sha256_file(zip_local)
            if got != job.expected_sha:
                raise RuntimeError(f"sha256 mismatch got={got}")

        extract_tmp = Path(tempfile.mkdtemp(prefix="csv_tool_bin_extract_"))
        try:
            _phase(
                job.log_path,
                ui,
                "extract",
                ui_msgs.get("UPDATER_PHASE_EXTRACT_TITLE", "extracting"),
                ui_msgs.get("UPDATER_PHASE_EXTRACT_MESSAGE", "Extracting update archive."),
                50,
            )
            shutil.unpack_archive(str(zip_local), str(extract_tmp))
            _phase(
                job.log_path,
                ui,
                "apply",
                ui_msgs.get("UPDATER_PHASE_APPLY_TITLE", "applying"),
                ui_msgs.get("UPDATER_PHASE_APPLY_MESSAGE", "Applying update."),
                75,
            )

            if job.apply_mode == "patch":
                patch_app = extract_tmp / "app" / "bin"
                patch_addin = extract_tmp / "addin"
                if not patch_app.exists() and not patch_addin.exists():
                    raise RuntimeError("invalid patch zip: need app/bin and/or addin")
                if patch_app.exists():
                    dst_bin = job.install_root / "app" / "bin"
                    dst_bin.mkdir(parents=True, exist_ok=True)
                    _copy_merge_tree(patch_app, dst_bin)
                if patch_addin.exists():
                    dst_addin = job.install_root / "addin"
                    dst_addin.mkdir(parents=True, exist_ok=True)
                    _copy_merge_tree(patch_addin, dst_addin)
                _apply_delete_list(job.install_root, extract_tmp / "__delete_list.txt")
                if job.target_bin_version:
                    (job.install_root / "VERSION.txt").write_text(job.target_bin_version + "\n", encoding="utf-8")
                _append_with_cap(job.log_path, f"{_ts()} apply_bin: patch merged TargetBinVersion={job.target_bin_version}\n")
            else:
                src_bin = extract_tmp / "app" / "bin"
                if not src_bin.is_dir():
                    raise RuntimeError("invalid zip: missing app/bin")
                _mirror_tree(src_bin, job.install_root / "app" / "bin")
                vsrc = extract_tmp / "VERSION.txt"
                if job.target_bin_version:
                    (job.install_root / "VERSION.txt").write_text(job.target_bin_version + "\n", encoding="utf-8")
                elif vsrc.is_file():
                    shutil.copy2(vsrc, job.install_root / "VERSION.txt")
                addin_src = extract_tmp / "addin"
                if addin_src.is_dir():
                    _mirror_tree(addin_src, job.install_root / "addin")

            _try_set_display_values(job.display_version, job.log_path)
            _write_marker(job.notify_marker_path, job.target_bin_version, job.display_version, job.log_path)
            _append_with_cap(job.log_path, f"{_ts()} apply_bin: success\n")
            _phase(
                job.log_path,
                ui,
                "done",
                ui_msgs.get("UPDATER_PHASE_DONE_TITLE", "done"),
                ui_msgs.get("UPDATER_PHASE_DONE_MESSAGE", "Update completed. Please restart Excel."),
                100,
            )
            _message_box(
                ui_msgs.get("UPDATER_SUCCESS_MESSAGE", "bin update completed.\nPlease restart Excel."),
                ui_msgs.get("UPDATER_WINDOW_TITLE", "CSV Tool update"),
                0x40,
            )
            return 0
        finally:
            shutil.rmtree(extract_tmp, ignore_errors=True)
            shutil.rmtree(dl_tmp, ignore_errors=True)
    except Exception as e:
        # Keep same keyword "ERROR" used by existing operations.
        log_path = Path(tempfile.gettempdir()) / "csv_tool" / "hc_update.log"
        try:
            if "job" in locals():
                log_path = job.log_path
        except Exception:
            pass
        _append_with_cap(log_path, f"{_ts()} apply_bin: ERROR {e}\n")
        err_tpl = ui_msgs.get("UPDATER_ERROR_TEMPLATE", "bin update failed.\n\n{error}\n\nlog: {log_path}")
        try:
            err_msg = err_tpl.format(error=e, log_path=log_path)
        except Exception:
            err_msg = f"bin update failed.\n\n{e}\n\nlog: {log_path}"
        _message_box(err_msg, ui_msgs.get("UPDATER_WINDOW_TITLE", "CSV Tool update"), 0x30)
        return 1
    finally:
        try:
            if "job" in locals():
                _remove_path(job.cleanup_path)
        except Exception:
            pass
        _remove_path(str(job_path))
        ui.close()


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--job", required=True, help="Path to updater job json")
    args = ap.parse_args()
    return run(Path(args.job))


if __name__ == "__main__":
    raise SystemExit(main())

