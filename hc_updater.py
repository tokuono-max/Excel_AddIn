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
import sys
import tempfile
import time
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from core.win_running_file_replace import (
    cleanup_stale_sidecar_files,
    collect_self_sidecar_dst_paths,
    copy_file_with_sharing_fallback,
)
from core.update_process_cleanup import (
    ensure_packaged_children_stopped,
    mutex_snapshot,
    probe_tasklist_line,
    sleep_with_ui_pulse,
    taskkill_other_hc_updater_processes,
)
from core.update_housekeeping import cleanup_update_payload_dir, post_deferred_bin_success_housekeeping
from core.packaged_update import (
    _bin_apply_success_marker_path,
    _resolve_install_scope,
    display_name_for_install_scope,
    notify_installed_apps_list_changed,
    updater_result_path,
)
from core.update_state import load_runtime_config


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


def _load_update_messages_for_install(install_root: Path) -> dict[str, str]:
    msgs: dict[str, str] = {}
    try:
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


def _load_update_messages_for_job(job_path: Path) -> dict[str, str]:
    msgs: dict[str, str] = {}
    try:
        raw = json.loads(job_path.read_text(encoding="utf-8-sig"))
        install_root = Path(str(raw.get("InstallRoot", "")).strip())
        if not install_root:
            return msgs
        return _load_update_messages_for_install(install_root)
    except Exception:
        pass
    return msgs


def _is_patch_apply_mode(apply_mode: str) -> bool:
    return str(apply_mode or "").strip().lower() == "patch"


def updater_busy_title(ui_msgs: dict[str, str], apply_mode: str) -> str:
    if _is_patch_apply_mode(apply_mode):
        return str(ui_msgs.get("UPDATER_PHASE_BUSY_TITLE_PATCH") or "差分更新中")
    return str(ui_msgs.get("UPDATER_PHASE_BUSY_TITLE_FULL") or "フル更新中")


def updater_busy_body(
    ui_msgs: dict[str, str],
    apply_mode: str,
    message_key: str,
    default: str,
    *,
    apply_phase: bool = False,
) -> str:
    if apply_phase:
        if _is_patch_apply_mode(apply_mode):
            return str(
                ui_msgs.get("UPDATER_PHASE_APPLY_MESSAGE_PATCH")
                or "差分パッケージをインストールしています"
            )
        return str(
            ui_msgs.get("UPDATER_PHASE_APPLY_MESSAGE_FULL") or "フルパッケージをインストールしています"
        )
    prefix = "差分" if _is_patch_apply_mode(apply_mode) else "フル"
    base = str(ui_msgs.get(message_key) or default).strip() or default
    if base.startswith(prefix):
        return base
    return f"{prefix} {base}"


def _self_hc_updater_dst_paths() -> set[Path]:
    """Paths under app\\bin that this updater run may replace via sidecar."""
    extra: set[Path] = set()
    try:
        here = Path(__file__).resolve()
        if here.suffix.lower() == ".py":
            sibling = here.parent / "hc_updater.exe"
            if sibling.is_file():
                extra.add(sibling.resolve())
    except OSError:
        pass
    return collect_self_sidecar_dst_paths(extra=extra)


def _cleanup_stale_renamed_updaters(bin_dir: Path, log_path: Path | None) -> None:
    def _log(msg: str) -> None:
        if log_path is not None:
            _append_with_cap(log_path, f"{_ts()} apply_bin: {msg}\n")

    cleanup_stale_sidecar_files(bin_dir, _log)


def _copy_merge_tree(
    src_root: Path,
    dst_root: Path,
    log_path: Path | None = None,
    *,
    ui: _ProgressUi | None = None,
    ui_title: str = "更新中",
    ui_message: str = "",
    progress_lo: int = 75,
    progress_hi: int = 95,
) -> None:
    if not src_root.exists():
        return
    proactive = _self_hc_updater_dst_paths()

    def _log(msg: str) -> None:
        if log_path is not None:
            _append_with_cap(log_path, f"{_ts()} apply_bin: {msg}\n")

    files = [p for p in src_root.rglob("*") if p.is_file()]
    total = max(len(files), 1)
    for idx, fp in enumerate(files, 1):
        rel = fp.relative_to(src_root)
        dst = dst_root / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        try:
            copy_file_with_sharing_fallback(
                fp,
                dst,
                proactive_sidecar=proactive,
                log=_log,
                rel_label=rel.as_posix(),
            )
        except OSError as e:
            if log_path is not None:
                _append_with_cap(
                    log_path,
                    f"{_ts()} apply_bin: copy_merge_fail rel={rel.as_posix()} err={e!s} "
                    f"winerror={getattr(e, 'winerror', None)}\n",
                )
            raise
        if ui is not None and ui._ok:
            pct = progress_lo + int((progress_hi - progress_lo) * idx / total)
            ui.set(ui_title, ui_message, pct)


def _running_from_tree_root(dst_root: Path) -> bool:
    """True when this process executable lives under dst_root (cannot safely rmtree dst)."""
    try:
        exe_parent = Path(sys.executable).resolve().parent
        return exe_parent == dst_root.resolve()
    except OSError:
        return False


def _mirror_tree(
    src_root: Path,
    dst_root: Path,
    log_path: Path,
    *,
    ui: _ProgressUi | None = None,
    ui_title: str = "更新中",
    ui_message: str = "",
    progress_lo: int = 75,
    progress_hi: int = 95,
) -> None:
    # Equivalent intent to robocopy /MIR used by old worker.
    if _running_from_tree_root(dst_root):
        raise RuntimeError(
            f"mirror refused: updater runs from dst={dst_root} (use copy_merge for app/bin)"
        )
    _append_with_cap(log_path, f"{_ts()} apply_bin: mirror rmtree_if_exists dst={dst_root}\n")
    proactive = _self_hc_updater_dst_paths()
    self_dst = dst_root / "hc_updater.exe"
    if dst_root.exists() and self_dst.is_file():
        try:
            if self_dst.resolve() in proactive:
                side = Path(tempfile.gettempdir()) / f"csv_tool_hc_updater_mirror_stale_{os.getpid()}.exe"
                try:
                    if side.is_file():
                        side.unlink(missing_ok=True)
                except OSError:
                    pass
                os.replace(self_dst, side)
                _append_with_cap(
                    log_path,
                    f"{_ts()} apply_bin: mirror moved_running_updater_exe to={side}\n",
                )
        except OSError as e:
            _append_with_cap(
                log_path,
                f"{_ts()} apply_bin: mirror move_running_updater_exe failed err={e!s}\n",
            )
    if dst_root.exists():
        shutil.rmtree(dst_root, ignore_errors=True)
    files = [p for p in src_root.rglob("*") if p.is_file()]
    total = max(len(files), 1)

    def _log(msg: str) -> None:
        _append_with_cap(log_path, f"{_ts()} apply_bin: {msg}\n")

    try:
        for idx, fp in enumerate(files, 1):
            rel = fp.relative_to(src_root)
            dst = dst_root / rel
            dst.parent.mkdir(parents=True, exist_ok=True)
            copy_file_with_sharing_fallback(
                fp,
                dst,
                proactive_sidecar=proactive,
                log=_log,
                rel_label=rel.as_posix(),
            )
            if ui is not None and ui._ok:
                pct = progress_lo + int((progress_hi - progress_lo) * idx / total)
                ui.set(ui_title, ui_message, pct)
    except OSError as e:
        _append_with_cap(
            log_path,
            f"{_ts()} apply_bin: copytree_fail src={src_root} dst={dst_root} err={e!s} winerror={getattr(e, 'winerror', None)}\n",
        )
        if int(getattr(e, "winerror", 0) or 0) == 32:
            _append_with_cap(
                log_path,
                f"{_ts()} apply_bin: winerror32 probe={probe_tasklist_line()}\n",
            )
            _append_with_cap(
                log_path,
                f"{_ts()} apply_bin: winerror32 mutex={mutex_snapshot()}\n",
            )
        raise


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


def _try_set_display_values(display_version: str, log_path: Path, install_root: Path) -> None:
    if not display_version:
        return
    try:
        import winreg  # type: ignore
    except Exception:
        return
    scope = _resolve_install_scope(install_root)
    sub = r"Software\Microsoft\Windows\CurrentVersion\Uninstall\{B5E8C4A2-1D3F-4E6B-9C0D-1A2B3C4D5E6F}_is1"
    candidates = [
        (winreg.HKEY_LOCAL_MACHINE, sub),
        (winreg.HKEY_LOCAL_MACHINE, r"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{B5E8C4A2-1D3F-4E6B-9C0D-1A2B3C4D5E6F}_is1"),
        (winreg.HKEY_CURRENT_USER, sub),
        (winreg.HKEY_CURRENT_USER, r"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\{B5E8C4A2-1D3F-4E6B-9C0D-1A2B3C4D5E6F}_is1"),
    ]
    for root, subkey in candidates:
        root_name = "HKLM" if root == winreg.HKEY_LOCAL_MACHINE else "HKCU"
        display_name = display_name_for_install_scope(scope, registry_root=root_name)
        try:
            with winreg.OpenKey(root, subkey, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, display_name)
                winreg.SetValueEx(key, "DisplayVersion", 0, winreg.REG_SZ, display_version)
            _append_with_cap(
                log_path,
                f"{_ts()} apply_bin: DisplayName/DisplayVersion updated name={display_name} value={display_version} key={subkey}\n",
            )
            notify_installed_apps_list_changed()
            _append_with_cap(log_path, f"{_ts()} apply_bin: installed_apps_list SHChangeNotify sent\n")
            return
        except OSError:
            continue
    display_name = display_name_for_install_scope(scope, registry_root="HKCU")
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
                f"{_ts()} apply_bin: DisplayName/DisplayVersion updated (HKCU created) name={display_name} value={display_version} key={subkey}\n",
            )
            notify_installed_apps_list_changed()
            _append_with_cap(log_path, f"{_ts()} apply_bin: installed_apps_list SHChangeNotify sent\n")
            return
        except OSError:
            continue
    _append_with_cap(log_path, f"{_ts()} apply_bin: DisplayName/DisplayVersion key not found or not writable value={display_version}\n")


_MB_SETFOREGROUND = 0x00010000
_MB_TOPMOST = 0x00040000


def _message_box(text: str, title: str = "CSV Tool update", icon: int = 0x40) -> None:
    try:
        import ctypes

        style = _MB_SETFOREGROUND | _MB_TOPMOST | int(icon)
        ctypes.windll.user32.MessageBoxW(0, text, title, style)
    except Exception as e:
        try:
            _append_with_cap(
                Path(tempfile.gettempdir()) / "csv_tool" / "hc_update.log",
                f"{_ts()} apply_bin: MessageBoxW failed type={type(e).__name__}: {e}\n",
            )
        except OSError:
            pass


def _excel_wait_timeout_sec(install_root: Path) -> int:
    cfg = load_runtime_config(install_root)
    try:
        return max(60, int(cfg.get("UPDATER_EXCEL_WAIT_TIMEOUT_SEC", 600)))
    except (TypeError, ValueError):
        return 600


def _write_updater_result(
    result_path: Path | None,
    *,
    ok: bool,
    error: str = "",
    target_bin_version: str = "",
    display_version: str = "",
) -> None:
    if result_path is None:
        return
    try:
        result_path.parent.mkdir(parents=True, exist_ok=True)
        payload: dict[str, Any] = {
            "ok": bool(ok),
            "error": str(error or "").strip(),
            "target_bin_version": str(target_bin_version or "").strip(),
            "display_version": str(display_version or "").strip(),
            "ts": _ts(),
        }
        tmp = result_path.with_suffix(result_path.suffix + ".new")
        tmp.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        os.replace(tmp, result_path)
    except OSError as e:
        try:
            _append_with_cap(
                result_path.parent / "hc_update.log",
                f"{_ts()} apply_bin: write ResultPath failed err={e!s}\n",
            )
        except OSError:
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
            root.title(self._messages.get("UPDATER_WINDOW_TITLE", "CSV Tool の更新"))
            root.attributes("-topmost", True)
            root.geometry("540x190")
            root.resizable(False, False)
            init_status = self._messages.get("UPDATER_INITIAL_STATUS", "Excel の終了を待っています")
            self._title = tk.StringVar(value=f"状態: {init_status}")
            self._msg = tk.StringVar(value=self._messages.get("UPDATER_INITIAL_MESSAGE", ""))
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

    @property
    def active(self) -> bool:
        return self._ok

    def set(self, title: str, message: str, progress: int) -> None:
        if not self._ok:
            return
        title_var = self._title
        msg_var = self._msg
        bar = self._bar
        root = self._root
        if title_var is None or msg_var is None or bar is None or root is None:
            return
        try:
            title_var.set(f"状態: {title}")
            msg_var.set(message)
            bar["value"] = max(0, min(100, int(progress)))
            root.update_idletasks()
            root.update()
        except Exception:
            pass

    def close(self) -> None:
        if not self._ok:
            return
        root = self._root
        if root is None:
            return
        try:
            root.destroy()
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


def _write_ui_ready_marker(path: Path | None, *, ui_active: bool) -> None:
    if path is None or not str(path).strip():
        return
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(
            json.dumps(
                {"ready": True, "ui_active": bool(ui_active), "pid": os.getpid(), "ts": _ts()},
                ensure_ascii=False,
            ),
            encoding="utf-8",
        )
    except OSError:
        pass


def _job_from_deferred_inline_result(
    install_root: Path,
    res: dict[str, Any],
    log_path: Path,
) -> Job:
    zip_s = str(res.get("worker_zip_path") or "").strip()
    return Job(
        install_root=install_root,
        zip_path=Path(zip_s),
        expected_sha=str(res.get("worker_zip_sha") or "").strip().lower(),
        log_path=log_path,
        display_version=str(res.get("display_version") or "").strip(),
        apply_mode=str(res.get("worker_apply_mode") or "full").strip().lower() or "full",
        target_bin_version=str(res.get("target_bin_version") or "").strip(),
        cleanup_path="",
        notify_marker_path=str(_bin_apply_success_marker_path(install_root).resolve()),
    )


def _run_apply_pending_job(job_path: Path, raw: dict[str, Any]) -> int:
    """apply_pending ジョブ: pending を適用（昇格時は inline bin、通常 defer 可）。"""
    import traceback

    install_root = Path(str(raw.get("InstallRoot", "")).strip())
    result_s = str(raw.get("ResultPath", "")).strip()
    result_path = Path(result_s) if result_s else Path()
    log_s = str(raw.get("LogPath", "")).strip()
    log_line = Path(log_s) if log_s else (install_root / "logs" / "hc_update.log")
    source_s = str(raw.get("Source", "")).strip()
    ui_ready_s = str(raw.get("UiReadyPath") or "").strip()
    ui_ready_path = Path(ui_ready_s) if ui_ready_s else None
    inline_raw = raw.get("InlineBin")
    inline_bin = True if inline_raw is None else bool(inline_raw)
    defer_ui: _ProgressUi | None = None
    handed_off_ui = False
    try:
        if not install_root.is_dir():
            raise RuntimeError(f"InstallRoot not found: {install_root}")
        _append_with_cap(
            log_line,
            f"{_ts()} apply_pending: worker start pid={os.getpid()} install_root={install_root} "
            f"inline_bin={inline_bin} source={source_s or '-'}\n",
        )
        from bootstrap.update_bootstrap import (
            apply_pending_update,
            clear_external_progress_pulse,
            set_external_progress_pulse,
        )

        ui_msgs: dict[str, str] = {}
        if not inline_bin:
            ui_msgs = _load_update_messages_for_install(install_root)
            defer_ui = _ProgressUi(ui_msgs)
            _write_ui_ready_marker(ui_ready_path, ui_active=defer_ui.active)
            if defer_ui.active:
                defer_ui.set(
                    ui_msgs.get("PROGRESS_PREPARE_TITLE", "準備中"),
                    ui_msgs.get(
                        "PROGRESS_PREPARE_MSG",
                        "更新に必要なファイルを用意しています。\nしばらくお待ちください。",
                    ),
                    5,
                )
                set_external_progress_pulse(
                    lambda title, message, progress: defer_ui.set(title, message, int(progress))
                )
            os.environ["HC_BOOTSTRAP_NO_TK"] = "1"
            os.environ["CSV_TOOL_HC_UPDATER_CONTINUOUS_BIN"] = "1"

        if inline_bin:
            os.environ["CSV_TOOL_APPLY_PENDING_INLINE_BIN"] = "1"
        else:
            os.environ.pop("CSV_TOOL_APPLY_PENDING_INLINE_BIN", None)
        try:
            res = apply_pending_update(install_root)
        finally:
            clear_external_progress_pulse()
            os.environ.pop("CSV_TOOL_APPLY_PENDING_INLINE_BIN", None)
            os.environ.pop("HC_BOOTSTRAP_NO_TK", None)
            os.environ.pop("CSV_TOOL_HC_UPDATER_CONTINUOUS_BIN", None)

        out: dict[str, Any] = {"ok": False, "applied": False}
        if isinstance(res, dict):
            out.update(res)
            out["ok"] = bool(res.get("ok", True))
        else:
            out = {"ok": True, "applied": False}

        if (
            not inline_bin
            and isinstance(res, dict)
            and res.get("deferred_inline_bin_apply")
            and defer_ui is not None
        ):
            if defer_ui.active:
                defer_ui.set(
                    ui_msgs.get("PROGRESS_DEFER_DONE_TITLE", "準備完了"),
                    ui_msgs.get(
                        "PROGRESS_DEFER_DONE_TEMPLATE",
                        "更新の準備が終わりました。\n\n開いている Microsoft Excel をすべて終了してください。",
                    ),
                    100,
                )
            job = _job_from_deferred_inline_result(install_root, res, log_line)
            bin_result_path = result_path if result_s else updater_result_path(install_root)
            handed_off_ui = True
            exit_code = _run_bin_apply_job(
                job,
                defer_ui,
                ui_msgs,
                result_path=bin_result_path if str(bin_result_path) else None,
                job_path=None,
            )
            if result_s:
                result_path.parent.mkdir(parents=True, exist_ok=True)
                result_path.write_text(
                    json.dumps(
                        {
                            "ok": exit_code == 0,
                            "applied": exit_code == 0,
                            "deferred_to_updater": False,
                            "deferred_inline_bin_apply": True,
                        },
                        ensure_ascii=False,
                    ),
                    encoding="utf-8",
                )
            _append_with_cap(
                log_line,
                f"{_ts()} apply_pending: continuous_bin_apply done exit={exit_code}\n",
            )
            return exit_code

        if defer_ui is not None:
            defer_ui.close()
        if result_s:
            result_path.parent.mkdir(parents=True, exist_ok=True)
            result_path.write_text(json.dumps(out, ensure_ascii=False), encoding="utf-8")
        _append_with_cap(
            log_line,
            f"{_ts()} apply_pending: done ok={out.get('ok')} applied={out.get('applied')} "
            f"deferred_to_updater={out.get('deferred_to_updater')}\n",
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
        if defer_ui is not None and not handed_off_ui:
            try:
                defer_ui.close()
            except Exception:
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


def _run_bin_apply_job(
    job: Job,
    ui: _ProgressUi,
    ui_msgs: dict[str, str],
    *,
    result_path: Path | None,
    job_path: Path | None = None,
) -> int:
    exit_code = 1
    result_error = ""
    busy_title = updater_busy_title(ui_msgs, job.apply_mode)
    apply_title = busy_title
    apply_msg = updater_busy_body(
        ui_msgs, job.apply_mode, "UPDATER_PHASE_APPLY_MESSAGE", "更新ファイルを適用しています。", apply_phase=True
    )
    try:
        _append_with_cap(
            job.log_path,
            f"{_ts()} apply_bin: updater job loaded pid={os.getpid()} job={job_path or '-'}\n",
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
            ui_msgs.get("UPDATER_PHASE_WAIT_TITLE", "Excel の終了を待っています"),
            ui_msgs.get("UPDATER_PHASE_WAIT_MESSAGE", ""),
            5,
        )
        wait_title = ui_msgs.get("UPDATER_PHASE_WAIT_TITLE", "Excel の終了を待っています")
        wait_msg = ui_msgs.get("UPDATER_PHASE_WAIT_MESSAGE", "")
        wait_deadline = time.time() + _excel_wait_timeout_sec(job.install_root)

        def _wait_pulse() -> None:
            ui.set(wait_title, wait_msg, 5)

        while _is_excel_running():
            if time.time() > wait_deadline:
                timeout_tpl = ui_msgs.get(
                    "UPDATER_EXCEL_WAIT_TIMEOUT_TEMPLATE",
                    "Excel が終了しませんでした。\n\nすべての Excel ウィンドウを閉じてから、再度「すぐに更新」を実行してください。",
                )
                raise RuntimeError(timeout_tpl)
            sleep_with_ui_pulse(2.0, ui_pulse=_wait_pulse)
        sleep_with_ui_pulse(1.0, ui_pulse=_wait_pulse)

        _phase(
            job.log_path,
            ui,
            "start",
            busy_title,
            updater_busy_body(
                ui_msgs,
                job.apply_mode,
                "UPDATER_PHASE_START_MESSAGE",
                "更新処理を開始しています。",
            ),
            15,
        )
        dl_tmp = Path(tempfile.mkdtemp(prefix="csv_tool_bin_download_"))
        zip_local = dl_tmp / job.zip_path.name
        _phase(
            job.log_path,
            ui,
            "download_copy",
            busy_title,
            updater_busy_body(
                ui_msgs,
                job.apply_mode,
                "UPDATER_PHASE_DOWNLOAD_MESSAGE",
                "更新ファイルを取得しています。",
            ),
            30,
        )
        shutil.copy2(job.zip_path, zip_local)

        if job.expected_sha:
            got = _sha256_file(zip_local)
            if got != job.expected_sha:
                raise RuntimeError(f"sha256 mismatch got={got}")

        extract_tmp = Path(tempfile.mkdtemp(prefix="csv_tool_bin_extract_"))
        try:
            extract_title = busy_title
            extract_msg = updater_busy_body(
                ui_msgs,
                job.apply_mode,
                "UPDATER_PHASE_EXTRACT_MESSAGE",
                "更新ファイルを展開しています。",
            )
            _phase(
                job.log_path,
                ui,
                "extract",
                extract_title,
                extract_msg,
                50,
            )
            ui.set(extract_title, extract_msg, 55)
            from bootstrap.update_bootstrap import _pulse_while_blocking

            _pulse_while_blocking(
                ui,
                extract_title,
                extract_msg,
                55,
                lambda: shutil.unpack_archive(str(zip_local), str(extract_tmp)),
            )
            cfg = load_runtime_config(job.install_root)
            stop_msg = updater_busy_body(
                ui_msgs,
                job.apply_mode,
                "UPDATER_PHASE_STOP_PROCESSES_MESSAGE",
                "関連プロセスを終了しています…",
            )

            def _stop_pulse() -> None:
                ui.set(busy_title, stop_msg, 68)

            ui.set(busy_title, stop_msg, 68)
            ensure_packaged_children_stopped(
                lambda m: _append_with_cap(job.log_path, f"{_ts()} {m}\n"),
                cfg,
                phase="updater_before_bin_apply",
                force_taskkill=True,
                ui_pulse=_stop_pulse,
            )
            taskkill_other_hc_updater_processes(
                lambda m: _append_with_cap(job.log_path, f"{_ts()} {m}\n"),
            )
            _paths = sorted(str(p) for p in _self_hc_updater_dst_paths())
            _append_with_cap(
                job.log_path,
                f"{_ts()} apply_bin: self_hc_updater_paths={_paths} executable={sys.executable!r} __file__={__file__!r}\n",
            )
            _phase(
                job.log_path,
                ui,
                "apply",
                busy_title,
                apply_msg,
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
                    _cleanup_stale_renamed_updaters(dst_bin, job.log_path)
                    _copy_merge_tree(
                        patch_app,
                        dst_bin,
                        job.log_path,
                        ui=ui,
                        ui_title=apply_title,
                        ui_message=apply_msg,
                        progress_lo=75,
                        progress_hi=88,
                    )
                if patch_addin.exists():
                    dst_addin = job.install_root / "addin"
                    dst_addin.mkdir(parents=True, exist_ok=True)
                    _copy_merge_tree(
                        patch_addin,
                        dst_addin,
                        job.log_path,
                        ui=ui,
                        ui_title=apply_title,
                        ui_message=apply_msg,
                        progress_lo=88,
                        progress_hi=95,
                    )
                _apply_delete_list(job.install_root, extract_tmp / "__delete_list.txt")
                if job.target_bin_version:
                    (job.install_root / "VERSION.txt").write_text(job.target_bin_version + "\n", encoding="utf-8")
                _append_with_cap(job.log_path, f"{_ts()} apply_bin: patch merged TargetBinVersion={job.target_bin_version}\n")
            else:
                src_bin = extract_tmp / "app" / "bin"
                if not src_bin.is_dir():
                    raise RuntimeError("invalid zip: missing app/bin")
                dst_bin_full = job.install_root / "app" / "bin"
                dst_bin_full.mkdir(parents=True, exist_ok=True)
                _cleanup_stale_renamed_updaters(dst_bin_full, job.log_path)
                _append_with_cap(
                    job.log_path,
                    f"{_ts()} apply_bin: full_apply_mode=merge dst={dst_bin_full}\n",
                )
                _copy_merge_tree(
                    src_bin,
                    dst_bin_full,
                    job.log_path,
                    ui=ui,
                    ui_title=apply_title,
                    ui_message=apply_msg,
                    progress_lo=75,
                    progress_hi=92,
                )
                _apply_delete_list(job.install_root, extract_tmp / "__delete_list.txt")
                vsrc = extract_tmp / "VERSION.txt"
                if job.target_bin_version:
                    (job.install_root / "VERSION.txt").write_text(job.target_bin_version + "\n", encoding="utf-8")
                elif vsrc.is_file():
                    shutil.copy2(vsrc, job.install_root / "VERSION.txt")
                _append_with_cap(
                    job.log_path,
                    f"{_ts()} apply_bin: full merged TargetBinVersion={job.target_bin_version or '-'}\n",
                )
                addin_src = extract_tmp / "addin"
                if addin_src.is_dir():
                    _mirror_tree(
                        addin_src,
                        job.install_root / "addin",
                        job.log_path,
                        ui=ui,
                        ui_title=apply_title,
                        ui_message=apply_msg,
                        progress_lo=92,
                        progress_hi=98,
                    )

            _try_set_display_values(job.display_version, job.log_path, job.install_root)
            _write_marker(job.notify_marker_path, job.target_bin_version, job.display_version, job.log_path)
            _append_with_cap(job.log_path, f"{_ts()} apply_bin: success\n")
            try:
                post_deferred_bin_success_housekeeping(
                    job.install_root,
                    log=lambda m: _append_with_cap(job.log_path, f"{_ts()} {m}\n"),
                )
            except Exception as e:
                _append_with_cap(
                    job.log_path,
                    f"{_ts()} apply_bin: post_success_housekeeping err={type(e).__name__}: {e}\n",
                )
            _phase(
                job.log_path,
                ui,
                "done",
                ui_msgs.get("UPDATER_PHASE_DONE_TITLE", "完了"),
                ui_msgs.get("UPDATER_PHASE_DONE_MESSAGE", "更新が完了しました。"),
                100,
            )
            _write_updater_result(
                result_path,
                ok=True,
                target_bin_version=job.target_bin_version,
                display_version=job.display_version,
            )
            succ_tpl = ui_msgs.get(
                "UPDATER_SUCCESS_MESSAGE",
                "更新が完了しました。\n\nCSV Tool 版: {target_bin}\n"
                "セットバージョン: {display_version}\n\n"
                "Microsoft Excel を起動してください。",
            )
            try:
                succ_msg = succ_tpl.format(
                    target_bin=job.target_bin_version or "-",
                    display_version=job.display_version or "-",
                )
            except (KeyError, ValueError):
                succ_msg = succ_tpl
            _message_box(
                succ_msg,
                ui_msgs.get("UPDATER_WINDOW_TITLE", "CSV Tool の更新"),
                0x40,
            )
            exit_code = 0
            return exit_code
        finally:
            shutil.rmtree(extract_tmp, ignore_errors=True)
            shutil.rmtree(dl_tmp, ignore_errors=True)
    except Exception as e:
        result_error = f"{type(e).__name__}: {e}"
        log_path = job.log_path
        wn = getattr(e, "winerror", None) if isinstance(e, OSError) else None
        en = getattr(e, "errno", None) if isinstance(e, OSError) else None
        _append_with_cap(
            log_path,
            f"{_ts()} apply_bin: ERROR type={type(e).__name__} winerror={wn} errno={en} msg={e}\n",
        )
        try:
            _append_with_cap(
                log_path,
                f"{_ts()} apply_bin: ERROR probe {probe_tasklist_line()}\n",
            )
            _append_with_cap(
                log_path,
                f"{_ts()} apply_bin: ERROR mutex {mutex_snapshot()}\n",
            )
        except Exception:
            pass
        _append_with_cap(
            log_path,
            f"{_ts()} apply_bin: ERROR traceback {traceback.format_exc()[:2000]}\n",
        )
        if job.install_root.is_dir():
            try:
                cleanup_update_payload_dir(
                    job.install_root,
                    log=lambda m: _append_with_cap(log_path, f"{_ts()} {m}\n"),
                )
            except Exception:
                pass
        err_tpl = ui_msgs.get(
            "UPDATER_ERROR_TEMPLATE",
            "CSV Tool の更新に失敗しました。\n\n{error}\n\n"
            "対処: 開いている Excel をすべて閉じてから、再度「すぐに更新」を実行してください。\n"
            "解消しない場合は再インストールを検討してください。\n\n"
            "ログ: {log_path}",
        )
        try:
            err_msg = err_tpl.format(error=e, log_path=log_path)
        except (KeyError, ValueError):
            try:
                err_msg = err_tpl.format(error=e)
            except (KeyError, ValueError):
                err_msg = f"CSV Tool の更新に失敗しました。\n\n{e}"
        _message_box(err_msg, ui_msgs.get("UPDATER_WINDOW_TITLE", "CSV Tool の更新"), 0x30)
        exit_code = 1
        return exit_code
    finally:
        result_kw: dict[str, Any] = {"ok": exit_code == 0, "error": result_error}
        if exit_code == 0:
            result_kw["target_bin_version"] = job.target_bin_version
            result_kw["display_version"] = job.display_version
        _write_updater_result(result_path, **result_kw)
        try:
            _remove_path(job.cleanup_path)
        except Exception:
            pass
        if job_path is not None:
            _remove_path(str(job_path))
        try:
            ui.close()
        except Exception:
            pass


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

    result_s = str(raw_head.get("ResultPath", "")).strip()
    result_path = Path(result_s) if result_s else None

    bootstrap_log_path = Path(tempfile.gettempdir()) / "csv_tool" / "hc_update.log"
    _append_with_cap(
        bootstrap_log_path,
        f"{_ts()} apply_bin: updater process start pid={os.getpid()} job={job_path}\n",
    )
    ui_msgs = _load_update_messages_for_job(job_path)
    ui = _ProgressUi(ui_msgs)
    try:
        job = _load_job(job_path)
    except Exception as e:
        _append_with_cap(
            bootstrap_log_path,
            f"{_ts()} apply_bin: job load failed {type(e).__name__}: {e}\n",
        )
        ui.close()
        return 1
    return _run_bin_apply_job(job, ui, ui_msgs, result_path=result_path, job_path=job_path)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--job", required=True, help="Path to updater job json")
    args = ap.parse_args()
    return run(Path(args.job))


if __name__ == "__main__":
    raise SystemExit(main())

