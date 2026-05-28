from __future__ import annotations

import hashlib
import ctypes
import json
import os
import re
import shutil
import subprocess
import tempfile
import threading
import time
import traceback
import zipfile
from pathlib import Path
from typing import Any

from core.patch_manifest import materialize_manifest_patch_zip
from core.update_state import build_paths, clear_pending, load_runtime_config, read_pending, write_pending

# 起動シーケンスで apply が二重に掛からないようにする
_APPLY_SINGLE_FLIGHT = threading.Lock()
_SYNCHRONIZE = 0x00100000
_MUTEX_NAME_UI = "Global\\HC_QT_UI_SERVER"
_MUTEX_NAME_SVC = "Global\\HC_SVC_SERVER"
_MUTEX_NAME_MAIN = "Global\\HC_MAIN_RUNNER"
_MUTEX_NAME_MAIN_LEGACY = "Global\\HC_BRIDGE_RUNNER"


class UpdateApplyError(RuntimeError):
    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = str(code or "E_UNKNOWN")


def _ts() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime())


def _append(path: Path, line: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8", errors="replace") as f:
        f.write(f"{_ts()} {line}\n")


class _ProgressUi:
    def __init__(self) -> None:
        self._ok = False
        self._cancel = False
        self._root = None
        self._title = None
        self._msg = None
        self._bar = None
        try:
            import tkinter as tk
            from tkinter import messagebox, ttk

            self._tk = tk
            self._messagebox = messagebox
            root = tk.Tk()
            root.title("CSV Tool 更新")
            root.geometry("560x210")
            root.resizable(False, False)
            root.attributes("-topmost", True)
            self._title = tk.StringVar(value="状態: 待機中")
            self._msg = tk.StringVar(value="更新準備を確認しています。")
            ttk.Label(root, textvariable=self._title).pack(anchor="w", padx=16, pady=(16, 6))
            ttk.Label(root, textvariable=self._msg).pack(anchor="w", padx=16, pady=(0, 8))
            self._bar = ttk.Progressbar(root, orient="horizontal", mode="determinate", maximum=100, length=520)
            self._bar.pack(anchor="w", padx=16, pady=(0, 6))

            def _on_close() -> None:
                yes = self._messagebox.askyesno("CSV Tool 更新", "更新を中断しますか？\n\n「いいえ」で更新を継続します。")
                if yes:
                    self._cancel = True

            root.protocol("WM_DELETE_WINDOW", _on_close)
            root.update_idletasks()
            root.update()
            self._root = root
            self._ok = True
        except Exception:
            self._ok = False

    def set(self, title: str, message: str, progress: float) -> None:
        if not self._ok:
            return
        try:
            self._title.set(f"状態: {title}")
            self._msg.set(message)
            self._bar["value"] = max(0, min(100, int(progress)))
            self._root.update_idletasks()
            self._root.update()
        except Exception:
            pass

    @property
    def cancelled(self) -> bool:
        return self._cancel

    def close(self) -> None:
        if not self._ok:
            return
        try:
            self._root.destroy()
        except Exception:
            pass

    def notify_done(self, title: str, message: str) -> None:
        if not self._ok:
            return
        try:
            self._messagebox.showinfo(title, message)
        except Exception:
            pass


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().lower()


def _phase(log_path: Path, ui: _ProgressUi, title: str, msg: str, progress: float) -> None:
    _append(log_path, f"bootstrap phase={title} message={msg}")
    ui.set(title, msg, progress)


def _is_mutex_exists(name: str) -> bool:
    try:
        kernel32 = ctypes.windll.kernel32  # type: ignore[attr-defined]
        handle = kernel32.OpenMutexW(_SYNCHRONIZE, False, name)
        if handle:
            kernel32.CloseHandle(handle)
            return True
        return False
    except Exception:
        return False


def _probe_processes() -> str:
    if os.name != "nt":
        return "tasklist=unsupported"
    try:
        cp = subprocess.run(
            ["tasklist"],
            capture_output=True,
            text=True,
            check=False,
        )
        out = (cp.stdout or "").lower()
        flags = {
            "hc_main": ("hc_main.exe" in out),
            "hc_svc_server": ("hc_svc_server.exe" in out),
            "hc_ui_server": ("hc_ui_server.exe" in out),
            "excel": ("excel.exe" in out),
        }
        return "tasklist_rc={rc} running={flags}".format(rc=cp.returncode, flags=flags)
    except Exception as e:
        return f"tasklist_probe_failed={type(e).__name__}: {e}"


def _write_shutdown_flags() -> None:
    try:
        from ui_qt import ipc_file

        ipc_file.write_shutdown_flag()
    except Exception:
        pass
    try:
        from svc.svc_host import _write_svc_shutdown_flag

        _write_svc_shutdown_flag()
    except Exception:
        pass


def _mutex_snapshot() -> dict[str, bool]:
    return {
        "main": _is_mutex_exists(_MUTEX_NAME_MAIN),
        "main_legacy": _is_mutex_exists(_MUTEX_NAME_MAIN_LEGACY),
        "svc": _is_mutex_exists(_MUTEX_NAME_SVC),
        "ui": _is_mutex_exists(_MUTEX_NAME_UI),
    }


def _wait_mutex_clear(timeout_sec: int = 20, poll_sec: float = 0.5) -> tuple[bool, dict[str, bool]]:
    t0 = time.time()
    last = _mutex_snapshot()
    while time.time() - t0 < max(1, int(timeout_sec)):
        last = _mutex_snapshot()
        if not any(last.values()):
            return True, last
        time.sleep(max(0.05, float(poll_sec)))
    return False, last


def _mode_text(mode: str) -> str:
    return "差分" if str(mode or "").strip().lower() == "patch" else "フル"


def _progress_message(base: str, *, target_bin: str, mode: str) -> str:
    tv = str(target_bin or "").strip() or "-"
    return f"{base}\n更新版: {tv}\n適用方式: {_mode_text(mode)}"


def _copy_with_progress(
    src: Path,
    dst: Path,
    log_path: Path,
    ui: _ProgressUi,
    p0: float,
    p1: float,
    *,
    progress_msg: str,
) -> None:
    total = max(src.stat().st_size, 1)
    done = 0
    dst.parent.mkdir(parents=True, exist_ok=True)
    with src.open("rb") as rf, dst.open("wb") as wf:
        while True:
            buf = rf.read(1024 * 1024)
            if not buf:
                break
            wf.write(buf)
            done += len(buf)
            pct = p0 + (p1 - p0) * (done / total)
            ui.set("取得中", progress_msg, pct)
            if ui.cancelled:
                raise UpdateApplyError("E_USER_CANCELLED", "更新はユーザーにより中断されました。")
    shutil.copystat(src, dst)
    _append(log_path, f"bootstrap copy done src={src} dst={dst} bytes={done}")


def _extract_with_progress(zip_path: Path, dst: Path, ui: _ProgressUi, p0: float, p1: float, *, progress_msg: str) -> None:
    with zipfile.ZipFile(zip_path, "r") as zf:
        infos = zf.infolist()
        total = max(len(infos), 1)
        for idx, info in enumerate(infos, start=1):
            zf.extract(info, dst)
            pct = p0 + (p1 - p0) * (idx / total)
            ui.set("展開中", progress_msg, pct)
            if ui.cancelled:
                raise UpdateApplyError("E_USER_CANCELLED", "更新はユーザーにより中断されました。")


def _apply_delete_list(install_root: Path, delete_list_path: Path) -> None:
    if not delete_list_path.is_file():
        return
    for raw in delete_list_path.read_text(encoding="utf-8-sig").splitlines():
        rel = raw.strip().replace("\\", "/").lstrip("/")
        if not rel:
            continue
        target = install_root / Path(rel)
        try:
            if target.is_file():
                target.unlink(missing_ok=True)
            elif target.is_dir():
                shutil.rmtree(target, ignore_errors=True)
        except Exception:
            pass


def _copy_tree_progress(
    src_root: Path,
    dst_root: Path,
    ui: _ProgressUi,
    title: str,
    p0: float,
    p1: float,
    *,
    progress_msg: str,
) -> None:
    files = [p for p in src_root.rglob("*") if p.is_file()]
    total = max(len(files), 1)
    for idx, fp in enumerate(files, start=1):
        rel = fp.relative_to(src_root)
        dst = dst_root / rel
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(fp, dst)
        pct = p0 + (p1 - p0) * (idx / total)
        ui.set(title, progress_msg, pct)
        if ui.cancelled:
            raise UpdateApplyError("E_USER_CANCELLED", "更新はユーザーにより中断されました。")


def _split_error(e: Exception) -> tuple[str, str]:
    if isinstance(e, UpdateApplyError):
        return e.code, str(e)
    msg = str(e) or "更新中に不明なエラーが発生しました。"
    return "E_UNKNOWN", msg


def _normalize_bootstrap_version(raw: Any) -> str:
    txt = str(raw or "").strip()
    m = re.fullmatch(r"(\d+)\.(\d+)\.(\d+)", txt)
    if not m:
        return ""
    return f"{int(m.group(1))}.{int(m.group(2))}.{int(m.group(3))}"


def _is_immediate_full_error(code: str) -> bool:
    return code in {"E_PATCH_SHA_MISMATCH", "E_PATCH_MANIFEST_INVALID"}


def _apply_zip(install_root: Path, zip_path: Path, expected_sha: str, mode: str, target_bin: str, ui: _ProgressUi, log_path: Path) -> None:
    dl_tmp = Path(tempfile.mkdtemp(prefix="csv_tool_boot_dl_"))
    ex_tmp = Path(tempfile.mkdtemp(prefix="csv_tool_boot_ex_"))
    try:
        zip_local = dl_tmp / zip_path.name
        _append(
            log_path,
            "apply_zip: start mode={m} src={src} expected_sha={sha}".format(
                m=mode,
                src=zip_path,
                sha=expected_sha or "-",
            ),
        )
        download_msg = _progress_message("更新ファイルを取得しています。", target_bin=target_bin, mode=mode)
        extract_msg = _progress_message("更新ファイルを展開しています。", target_bin=target_bin, mode=mode)
        apply_msg = _progress_message("更新を適用しています。", target_bin=target_bin, mode=mode)
        _phase(log_path, ui, "取得中", download_msg, 5)
        _copy_with_progress(zip_path, zip_local, log_path, ui, 5, 35, progress_msg=download_msg)
        if expected_sha:
            got = _sha256_file(zip_local)
            if got != expected_sha:
                if mode == "patch":
                    raise UpdateApplyError("E_PATCH_SHA_MISMATCH", "更新ファイルの整合性検証に失敗しました（sha256不一致）。")
                raise UpdateApplyError("E_FULL_SHA_MISMATCH", "更新ファイルの整合性検証に失敗しました（sha256不一致）。")
        _phase(log_path, ui, "展開中", extract_msg, 40)
        _extract_with_progress(zip_local, ex_tmp, ui, 40, 65, progress_msg=extract_msg)
        _append(log_path, f"apply_zip: extract_done mode={mode} temp={ex_tmp}")
        _phase(log_path, ui, "適用中", apply_msg, 70)
        if mode == "patch":
            p_app = ex_tmp / "app" / "bin"
            p_addin = ex_tmp / "addin"
            if not p_app.exists() and not p_addin.exists():
                raise UpdateApplyError("E_PATCH_MANIFEST_INVALID", "差分更新ファイルが不正です（必要な構成が不足しています）。")
            if p_app.exists():
                _copy_tree_progress(p_app, install_root / "app" / "bin", ui, "適用中", 70, 90, progress_msg=apply_msg)
            if p_addin.exists():
                _copy_tree_progress(p_addin, install_root / "addin", ui, "適用中", 70, 95, progress_msg=apply_msg)
            _apply_delete_list(install_root, ex_tmp / "__delete_list.txt")
            if target_bin:
                (install_root / "VERSION.txt").write_text(target_bin + "\n", encoding="utf-8")
            _append(log_path, f"apply_zip: patch_apply_done target_bin={target_bin or '-'}")
        else:
            s_bin = ex_tmp / "app" / "bin"
            if not s_bin.is_dir():
                raise UpdateApplyError("E_FULL_LAYOUT_INVALID", "更新ファイルの構成が不正です（app/bin がありません）。")
            d_bin = install_root / "app" / "bin"
            if d_bin.exists():
                shutil.rmtree(d_bin, ignore_errors=True)
            _copy_tree_progress(s_bin, d_bin, ui, "適用中", 70, 95, progress_msg=apply_msg)
            vsrc = ex_tmp / "VERSION.txt"
            if target_bin:
                (install_root / "VERSION.txt").write_text(target_bin + "\n", encoding="utf-8")
            elif vsrc.is_file():
                shutil.copy2(vsrc, install_root / "VERSION.txt")
            a_src = ex_tmp / "addin"
            if a_src.is_dir():
                d_add = install_root / "addin"
                if d_add.exists():
                    shutil.rmtree(d_add, ignore_errors=True)
                _copy_tree_progress(a_src, d_add, ui, "適用中", 70, 98, progress_msg=apply_msg)
            _append(log_path, f"apply_zip: full_apply_done target_bin={target_bin or '-'}")
    except Exception as e:
        _append(
            log_path,
            "apply_zip: fatal mode={m} type={t} err={msg}".format(
                m=mode,
                t=type(e).__name__,
                msg=e,
            ),
        )
        if isinstance(e, OSError):
            _append(
                log_path,
                "apply_zip: os_error errno={eno} winerror={wno} filename={fn}".format(
                    eno=getattr(e, "errno", None),
                    wno=getattr(e, "winerror", None),
                    fn=getattr(e, "filename", None),
                ),
            )
        raise
    finally:
        shutil.rmtree(dl_tmp, ignore_errors=True)
        shutil.rmtree(ex_tmp, ignore_errors=True)


def _try_apply_bootstrap_swap(install_root: Path, pending: dict[str, Any], log_path: Path) -> tuple[bool, str | None]:
    b = pending.get("bootstrap")
    if not isinstance(b, dict):
        return False, None
    if not bool(b.get("pending_swap", False)):
        return False, None
    new_path_s = str(b.get("local_new_path") or "").strip()
    if not new_path_s:
        return False, "bootstrap 更新ファイルが指定されていません。"
    new_path = Path(new_path_s)
    if not new_path.is_file():
        return False, "bootstrap 更新ファイルが見つかりません。"
    dst = install_root / "bootstrap" / "update_bootstrap.exe"
    bak = install_root / "bootstrap" / "update_bootstrap.exe.bak"
    try:
        dst.parent.mkdir(parents=True, exist_ok=True)
        if dst.is_file():
            try:
                bak.unlink(missing_ok=True)
            except Exception:
                pass
            os.replace(dst, bak)
        os.replace(new_path, dst)
        b["pending_swap"] = False
        b["local_new_path"] = ""
        target_version = _normalize_bootstrap_version(b.get("target_version"))
        if target_version:
            (install_root / "bootstrap" / "VERSION.txt").write_text(target_version + "\n", encoding="utf-8")
        pending["bootstrap"] = b
        try:
            bak.unlink(missing_ok=True)
        except Exception:
            pass
        _append(log_path, "bootstrap_self_update: 成功しました。")
        return True, None
    except Exception as e:
        try:
            if bak.is_file() and not dst.exists():
                os.replace(bak, dst)
        except Exception:
            pass
        _append(log_path, f"bootstrap_self_update: 失敗しました err={type(e).__name__}: {e}")
        return False, "bootstrap 自己更新に失敗しました。"


def _resolve_payload(install_root: Path, payload: dict[str, Any], catalog_path: str) -> tuple[Path | None, str]:
    local = str(payload.get("local_path") or "").strip()
    if local and Path(local).is_file():
        return Path(local), str(payload.get("sha256") or "").strip().lower()
    rel = str(payload.get("relative_path") or "").strip()
    if not rel:
        return None, ""
    p = Path(rel)
    if p.is_absolute():
        return p if p.is_file() else None, str(payload.get("sha256") or "").strip().lower()
    cat = Path(catalog_path)
    if cat.is_file():
        cand = (cat.parent / p).resolve()
        return cand if cand.is_file() else None, str(payload.get("sha256") or "").strip().lower()
    cand2 = (install_root / "update" / "payload" / p.name).resolve()
    return cand2 if cand2.is_file() else None, str(payload.get("sha256") or "").strip().lower()


def _confirm_pending_apply_before_progress(install_root: Path, pending: dict[str, Any]) -> bool:
    """予約適用の直前に操作者へ Yes/No。False=今回はスキップ（pending は残す）。"""
    try:
        import tkinter as tk
        from tkinter import messagebox
    except Exception:
        return True
    cfg_path = install_root / "config" / "ui_update_check.json"
    title = "CSV Tool 更新"
    body = (
        "予約された更新を適用しますか？\n\n"
        "目標 bin 版: {target_bin}\n"
        "適用方式: {mode_text}\n"
        "bootstrap 同梱: {has_bootstrap}\n"
        "スコープ: {apply_scope}\n\n"
        "「いいえ」は今回スキップし、次回起動で再確認します。"
    )
    try:
        if cfg_path.is_file():
            raw = json.loads(cfg_path.read_text(encoding="utf-8-sig"))
            msg = raw.get("MESSAGES") if isinstance(raw, dict) else {}
            if isinstance(msg, dict):
                title = str(msg.get("PENDING_APPLY_CONFIRM_TITLE") or title).strip() or title
                body = str(msg.get("PENDING_APPLY_CONFIRM_TEMPLATE") or body)
    except Exception:
        pass
    mode = str(pending.get("mode") or "patch").strip().lower()
    target_bin = str(pending.get("target_bin_version") or "").strip() or "-"
    b = pending.get("bootstrap") if isinstance(pending.get("bootstrap"), dict) else {}
    has_bs = "あり" if bool(b.get("pending_swap")) else "なし"
    scope = str(pending.get("apply_scope") or "").strip() or "（従来）"
    body = (
        body.replace("{target_bin}", target_bin)
        .replace("{mode_text}", _mode_text(mode))
        .replace("{has_bootstrap}", has_bs)
        .replace("{apply_scope}", scope)
    )
    root = tk.Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        return bool(messagebox.askyesno(title, body, parent=root))
    finally:
        try:
            root.destroy()
        except Exception:
            pass


def apply_pending_update(install_root: Path) -> dict[str, Any]:
    if not _APPLY_SINGLE_FLIGHT.acquire(blocking=False):
        return {"ok": True, "applied": False, "skipped": "concurrent_apply"}
    try:
        return _apply_pending_update_impl(install_root)
    finally:
        _APPLY_SINGLE_FLIGHT.release()


def _apply_pending_update_impl(install_root: Path) -> dict[str, Any]:
    paths = build_paths(install_root)
    cfg = load_runtime_config(install_root)
    pending = read_pending(paths)
    if not pending:
        return {"ok": True, "applied": False}
    if paths.lock_path.exists():
        return {"ok": False, "applied": False, "error": "apply lock exists"}
    _append(
        paths.log_path,
        "pending_apply: start apply_scope={scope} mode={mode} target_bin={target} catalog={cat}".format(
            scope=str(pending.get("apply_scope") or "-"),
            mode=str(pending.get("mode") or "-"),
            target=str(pending.get("target_bin_version") or "-"),
            cat=str(pending.get("catalog_path") or "-"),
        ),
    )
    _append(
        paths.log_path,
        "pending_apply: lock_probe main={main} main_legacy={main_old} svc={svc} ui={ui}".format(
            main=_is_mutex_exists(_MUTEX_NAME_MAIN),
            main_old=_is_mutex_exists(_MUTEX_NAME_MAIN_LEGACY),
            svc=_is_mutex_exists(_MUTEX_NAME_SVC),
            ui=_is_mutex_exists(_MUTEX_NAME_UI),
        ),
    )
    _append(paths.log_path, f"pending_apply: {_probe_processes()}")
    snap0 = _mutex_snapshot()
    if any(snap0.values()):
        _append(paths.log_path, "pending_apply: request_shutdown_flags=true")
        _write_shutdown_flags()
        ok_clear, snap1 = _wait_mutex_clear(timeout_sec=20, poll_sec=0.5)
        _append(paths.log_path, f"pending_apply: mutex_after_wait ok_clear={ok_clear} state={snap1}")
        if not ok_clear:
            return {
                "ok": False,
                "applied": False,
                "error": f"blocked_by_running_process mutex={snap1}",
            }
    if not _confirm_pending_apply_before_progress(install_root, pending):
        _append(paths.log_path, "pending_apply: user_decision=no deferred=true")
        return {"ok": True, "applied": False, "deferred": True}
    paths.lock_path.parent.mkdir(parents=True, exist_ok=True)
    paths.lock_path.write_text(str(_ts()), encoding="utf-8")
    ui = _ProgressUi()
    t0 = time.time()
    try:
        timeout_sec = max(30, int(cfg.get("BOOTSTRAP_APPLY_TIMEOUT_SEC", 120)))
        retries = max(1, int(cfg.get("PATCH_RETRY_IN_RUN_MAX", 3)))
        wait1 = max(0, int(cfg.get("PATCH_RETRY_WAIT_SEC_1", 2)))
        wait2 = max(0, int(cfg.get("PATCH_RETRY_WAIT_SEC_2", 5)))

        # bootstrap 自己更新は「1起動内のみ」で最大3回。起動またぎ累積はしない。
        p_retry = pending.get("retry") if isinstance(pending.get("retry"), dict) else {}
        p_retry["bootstrap_retry_in_run"] = 0
        pending["retry"] = p_retry
        write_pending(paths, pending)
        b_err: str | None = None
        for bi in range(3):
            p_retry = pending.get("retry") if isinstance(pending.get("retry"), dict) else {}
            p_retry["bootstrap_retry_in_run"] = bi + 1
            pending["retry"] = p_retry
            write_pending(paths, pending)
            _, b_err = _try_apply_bootstrap_swap(install_root, pending, paths.log_path)
            if not b_err:
                break
            if bi < 2:
                time.sleep(1 if bi == 0 else 2)
        if b_err:
            _append(paths.log_path, f"bootstrap_self_update: 失敗しました err={b_err}")
            p_retry = pending.get("retry") if isinstance(pending.get("retry"), dict) else {}
            p_retry["last_error_code"] = "E_BOOTSTRAP_SWAP_FAILED"
            p_retry["last_error_message"] = b_err
            p_retry["last_failed_at"] = _ts()
            p_retry["bootstrap_retry_in_run"] = 0
            pending["retry"] = p_retry
            write_pending(paths, pending)
        else:
            p_retry = pending.get("retry") if isinstance(pending.get("retry"), dict) else {}
            p_retry["bootstrap_retry_in_run"] = 0
            pending["retry"] = p_retry
            write_pending(paths, pending)

        apply_scope = str(pending.get("apply_scope") or "").strip()
        if apply_scope == "bootstrap_only":
            if b_err:
                _append(paths.log_path, f"bootstrap_only: swap aborted err={b_err}")
                return {"ok": False, "applied": False, "error": b_err or "bootstrap swap failed"}
            bt = pending.get("bootstrap") if isinstance(pending.get("bootstrap"), dict) else {}
            bt_ver = str(bt.get("target_version") or "-")
            pending["state"] = "done"
            pending["retry"] = {
                "patch_retry_in_run": 0,
                "patch_fail_total": 0,
                "full_fail_total": 0,
                "last_error_code": "",
                "last_error_message": "",
                "last_failed_at": "",
            }
            write_pending(paths, pending)
            done_msg = (
                f"bootstrap を更新しました。Excel を再起動してください。\n\n適用後 bootstrap 版: {bt_ver}"
            )
            _phase(paths.log_path, ui, "完了", done_msg, 100)
            _append(paths.log_path, f"apply_bootstrap_only: success version={bt_ver}")
            ui.notify_done(
                "CSV Tool 更新",
                "bootstrap の更新が完了しました。\n\n"
                f"適用後 bootstrap 版: {bt_ver}\n\n"
                "Excel を再起動してください。",
            )
            clear_pending(paths)
            try:
                shutil.rmtree(paths.payload_root, ignore_errors=True)
            except Exception:
                pass
            return {"ok": True, "applied": True}

        target_bin = str(pending.get("target_bin_version") or "").strip()
        mode = str(pending.get("mode") or "patch").strip().lower()
        cat_path = str(pending.get("catalog_path") or "").strip()
        retry = pending.get("retry") if isinstance(pending.get("retry"), dict) else {}
        patch_total = int(retry.get("patch_fail_total", 0) or 0)
        full_total = int(retry.get("full_fail_total", 0) or 0)

        if mode not in ("patch", "full"):
            mode = "patch"
        _append(
            paths.log_path,
            "apply_bin: apply_mode_selected={m} target_bin={t}".format(
                m=mode,
                t=target_bin or "-",
            ),
        )
        pending["state"] = "applying_patch" if mode == "patch" else "applying_full"
        write_pending(paths, pending)
        payload_patch = pending.get("patch") if isinstance(pending.get("patch"), dict) else {}
        payload_full = pending.get("full") if isinstance(pending.get("full"), dict) else {}

        def _update_err(msg: str) -> None:
            p_retry = pending.get("retry") if isinstance(pending.get("retry"), dict) else {}
            p_retry["last_error_message"] = msg
            p_retry["last_failed_at"] = _ts()
            p_retry["last_error_code"] = str(p_retry.get("last_error_code") or "")
            pending["retry"] = p_retry

        if mode == "patch":
            patch_path, patch_sha = _resolve_payload(install_root, payload_patch, cat_path)
            if patch_path is None:
                patch_total += 1
                _update_err("差分更新ファイルを取得できません。")
                pending["retry"] = {"patch_retry_in_run": 0, "patch_fail_total": patch_total, "full_fail_total": full_total, "last_error_code": "E_PATCH_PAYLOAD_MISSING", "last_error_message": pending.get("retry", {}).get("last_error_message", ""), "last_failed_at": _ts()}
                pending["state"] = "applying_full"
                mode = "full"
                write_pending(paths, pending)
            else:
                succeeded = False
                for i in range(retries):
                    if time.time() - t0 > timeout_sec:
                        raise UpdateApplyError("E_APPLY_TIMEOUT", "更新処理がタイムアウトしました。")
                    try:
                        pending["retry"] = {"patch_retry_in_run": i + 1, "patch_fail_total": patch_total, "full_fail_total": full_total, "last_error_code": "", "last_error_message": "", "last_failed_at": ""}
                        write_pending(paths, pending)
                        t_mat0 = time.perf_counter()
                        try:
                            mz, mclean, mstats, merr = materialize_manifest_patch_zip(
                                install_root=install_root,
                                patch_zip=patch_path,
                                target_bin_version=target_bin,
                            )
                        finally:
                            sec = time.perf_counter() - t_mat0
                            _tm = int(round(sec * 1000))
                            _h, _r = divmod(_tm, 3600000)
                            _m, _r = divmod(_r, 60000)
                            _s, _ms = divmod(_r, 1000)
                            msg = f"[et {_h:d}:{_m:02d}:{_s:02d}.{_ms:03d}]  patch materialize zip={patch_path.name}"
                            _append(paths.log_path, msg)
                            print(msg, flush=True)
                        if mstats is not None:
                            _append(paths.log_path, f"patch materialize ok stats={mstats}")
                        elif merr:
                            _append(
                                paths.log_path,
                                f"patch materialize note err={merr}",
                            )
                        apply_sha = "" if mstats is not None else patch_sha
                        try:
                            _apply_zip(install_root, mz, apply_sha, "patch", target_bin, ui, paths.log_path)
                        finally:
                            if mclean and str(mclean) and Path(mclean).is_dir():
                                shutil.rmtree(mclean, ignore_errors=True)
                        succeeded = True
                        break
                    except Exception as e:
                        code, msg = _split_error(e)
                        _append(paths.log_path, f"bootstrap patch failed try={i+1}/{retries} err={msg}")
                        if _is_immediate_full_error(code):
                            patch_total += 1
                            mode = "full"
                            pending["state"] = "applying_full"
                            p_retry = pending.get("retry") if isinstance(pending.get("retry"), dict) else {}
                            p_retry["last_error_code"] = code
                            pending["retry"] = p_retry
                            _update_err("更新ファイルの整合性検証または差分情報の解析に失敗しました。")
                            _append(paths.log_path, "apply_bin: fallback_patch_to_full=true reason={c}".format(c=code))
                            write_pending(paths, pending)
                            break
                        if i + 1 < retries:
                            time.sleep(wait1 if i == 0 else wait2)
                if not succeeded and mode == "patch":
                    patch_total += 1
                    mode = "full"
                    pending["state"] = "applying_full"
                    p_retry = pending.get("retry") if isinstance(pending.get("retry"), dict) else {}
                    p_retry["last_error_code"] = str(p_retry.get("last_error_code") or "E_PATCH_RETRY_EXHAUSTED")
                    pending["retry"] = p_retry
                    _update_err("差分更新の適用に失敗したためフル更新へ切り替えます。")
                    _append(paths.log_path, "apply_bin: fallback_patch_to_full=true reason=retry_exhausted")
                    write_pending(paths, pending)

        if mode == "full":
            full_path, full_sha = _resolve_payload(install_root, payload_full, cat_path)
            if full_path is None:
                _append(paths.log_path, "bootstrap full failed: フル更新ファイルを取得できません。")
                pending["state"] = "failed"
                pending["retry"] = {"patch_retry_in_run": 0, "patch_fail_total": patch_total, "full_fail_total": full_total + 1, "last_error_code": "E_FULL_PAYLOAD_MISSING", "last_error_message": "フル更新ファイルを取得できません。", "last_failed_at": _ts()}
                write_pending(paths, pending)
                _append(paths.log_path, "apply_bin: apply_mode_final={m} apply_result=failed restart_required=false".format(m=mode))
                return {"ok": False, "applied": False, "error": "フル更新ファイルを取得できません。"}
            full_apply_err: Exception | None = None
            for j in range(3):
                try:
                    _append(paths.log_path, f"apply_full: try={j+1}/3 start")
                    _apply_zip(install_root, full_path, full_sha, "full", target_bin, ui, paths.log_path)
                    full_apply_err = None
                    break
                except PermissionError as e:
                    full_apply_err = e
                    if int(getattr(e, "winerror", 0) or 0) == 32 and j < 2:
                        _append(paths.log_path, f"apply_full: retry_on_winerror32 try={j+1}/3")
                        time.sleep(1.5 if j == 0 else 3.0)
                        continue
                    break
                except Exception as e:
                    full_apply_err = e
                    break
            if full_apply_err is not None:
                code, msg = _split_error(full_apply_err)
                pending["state"] = "failed"
                pending["retry"] = {"patch_retry_in_run": 0, "patch_fail_total": patch_total, "full_fail_total": full_total + 1, "last_error_code": code, "last_error_message": msg, "last_failed_at": _ts()}
                write_pending(paths, pending)
                _append(paths.log_path, "apply_bin: apply_mode_final={m} apply_result=failed restart_required=false".format(m=mode))
                return {"ok": False, "applied": False, "error": msg}

        pending["state"] = "done"
        pending["retry"] = {"patch_retry_in_run": 0, "patch_fail_total": patch_total, "full_fail_total": full_total, "last_error_code": "", "last_error_message": "", "last_failed_at": ""}
        write_pending(paths, pending)
        done_msg = _progress_message("更新が完了しました。Excel を再起動してください。", target_bin=target_bin, mode=mode)
        _phase(paths.log_path, ui, "完了", done_msg, 100)
        _append(
            paths.log_path,
            "apply_bin: apply_mode_final={m} apply_result=success restart_required=true target_bin={t}".format(
                m=mode,
                t=target_bin or "-",
            ),
        )
        if str(cat_path or "").strip():
            try:
                from core.packaged_update import sync_uninstall_display_version_from_catalog

                if sync_uninstall_display_version_from_catalog(cat_path, install_root):
                    _append(paths.log_path, "display_version: Windows 設定用レジストリを catalog に合わせて更新しました")
                else:
                    _append(
                        paths.log_path,
                        "display_version: レジストリ同期をスキップまたは失敗（詳細はインストール先 logs/hc_update.log）",
                    )
            except Exception as e:
                _append(
                    paths.log_path,
                    "display_version: 同期処理で例外 {t}: {m}".format(
                        t=type(e).__name__,
                        m=e,
                    ),
                )
        ui.notify_done(
            "CSV Tool 更新",
            "更新が完了しました。\n\n適用後バージョン: {v}\n最終適用方式: {m}\n\nExcel を再起動してください。".format(
                v=(target_bin or "-"),
                m=_mode_text(mode),
            ),
        )
        clear_pending(paths)
        try:
            shutil.rmtree(paths.payload_root, ignore_errors=True)
        except Exception:
            pass
        return {"ok": True, "applied": True}
    except Exception as e:
        _append(
            paths.log_path,
            "pending_apply: fatal_exception type={t} err={m}".format(
                t=type(e).__name__,
                m=e,
            ),
        )
        tb = traceback.format_exc().strip().replace("\r", " ").replace("\n", " | ")
        _append(paths.log_path, f"pending_apply: traceback={tb[:1200]}")
        return {"ok": False, "applied": False, "error": f"{type(e).__name__}: {e}"}
    finally:
        try:
            paths.lock_path.unlink(missing_ok=True)
        except Exception:
            pass
        ui.close()
