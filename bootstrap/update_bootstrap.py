from __future__ import annotations

import hashlib
import ctypes
import json
import os
import re
import shutil
import tempfile
import threading
import time
import traceback
import zipfile
from pathlib import Path
from typing import Any, Callable, TypeVar, cast

from core.patch_manifest import materialize_manifest_patch_zip
from core.update_process_cleanup import (
    ensure_packaged_children_stopped,
    mutex_blocks_pending_apply,
    should_skip_mutex_gate_before_deferred_prep,
    mutex_snapshot,
    probe_tasklist_line,
    should_relax_svc_mutex_for_interactive_defer,
)
from core.update_state import build_paths, clear_pending, load_runtime_config, read_pending, write_pending
from core.win_running_file_replace import (
    collect_self_sidecar_dst_paths,
    copy_file_with_sharing_fallback,
    process_bin_dir,
    replace_via_sidecar,
)

_T = TypeVar("_T")
_ProgressPulseFn = Callable[[str, str, float], None]
_external_progress_pulse: _ProgressPulseFn | None = None


def set_external_progress_pulse(fn: _ProgressPulseFn | None) -> None:
    """hc_updater continuous 経路: bootstrap tk 無効時に defer_ui へ進捗を転送。"""
    global _external_progress_pulse
    _external_progress_pulse = fn


def clear_external_progress_pulse() -> None:
    set_external_progress_pulse(None)


def _progress_pulse(ui: _ProgressUi, title: str, msg: str, progress: float) -> None:
    if ui.active:
        try:
            ui.set(title, msg, progress)
        except Exception:
            pass
        return
    ext = _external_progress_pulse
    if ext is not None:
        try:
            ext(title, msg, progress)
        except Exception:
            pass


def _ui_pulse_fn(ui: _ProgressUi, title: str, msg: str, progress: float) -> Callable[[], None] | None:
    if not ui.active and _external_progress_pulse is None:
        return None
    return lambda: _progress_pulse(ui, title, msg, progress)


def _pulse_while_blocking(
    ui: _ProgressUi,
    title: str,
    msg: str,
    progress: float,
    fn: Callable[[], _T],
    *,
    poll_sec: float = 0.25,
) -> _T:
    """重い処理を worker スレッドで実行し、メインスレッドで tk を pump。"""
    result: list[_T] = []
    errors: list[BaseException] = []

    def _worker() -> None:
        try:
            result.append(fn())
        except BaseException as e:
            errors.append(e)

    worker = threading.Thread(target=_worker, daemon=True)
    worker.start()
    while worker.is_alive():
        _progress_pulse(ui, title, msg, progress)
        worker.join(timeout=max(0.05, float(poll_sec)))
    _progress_pulse(ui, title, msg, progress)
    if errors:
        raise errors[0]
    return result[0]


# 起動シーケンスで apply が二重に掛からないようにする
_APPLY_SINGLE_FLIGHT = threading.Lock()


class UpdateApplyError(RuntimeError):
    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = str(code or "E_UNKNOWN")


def _ts() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime())


def _as_dict(value: Any) -> dict[str, Any]:
    if isinstance(value, dict):
        return cast(dict[str, Any], value)
    return {}


def _append(path: Path, line: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a", encoding="utf-8", errors="replace") as f:
        f.write(f"{_ts()} {line}\n")


def _env_truthy(name: str) -> bool:
    return str(os.environ.get(name) or "").strip().lower() in {"1", "true", "yes", "on"}


def _try_import_tk() -> tuple[Any | None, Any | None, Any | None]:
    """Return tkinter unless HC_BOOTSTRAP_NO_TK=1.

    Release builds use Nuitka --enable-plugin=tk-inter (see build_nuitka_bootstrap.bat).
    """
    if _env_truthy("HC_BOOTSTRAP_NO_TK"):
        return None, None, None
    try:
        import tkinter as tk
        from tkinter import messagebox, ttk

        return tk, messagebox, ttk
    except BaseException:
        return None, None, None


def _win_message_box_yes_no(title: str, body: str) -> bool:
    if os.name != "nt":
        return True
    try:
        user32: Any = ctypes.windll.user32
        mb_yesno = 0x00000004
        mb_iconinfo = 0x00000040
        mb_setforeground = 0x00010000
        mb_topmost = 0x00040000
        id_yes = 6
        rc = user32.MessageBoxW(None, body, title, mb_yesno | mb_iconinfo | mb_setforeground | mb_topmost)
        return int(rc) == id_yes
    except Exception:
        return True


def _win_message_box_ok(title: str, body: str, *, icon_error: bool = True) -> None:
    if os.name != "nt":
        return
    try:
        user32: Any = ctypes.windll.user32
        flags = 0x00000000 | 0x00010000 | 0x00040000  # OK | SETFOREGROUND | TOPMOST
        if icon_error:
            flags |= 0x00000010  # MB_ICONERROR
        else:
            flags |= 0x00000040  # MB_ICONINFORMATION
        user32.MessageBoxW(None, body, title, flags)
    except Exception:
        pass


def _notify_apply_failure(install_root: Path, error: str, *, log_path: Path | None = None) -> None:
    """apply_pending_update 失敗時に操作者へ通知（Tk 不可時は Win32 MessageBox）。"""
    err = str(error or "").strip() or "更新に失敗しました。"
    lp = log_path
    if lp is None:
        lp = install_root / "logs" / "hc_update.log"
    try:
        _append(lp, f"notify_apply_failure: {err}")
    except Exception:
        pass
    title = _ui_update_message(install_root, "UPDATER_WINDOW_TITLE", "CSV Tool の更新")
    tpl = _ui_update_message(
        install_root,
        "UPDATER_ERROR_TEMPLATE",
        "CSV Tool の更新に失敗しました。\n\n{error}",
    )
    try:
        body = tpl.format(error=err, log_path=str(lp))
    except Exception:
        body = f"CSV Tool の更新に失敗しました。\n\n{err}\n\nログ: {lp}"
    _win_message_box_ok(title, body, icon_error=True)


def _notify_apply_result_if_failed(install_root: Path, res: dict[str, Any]) -> None:
    if res.get("deferred") or res.get("skipped"):
        return
    if res.get("deferred_to_updater"):
        return
    if res.get("ok", True):
        return
    err = str(res.get("error") or "").strip() or "更新に失敗しました。"
    paths = build_paths(install_root)
    _notify_apply_failure(install_root, err, log_path=paths.log_path)


class _ProgressUi:
    def __init__(self) -> None:
        self._ok = False
        self._cancel = False
        self._root = None
        self._title = None
        self._msg = None
        self._bar = None
        try:
            tk, messagebox, ttk = _try_import_tk()
            if tk is None or messagebox is None or ttk is None:
                return

            self._tk = tk
            self._messagebox = messagebox
            root = tk.Tk()
            root.title("CSV Tool の更新")
            root.geometry("560x210")
            root.resizable(False, False)
            root.attributes("-topmost", True)
            self._title = tk.StringVar(value="状態: 準備中")
            self._msg = tk.StringVar(
                value="更新に必要なファイルを用意しています。\nしばらくお待ちください。"
            )
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
        except BaseException:
            self._ok = False

    @property
    def active(self) -> bool:
        return self._ok

    def set(self, title: str, message: str, progress: float) -> None:
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

    @property
    def cancelled(self) -> bool:
        return self._cancel

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


def _sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest().lower()


def _phase(log_path: Path, ui: _ProgressUi, title: str, msg: str, progress: float) -> None:
    _append(log_path, f"bootstrap phase={title} message={msg}")
    _progress_pulse(ui, title, msg, progress)


def _ui_pulse(ui: _ProgressUi, title: str, msg: str, progress: float) -> None:
    """長時間ブロック処理の前後で進捗 UI を更新（無応答に見えるのを軽減）。"""
    _progress_pulse(ui, title, msg, progress)


def _mode_text(mode: str) -> str:
    return "差分" if str(mode or "").strip().lower() == "patch" else "フル"


def _ui_update_message(install_root: Path, key: str, default: str) -> str:
    try:
        cfg_path = install_root / "config" / "ui_update_check.json"
        raw_obj = json.loads(cfg_path.read_text(encoding="utf-8-sig"))
        if isinstance(raw_obj, dict):
            raw = cast(dict[str, Any], raw_obj)
            msgs = _as_dict(raw.get("MESSAGES"))
            if key in msgs and isinstance(msgs[key], str):
                return msgs[key]
    except Exception:
        pass
    return default


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
    log_path: Path,
) -> None:
    files = [p for p in src_root.rglob("*") if p.is_file()]
    total = max(len(files), 1)
    proactive: set[Path] = set()
    try:
        pbin = process_bin_dir()
        if pbin is not None and dst_root.resolve() == pbin.resolve():
            proactive = collect_self_sidecar_dst_paths()
    except OSError:
        proactive = set()

    def _log(msg: str) -> None:
        _append(log_path, msg)

    for idx, fp in enumerate(files, start=1):
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
            _append(
                log_path,
                "copy_tree: copy_failed rel={rel} src={src} dst={dst} errno={eno} winerror={wno} err={msg}".format(
                    rel=rel.as_posix(),
                    src=fp,
                    dst=dst,
                    eno=getattr(e, "errno", None),
                    wno=getattr(e, "winerror", None),
                    msg=e,
                ),
            )
            raise
        pct = p0 + (p1 - p0) * (idx / total)
        ui.set(title, progress_msg, pct)
        if ui.cancelled:
            raise UpdateApplyError("E_USER_CANCELLED", "更新はユーザーにより中断されました。")


def _split_error(e: Exception) -> tuple[str, str]:
    if isinstance(e, UpdateApplyError):
        return e.code, str(e)
    base = str(e) or "更新中に不明なエラーが発生しました。"
    hint = ""
    if isinstance(e, OSError) and int(getattr(e, "winerror", 0) or 0) == 32:
        hint = " ヒント: ファイルが他プロセスに使用中です。すべての Excel を終了してから再度お試しください。"
    msg = base + hint
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
        cfg_apply = load_runtime_config(install_root)
        _ui_pulse(ui, "処理中", "関連プロセスを終了しています…", 66)
        ensure_packaged_children_stopped(
            lambda m: _append(log_path, m),
            cfg_apply,
            phase="before_bin_apply",
            force_taskkill=True,
            ui_pulse=_ui_pulse_fn(ui, "処理中", "関連プロセスを終了しています…", 66),
        )
        _phase(log_path, ui, "適用中", apply_msg, 70)
        if mode == "patch":
            p_app = ex_tmp / "app" / "bin"
            p_addin = ex_tmp / "addin"
            if not p_app.exists() and not p_addin.exists():
                raise UpdateApplyError("E_PATCH_MANIFEST_INVALID", "差分更新ファイルが不正です（必要な構成が不足しています）。")
            if p_app.exists():
                _copy_tree_progress(
                    p_app,
                    install_root / "app" / "bin",
                    ui,
                    "適用中",
                    70,
                    90,
                    progress_msg=apply_msg,
                    log_path=log_path,
                )
            if p_addin.exists():
                _copy_tree_progress(
                    p_addin,
                    install_root / "addin",
                    ui,
                    "適用中",
                    70,
                    95,
                    progress_msg=apply_msg,
                    log_path=log_path,
                )
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
                _append(log_path, f"apply_zip: rmtree before full bin replace path={d_bin}")
                shutil.rmtree(d_bin, ignore_errors=True)
            _copy_tree_progress(
                s_bin,
                d_bin,
                ui,
                "適用中",
                70,
                95,
                progress_msg=apply_msg,
                log_path=log_path,
            )
            vsrc = ex_tmp / "VERSION.txt"
            if target_bin:
                (install_root / "VERSION.txt").write_text(target_bin + "\n", encoding="utf-8")
            elif vsrc.is_file():
                shutil.copy2(vsrc, install_root / "VERSION.txt")
            a_src = ex_tmp / "addin"
            if a_src.is_dir():
                d_add = install_root / "addin"
                if d_add.exists():
                    _append(log_path, f"apply_zip: rmtree before full addin replace path={d_add}")
                    shutil.rmtree(d_add, ignore_errors=True)
                _copy_tree_progress(
                    a_src,
                    d_add,
                    ui,
                    "適用中",
                    70,
                    98,
                    progress_msg=apply_msg,
                    log_path=log_path,
                )
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
        wno = getattr(e, "winerror", None)
        if wno == 32:
            _append(log_path, f"apply_zip: winerror32_context probe={probe_tasklist_line()}")
            _append(log_path, f"apply_zip: winerror32_context mutex={mutex_snapshot()}")
        tb = traceback.format_exc().strip().replace("\r", " ").replace("\n", " | ")
        _append(log_path, f"apply_zip: traceback={tb[:2000]}")
        raise
    finally:
        shutil.rmtree(dl_tmp, ignore_errors=True)
        shutil.rmtree(ex_tmp, ignore_errors=True)


def _try_apply_bootstrap_swap(install_root: Path, pending: dict[str, Any], log_path: Path) -> tuple[bool, str | None]:
    b = _as_dict(pending.get("bootstrap"))
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
            try:
                os.replace(dst, bak)
            except OSError as e:
                if int(getattr(e, "winerror", 0) or 0) == 32:
                    replace_via_sidecar(
                        new_path,
                        dst,
                        log=lambda m: _append(log_path, f"bootstrap_self_update: {m}"),
                        label="bootstrap_exe_sidecar",
                    )
                    b["pending_swap"] = False
                    b["local_new_path"] = ""
                    target_version = _normalize_bootstrap_version(b.get("target_version"))
                    if target_version:
                        (install_root / "bootstrap" / "VERSION.txt").write_text(
                            target_version + "\n", encoding="utf-8"
                        )
                    pending["bootstrap"] = b
                    try:
                        bak.unlink(missing_ok=True)
                    except Exception:
                        pass
                    try:
                        new_path.unlink(missing_ok=True)
                    except OSError:
                        pass
                    _append(log_path, "bootstrap_self_update: 成功しました（sidecar）。")
                    return True, None
                raise
            os.replace(new_path, dst)
        else:
            os.replace(new_path, dst)
        try:
            new_path.unlink(missing_ok=True)
        except OSError:
            pass
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


def _catalog_display_version_for_pending(catalog_path: str) -> str:
    try:
        p = Path(str(catalog_path or "").strip())
        if not p.is_file():
            return ""
        from core.packaged_update import load_catalog

        data = load_catalog(p)
        if isinstance(data, dict):
            return str(data.get("set_version") or "").strip()
    except Exception:
        pass
    return ""


def _persist_zip_for_hc_updater(src: Path, payload_root: Path) -> Path:
    payload_root.mkdir(parents=True, exist_ok=True)
    dest = payload_root / "pending_hc_updater_bin.zip"
    shutil.copy2(src, dest)
    return dest


def _confirm_pending_apply_before_progress(install_root: Path, pending: dict[str, Any]) -> bool:
    """予約適用の直前に操作者へ Yes/No。False=今回はスキップ（pending は残す）。"""
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
            raw_obj = json.loads(cfg_path.read_text(encoding="utf-8-sig"))
            raw = cast(dict[str, Any], raw_obj) if isinstance(raw_obj, dict) else {}
            msg = _as_dict(raw.get("MESSAGES"))
            if msg:
                title = str(msg.get("PENDING_APPLY_CONFIRM_TITLE") or title).strip() or title
                body = str(msg.get("PENDING_APPLY_CONFIRM_TEMPLATE") or body)
    except Exception:
        pass
    mode = str(pending.get("mode") or "patch").strip().lower()
    target_bin = str(pending.get("target_bin_version") or "").strip() or "-"
    b = _as_dict(pending.get("bootstrap"))
    has_bs = "あり" if bool(b.get("pending_swap")) else "なし"
    scope = str(pending.get("apply_scope") or "").strip() or "（従来）"
    body = (
        body.replace("{target_bin}", target_bin)
        .replace("{mode_text}", _mode_text(mode))
        .replace("{has_bootstrap}", has_bs)
        .replace("{apply_scope}", scope)
    )
    tk, messagebox, _ttk = _try_import_tk()
    if tk is None or messagebox is None:
        return _win_message_box_yes_no(title, body)
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
        res = _apply_pending_update_impl(install_root)
        _notify_apply_result_if_failed(install_root, res)
        return res
    finally:
        _APPLY_SINGLE_FLIGHT.release()


def _apply_pending_update_impl(install_root: Path) -> dict[str, Any]:
    paths = build_paths(install_root)
    cfg = load_runtime_config(install_root)
    pending_opt = read_pending(paths)
    if not pending_opt:
        return {"ok": True, "applied": False}
    pending = pending_opt
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
    snap_probe = mutex_snapshot()
    _append(
        paths.log_path,
        "pending_apply: lock_probe main={main} main_legacy={main_old} svc={svc} ui={ui}".format(
            main=snap_probe["main"],
            main_old=snap_probe["main_legacy"],
            svc=snap_probe["svc"],
            ui=snap_probe["ui"],
        ),
    )
    _append(paths.log_path, f"pending_apply: {probe_tasklist_line()}")
    # skip_apply_confirm は直後に pending から除去するため、relax 判定は pop より前に行う。
    relax_svc_self = should_relax_svc_mutex_for_interactive_defer(pending)
    skip_mutex_gate = should_skip_mutex_gate_before_deferred_prep(pending)
    if pending.get("skip_apply_confirm"):
        _append(
            paths.log_path,
            "pending_apply: skip_apply_confirm=True (ribbon single confirm); second dialog suppressed",
        )
        p_skip = dict(pending)
        p_skip.pop("skip_apply_confirm", None)
        write_pending(paths, p_skip)
        pending = p_skip
    elif not _confirm_pending_apply_before_progress(install_root, pending):
        _append(paths.log_path, "pending_apply: user_decision=no deferred=true")
        return {"ok": True, "applied": False, "deferred": True}
    paths.lock_path.parent.mkdir(parents=True, exist_ok=True)
    paths.lock_path.write_text(str(_ts()), encoding="utf-8")
    ui = _ProgressUi()
    if not ui.active:
        _append(
            paths.log_path,
            "progress_ui: tk unavailable (HC_BOOTSTRAP_NO_TK or tk init failed); progress bar disabled, log only",
        )
    wait_title = _ui_update_message(
        install_root, "UPDATER_PHASE_WAIT_TITLE", "Excel の終了を待っています"
    )
    wait_msg = _ui_update_message(
        install_root, "UPDATER_PHASE_WAIT_MESSAGE", "すべての Excel を閉じてください。"
    )
    if skip_mutex_gate:
        _append(
            paths.log_path,
            "pending_apply: skip_mutex_gate=True (interactive defer prep; Excel may stay open)",
        )
    else:
        _phase(paths.log_path, ui, wait_title, wait_msg, 2)
        if relax_svc_self:
            _append(
                paths.log_path,
                "pending_apply: mutex_gate relax_svc_self=True "
                "(interactive defer from hc_svc_server; svc mutex ignored)",
            )
        snap0 = mutex_snapshot()
        if mutex_blocks_pending_apply(snap0, relax_svc_self=relax_svc_self):
            stop_msg = _ui_update_message(
                install_root,
                "UPDATER_PHASE_STOP_PROCESSES_MESSAGE",
                "関連プロセスを終了しています…",
            )
            _ui_pulse(ui, wait_title, stop_msg, 3)
            _append(paths.log_path, "pending_apply: mutex_busy graceful_shutdown")
            ensure_packaged_children_stopped(
                lambda m: _append(paths.log_path, m),
                cfg,
                phase="pending_apply_graceful",
                force_taskkill=False,
                ui_pulse=_ui_pulse_fn(ui, wait_title, stop_msg, 3),
            )
            snap_mid = mutex_snapshot()
            if mutex_blocks_pending_apply(snap_mid, relax_svc_self=relax_svc_self):
                _append(paths.log_path, "pending_apply: mutex still busy; taskkill")
                _ui_pulse(ui, wait_title, stop_msg, 4)
                ensure_packaged_children_stopped(
                    lambda m: _append(paths.log_path, m),
                    cfg,
                    phase="pending_apply_force",
                    force_taskkill=True,
                    ui_pulse=_ui_pulse_fn(ui, wait_title, stop_msg, 4),
                )
            snap1 = mutex_snapshot()
            if mutex_blocks_pending_apply(snap1, relax_svc_self=relax_svc_self):
                ui.close()
                return {
                    "ok": False,
                    "applied": False,
                    "error": f"blocked_by_running_process mutex={snap1}",
                }
    _phase(
        paths.log_path,
        ui,
        _ui_update_message(install_root, "PROGRESS_PREPARE_TITLE", "準備中"),
        _ui_update_message(
            install_root,
            "PROGRESS_PREPARE_MSG",
            "更新に必要なファイルを用意しています。\nしばらくお待ちください。",
        ),
        5,
    )
    t0 = time.time()
    try:
        timeout_sec = max(30, int(cfg.get("BOOTSTRAP_APPLY_TIMEOUT_SEC", 120)))
        retries = max(1, int(cfg.get("PATCH_RETRY_IN_RUN_MAX", 3)))
        wait1 = max(0, int(cfg.get("PATCH_RETRY_WAIT_SEC_1", 2)))
        wait2 = max(0, int(cfg.get("PATCH_RETRY_WAIT_SEC_2", 5)))

        # bootstrap 自己更新は「1起動内のみ」で最大3回。起動またぎ累積はしない。
        p_retry = _as_dict(pending.get("retry"))
        p_retry["bootstrap_retry_in_run"] = 0
        pending["retry"] = p_retry
        write_pending(paths, pending)
        b_err: str | None = None
        for bi in range(3):
            p_retry = _as_dict(pending.get("retry"))
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
            p_retry = _as_dict(pending.get("retry"))
            p_retry["last_error_code"] = "E_BOOTSTRAP_SWAP_FAILED"
            p_retry["last_error_message"] = b_err
            p_retry["last_failed_at"] = _ts()
            p_retry["bootstrap_retry_in_run"] = 0
            pending["retry"] = p_retry
            write_pending(paths, pending)
        else:
            p_retry = _as_dict(pending.get("retry"))
            p_retry["bootstrap_retry_in_run"] = 0
            pending["retry"] = p_retry
            write_pending(paths, pending)

        apply_scope = str(pending.get("apply_scope") or "").strip()
        if apply_scope == "bootstrap_only":
            if b_err:
                _append(paths.log_path, f"bootstrap_only: swap aborted err={b_err}")
                return {"ok": False, "applied": False, "error": b_err or "bootstrap swap failed"}
            bt = _as_dict(pending.get("bootstrap"))
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
            clear_pending(paths)
            try:
                shutil.rmtree(paths.payload_root, ignore_errors=True)
            except Exception:
                pass
            return {"ok": True, "applied": True}

        target_bin = str(pending.get("target_bin_version") or "").strip()
        mode = str(pending.get("mode") or "patch").strip().lower()
        cat_path = str(pending.get("catalog_path") or "").strip()
        retry = _as_dict(pending.get("retry"))
        patch_total = int(retry.get("patch_fail_total", 0) or 0)
        full_total = int(retry.get("full_fail_total", 0) or 0)
        # 速度優先: 旧版 full_prev バックアップは作らない（差分→フルで新版適用に一本化）。

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
        payload_patch = _as_dict(pending.get("patch"))
        payload_full = _as_dict(pending.get("full"))
        worker_zip: Path | None = None
        worker_mode = "full"
        worker_sha = ""
        defer_bin_to_updater = os.environ.get("CSV_TOOL_APPLY_PENDING_INLINE_BIN") != "1"
        _append(
            paths.log_path,
            "apply_bin: defer_bin_to_updater={d} (inline_bin env set={e})".format(
                d=defer_bin_to_updater,
                e=os.environ.get("CSV_TOOL_APPLY_PENDING_INLINE_BIN", ""),
            ),
        )

        def _update_err(msg: str) -> None:
            p_retry = _as_dict(pending.get("retry"))
            p_retry["last_error_message"] = msg
            p_retry["last_failed_at"] = _ts()
            p_retry["last_error_code"] = str(p_retry.get("last_error_code") or "")
            pending["retry"] = p_retry

        if mode == "patch":
            patch_path, patch_sha = _resolve_payload(install_root, payload_patch, cat_path)
            if patch_path is None:
                patch_total += 1
                _update_err("差分更新ファイルを取得できません。")
                pending["retry"] = {
                    "patch_retry_in_run": 0,
                    "patch_fail_total": patch_total,
                    "full_fail_total": full_total,
                    "last_error_code": "E_PATCH_PAYLOAD_MISSING",
                    "last_error_message": str(_as_dict(pending.get("retry")).get("last_error_message") or ""),
                    "last_failed_at": _ts(),
                }
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
                        mat_title = _ui_update_message(
                            install_root, "PROGRESS_PREPARE_TITLE", "準備中"
                        )
                        mat_msg = _ui_update_message(
                            install_root,
                            "PROGRESS_PREPARE_MSG_PATCH_BUILD",
                            "差分パッケージを構築しています。\nしばらくお待ちください。",
                        )
                        _ui_pulse(ui, mat_title, mat_msg, 8)
                        t_mat0 = time.perf_counter()
                        try:
                            mz, mclean, mstats, merr = _pulse_while_blocking(
                                ui,
                                mat_title,
                                mat_msg,
                                8,
                                lambda: materialize_manifest_patch_zip(
                                    install_root=install_root,
                                    patch_zip=patch_path,
                                    target_bin_version=target_bin,
                                ),
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
                            if defer_bin_to_updater:
                                worker_zip = _pulse_while_blocking(
                                    ui,
                                    mat_title,
                                    mat_msg,
                                    10,
                                    lambda: _persist_zip_for_hc_updater(mz, paths.payload_root),
                                )
                                worker_mode = "patch"
                                worker_sha = apply_sha
                            else:
                                _apply_zip(install_root, mz, apply_sha, "patch", target_bin, ui, paths.log_path)
                        finally:
                            if mclean and str(mclean) and Path(mclean).is_dir():
                                shutil.rmtree(mclean, ignore_errors=True)
                        succeeded = True
                        break
                    except Exception as e:
                        code, msg = _split_error(e)
                        _append(
                            paths.log_path,
                            "bootstrap patch failed try={i}/{ret} code={c} type={t} winerror={w} errno={en} err={m}".format(
                                i=i + 1,
                                ret=retries,
                                c=code,
                                t=type(e).__name__,
                                w=getattr(e, "winerror", None) if isinstance(e, OSError) else None,
                                en=getattr(e, "errno", None) if isinstance(e, OSError) else None,
                                m=msg,
                            ),
                        )
                        if _is_immediate_full_error(code):
                            patch_total += 1
                            mode = "full"
                            pending["state"] = "applying_full"
                            p_retry = _as_dict(pending.get("retry"))
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
                    p_retry = _as_dict(pending.get("retry"))
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
            if defer_bin_to_updater:
                persist_title = _ui_update_message(
                    install_root, "PROGRESS_PREPARE_TITLE", "準備中"
                )
                persist_msg = _ui_update_message(
                    install_root,
                    "PROGRESS_PREPARE_MSG",
                    "更新に必要なファイルを用意しています。\nしばらくお待ちください。",
                )
                worker_zip = _pulse_while_blocking(
                    ui,
                    persist_title,
                    persist_msg,
                    12,
                    lambda: _persist_zip_for_hc_updater(full_path, paths.payload_root),
                )
                worker_mode = "full"
                worker_sha = full_sha
            else:
                full_apply_err: Exception | None = None
                for j in range(3):
                    try:
                        _append(paths.log_path, f"apply_full: try={j+1}/3 start")
                        _apply_zip(install_root, full_path, full_sha, "full", target_bin, ui, paths.log_path)
                        full_apply_err = None
                        break
                    except PermissionError as e:
                        full_apply_err = e
                        _append(
                            paths.log_path,
                            "apply_full: caught PermissionError try={t}/3 winerror={w} errno={en} msg={m}".format(
                                t=j + 1,
                                w=getattr(e, "winerror", None),
                                en=getattr(e, "errno", None),
                                m=e,
                            ),
                        )
                        if int(getattr(e, "winerror", 0) or 0) == 32 and j < 2:
                            _append(paths.log_path, f"apply_full: retry_on_winerror32 try={j+1}/3")
                            ensure_packaged_children_stopped(
                                lambda m: _append(paths.log_path, m),
                                cfg,
                                phase=f"full_apply_win32_retry_{j}",
                                force_taskkill=True,
                                ui_pulse=_ui_pulse_fn(
                                    ui,
                                    "処理中",
                                    "関連プロセスを終了しています…",
                                    66,
                                ),
                            )
                            time.sleep(1.5 if j == 0 else 3.0)
                            continue
                        break
                    except Exception as e:
                        full_apply_err = e
                        _append(
                            paths.log_path,
                            "apply_full: caught_exception try={t}/3 type={ty} winerror={w} errno={en} msg={m}".format(
                                t=j + 1,
                                ty=type(e).__name__,
                                w=getattr(e, "winerror", None) if isinstance(e, OSError) else None,
                                en=getattr(e, "errno", None) if isinstance(e, OSError) else None,
                                m=e,
                            ),
                        )
                        break
                if full_apply_err is not None:
                    code, msg = _split_error(full_apply_err)
                    _append(
                        paths.log_path,
                        "apply_full: failed_final code={c} type={t} winerror={w} errno={en} detail={m}".format(
                            c=code,
                            t=type(full_apply_err).__name__,
                            w=getattr(full_apply_err, "winerror", None)
                            if isinstance(full_apply_err, OSError)
                            else None,
                            en=getattr(full_apply_err, "errno", None)
                            if isinstance(full_apply_err, OSError)
                            else None,
                            m=msg,
                        ),
                    )
                    pending["state"] = "failed"
                    pending["retry"] = {"patch_retry_in_run": 0, "patch_fail_total": patch_total, "full_fail_total": full_total + 1, "last_error_code": code, "last_error_message": msg, "last_failed_at": _ts()}
                    write_pending(paths, pending)
                    _append(paths.log_path, "apply_bin: apply_mode_final={m} apply_result=failed restart_required=false".format(m=mode))
                    return {"ok": False, "applied": False, "error": msg}

        if defer_bin_to_updater:
            if worker_zip is None or not worker_zip.is_file():
                _append(paths.log_path, "apply_bin: deferred_to_updater aborted worker_zip missing")
                pending["state"] = "failed"
                pending["retry"] = {
                    "patch_retry_in_run": 0,
                    "patch_fail_total": patch_total,
                    "full_fail_total": full_total + 1,
                    "last_error_code": "E_WORKER_ZIP_MISSING",
                    "last_error_message": "hc_updater 用の zip を用意できませんでした。",
                    "last_failed_at": _ts(),
                }
                write_pending(paths, pending)
                return {
                    "ok": False,
                    "applied": False,
                    "error": "更新 zip の準備に失敗しました。ログを確認してください。",
                }
            display_ver = _catalog_display_version_for_pending(cat_path)
            from core.update_process_cleanup import is_hc_updater_process

            continuous_bin = is_hc_updater_process() or os.environ.get(
                "CSV_TOOL_HC_UPDATER_CONTINUOUS_BIN"
            ) == "1"
            if continuous_bin:
                clear_pending(paths)
                _append(
                    paths.log_path,
                    "apply_bin: apply_mode_final={m} apply_result=deferred_inline_bin_apply target_bin={t}".format(
                        m=worker_mode,
                        t=target_bin or "-",
                    ),
                )
                return {
                    "ok": True,
                    "applied": False,
                    "deferred_to_updater": True,
                    "deferred_inline_bin_apply": True,
                    "worker_zip_path": str(worker_zip.resolve()),
                    "worker_zip_sha": worker_sha,
                    "worker_apply_mode": worker_mode,
                    "target_bin_version": target_bin,
                    "display_version": display_ver,
                }
            try:
                pending = read_pending(paths) or {}
                pending.setdefault("retry", {})
                pending["retry"]["last_error_code"] = "E_LEGACY_DEFER_UNSUPPORTED"
                pending["retry"]["last_error_message"] = (
                    "この更新経路はサポートされなくなりました。"
                    "リボンの「更新確認」から「すぐに更新」を実行してください。"
                )
                write_pending(paths, pending)
            except Exception as e:
                _append(
                    paths.log_path,
                    "apply_bin: legacy_defer_pending_update_failed type={t} err={m}".format(
                        t=type(e).__name__,
                        m=e,
                    ),
                )
            _append(
                paths.log_path,
                "apply_bin: legacy_defer_spawn_blocked=true reason=not_hc_updater_continuous "
                "apply_mode={m} target_bin={t} hint=use_ribbon_update_check".format(
                    m=worker_mode,
                    t=target_bin or "-",
                ),
            )
            return {
                "ok": False,
                "applied": False,
                "error": (
                    "この更新経路はサポートされなくなりました。"
                    "リボンの「更新確認」から「すぐに更新」を実行してください。"
                ),
                "error_code": "E_LEGACY_DEFER_UNSUPPORTED",
            }

        pending["state"] = "done"
        pending["retry"] = {"patch_retry_in_run": 0, "patch_fail_total": patch_total, "full_fail_total": full_total, "last_error_code": "", "last_error_message": "", "last_failed_at": ""}
        write_pending(paths, pending)
        done_msg = _ui_update_message(
            install_root,
            "PROGRESS_INLINE_DONE_MSG",
            "更新が完了しました。",
        )
        inline_title = _ui_update_message(install_root, "PROGRESS_INLINE_DONE_TITLE", "完了")
        _phase(paths.log_path, ui, inline_title, done_msg, 100)
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
        if b_err:
            _append(paths.log_path, "apply_bin: bootstrap swap not fully successful; pending cleared after bin apply")
        clear_pending(paths)
        try:
            shutil.rmtree(paths.payload_root, ignore_errors=True)
        except Exception:
            pass
        return {"ok": True, "applied": True}
    except Exception as e:
        _append(
            paths.log_path,
            "pending_apply: fatal_exception type={t} winerror={w} errno={en} err={m}".format(
                t=type(e).__name__,
                w=getattr(e, "winerror", None) if isinstance(e, OSError) else None,
                en=getattr(e, "errno", None) if isinstance(e, OSError) else None,
                m=e,
            ),
        )
        try:
            _append(paths.log_path, f"pending_apply: fatal_probe {probe_tasklist_line()}")
            _append(paths.log_path, f"pending_apply: fatal_mutex {mutex_snapshot()}")
        except Exception:
            pass
        tb = traceback.format_exc().strip().replace("\r", " ").replace("\n", " | ")
        _append(paths.log_path, f"pending_apply: traceback={tb[:2000]}")
        return {"ok": False, "applied": False, "error": f"{type(e).__name__}: {e}"}
    finally:
        try:
            paths.lock_path.unlink(missing_ok=True)
        except Exception:
            pass
        ui.close()
