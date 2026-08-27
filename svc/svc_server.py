# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: svc/svc_server.py
Created: 2026-02-15
Updated: 2026-07-03
Version: 0.1.22
Purpose:
  別プロセスで常駐する svc サーバ。
  - hc_main からの svc_req_*.pkl を監視し、action に応じて svc/hc_<feature>.py を遅延 import して実行する。
  - import 済み handler はプロセス生存中キャッシュし、次回以降の import コストを排除する（目的）。
  - UI は ui_qt/ui_server.py（既存）に委譲し、IPC契約は維持する。

Shutdown (L1):
  - control/svc_shutdown.flag を検知したら停止する。

History (latest 3):
  - 0.1.22 (2026-07-03): resolve_fresh_book_after_ui_wait — UI 待ち後は古い Book 参照を使わず HWND 再取得（COM_NG 根本対策）。
  - 0.1.21 (2026-07-03): attach_book — UI/長処理 action は cache 無効化＋HWND 再取得。com_error 後 PyErr_Clear・1 回リトライ。操作後キャッシュ破棄。
  - 0.1.20 (2026-06-16): com_monitor 掃除時に svc_last_com_hwnd.txt を同期。warmup リスト見直しと整合。
"""

import importlib
import tempfile
import pickle
import os
import sys
import time
import traceback
import subprocess
import multiprocessing as _mp
import threading

from pathlib import Path

# ctypes より前に shared を登録すると、以降に読み込むネイティブ依存の探索順が安定しやすい。
if os.name == "nt":
    from core.shared_dll_bootstrap import ensure_shared_dll_search_path_for_layout

    ensure_shared_dll_search_path_for_layout(Path(sys.executable).resolve().parent)

from ctypes import wintypes
import ctypes
from typing import Any, Callable

from core import core_env
from core.core_log import get_logger

__version__ = "0.1.20"

logger = get_logger(__name__)

# ===== DEBUG: spawn detector (safe for asyncio) =====
_OrigPopen = subprocess.Popen  # keep original class

class LoggedPopen(_OrigPopen):  # keep it a class so asyncio can subclass it
    def __init__(self, *args, **kwargs):
        try:
            logger.debug(
                "[SPAWN_DETECT] Popen pid=%s exe=%s args=%s kwargs_keys=%s",
                os.getpid(),
                sys.executable,
                args,
                list(kwargs.keys()),
            )
        except Exception:
            pass
        super().__init__(*args, **kwargs)

subprocess.Popen = LoggedPopen
# ================================================

# ===== DEBUG: boot trace (temporary) =====
logger.debug(
    "[BOOT] pid=%s ppid=%s sys.executable=%s argv=%s",
    os.getpid(),
    getattr(os, "getppid", lambda: -1)(),
    sys.executable,
    sys.argv,
)
# ========================================

# -----------------------------------------------------------------------------
# Boot trace (who started this server?)
# -----------------------------------------------------------------------------
def _get_process_image_path(pid: int) -> str:
    """Return full image path of a process (Windows)."""
    if os.name != "nt" or pid <= 0:
        return ""
    try:
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        h = _kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, int(pid))
        if not h:
            return ""
        try:
            buf = ctypes.create_unicode_buffer(32768)
            size = wintypes.DWORD(len(buf))
            ok = _kernel32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size))
            return buf.value if ok else ""
        finally:
            _kernel32.CloseHandle(h)
    except Exception:
        return ""

def log_boot_origin(_logger) -> None:
    """Log boot origin for correlation (pid/ppid/image/argv)."""
    try:
        pid = os.getpid()
        ppid = os.getppid()
        self_img = _get_process_image_path(pid)
        parent_img = _get_process_image_path(ppid)
        _logger.info(
            "[BOOT] pid=%s ppid=%s self_img=%s parent_img=%s sys.executable=%s argv=%s spawn_id=%s",
            pid,
            ppid,
            self_img,
            parent_img,
            sys.executable,
            sys.argv,
            os.environ.get("CSV_TOOL_SPAWN_ID", ""),
        )
    except Exception:
        # never break server on logging
        pass

# -----------------------------------------------------------------------------
# Local IPC + Mutex (no new core modules)
# -----------------------------------------------------------------------------
_MUTEX_NAME = "Global\\HC_SVC_SERVER"
_ERROR_ALREADY_EXISTS = 183
_SYNCHRONIZE = 0x00100000
_kernel32 = ctypes.windll.kernel32  # type: ignore[attr-defined]

def _create_mutex(name: str) -> tuple[int, bool]:
    h = _kernel32.CreateMutexW(None, False, wintypes.LPCWSTR(name))
    already = bool(_kernel32.GetLastError() == _ERROR_ALREADY_EXISTS)
    return int(h) if h else 0, already

def _close_handle(h: int) -> None:
    try:
        if h:
            _kernel32.CloseHandle(int(h))
    except Exception:
        pass

def _ipc_root() -> Path:
    forced = core_env.ipc_dir_raw()
    if forced:
        d = Path(forced)
    else:
        d = Path(tempfile.gettempdir()) / "csv_tool"
    d.mkdir(parents=True, exist_ok=True)
    return d

def _req_dir() -> Path:
    d = _ipc_root() / "svc_requests"
    d.mkdir(parents=True, exist_ok=True)
    return d

def _res_dir() -> Path:
    d = _ipc_root() / "svc_results"
    d.mkdir(parents=True, exist_ok=True)
    return d

def _shutdown_flag() -> Path:
    d = _ipc_root() / "control"
    d.mkdir(parents=True, exist_ok=True)
    return d / "svc_shutdown.flag"

def _is_shutdown_requested() -> bool:
    try:
        return _shutdown_flag().exists()
    except Exception:
        return False

def _clear_shutdown_flag() -> None:
    try:
        _shutdown_flag().unlink(missing_ok=True)
    except Exception:
        pass

def _atomic_write_bytes(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(
        prefix=path.name + ".", suffix=".tmp", dir=str(path.parent)
    )
    try:
        with os.fdopen(fd, "wb") as f:
            f.write(data)
            f.flush()
            os.fsync(f.fileno())
        os.replace(tmp_name, path)
    finally:
        try:
            if os.path.exists(tmp_name):
                os.remove(tmp_name)
        except OSError:
            pass

def _write_pickle(path: Path, obj: object) -> None:
    _atomic_write_bytes(path, pickle.dumps(obj, protocol=pickle.HIGHEST_PROTOCOL))

def _read_pickle(path: Path) -> object:
    """
    Windows で稀に `PermissionError: [Errno 13]` が発生することがある（AV/索引/競合）。
    要求キューは atomic replace で書かれる前提だが、読み側は短いリトライで吸収する。
    """
    last_exc: Exception | None = None
    for i in range(12):
        try:
            with path.open("rb") as f:
                return pickle.load(f)
        except PermissionError as e:
            last_exc = e
            # 0ms/5ms/10ms... と短く待つ（最大 ~330ms）
            time.sleep(min(0.03, 0.005 * (i + 1)))
        except EOFError as e:
            # 書込途中/AV 介入などの極稀ケース。短く待って再試行。
            last_exc = e
            time.sleep(min(0.03, 0.005 * (i + 1)))
    if last_exc is not None:
        raise last_exc
    with path.open("rb") as f:
        return pickle.load(f)

def _new_res_path(req_path: Path) -> Path:
    stem = req_path.stem.replace("svc_req_", "svc_res_")
    return _res_dir() / f"{stem}.pkl"

# module load log (version)
try:
    logger.info(
        "[MODULE_LOAD] %s version=%s pid=%s file=%s",
        __name__,
        __version__,
        os.getpid(),
        __file__,
    )
except Exception:
    pass

_Handler = Callable[..., Any]
_HANDLER_CACHE: dict[str, _Handler] = {}

_ACTION_MAP: dict[str, tuple[str, str]] = {
    "csv_mg": ("svc.svc_csv_mg", "merge_csv"),
    "csv_ld": ("svc.svc_csv_ld", "load_csv"),
    "csv_sv": ("svc.svc_csv_sv", "save_csv"),
    "csv_sp": ("svc.svc_csv_sp", "split_csv"),
    "dupli": ("svc.svc_dupli", "check_duplicates"),
    "row_dl": ("svc.svc_row_dl", "delete_empty_rows"),
    "col_dl": ("svc.svc_col_dl", "delete_empty_cols"),
    "dt_ymd": ("svc.svc_dt_ymd", "convert_date_ymd"),
    "dt_hm": ("svc.svc_dt_hm", "convert_date_ymd_hm"),
    "trm_ex": ("svc.svc_trm_ex", "trim_cells"),
    "hd_nr": ("svc.svc_hd_nr", "insert_header"),
    "hd_in": ("svc.svc_hd_in", "insert_header"),
    "undo": ("svc.svc_undo", "exec_undo"),
    "help": ("svc.svc_help", "show_help"),
    "update_check": ("svc.svc_packaged_update", "check_for_updates"),
    "data_agg": ("svc.svc_data_agg", "run_data_agg"),
}

# bridge / hc_main との整合検証用（テスト・ドキュメント参照可）
SVC_SERVER_ACTION_KEYS: frozenset[str] = frozenset(_ACTION_MAP.keys())

def _bootstrap_sys_path() -> None:
    """プロジェクトルートを sys.path の先頭に入れ、パッケージ import を可能にする。"""
    here = Path(__file__).resolve().parent  # .../svc
    root = here.parent
    if str(root) not in sys.path:
        sys.path.insert(0, str(root))

def _load_handler(action: str) -> _Handler:
    if action in _HANDLER_CACHE:
        return _HANDLER_CACHE[action]

    if action not in _ACTION_MAP:
        raise ValueError("Unknown action: %s" % action)

    mod_name, fn_name = _ACTION_MAP[action]
    mod = importlib.import_module(mod_name)
    # Backward-compatible handler resolution:
    if hasattr(mod, fn_name):
        fn = getattr(mod, fn_name)
    else:
        candidates: tuple[str, ...]
        if action == "csv_mg":
            candidates = ("merge_csv", "main", "run", "execute", "merge")
        elif action == "csv_ld":
            candidates = ("load_csv", "main", "run", "execute", "load")
        elif action == "csv_sv":
            candidates = ("save_csv", "main", "run", "execute", "save")
        elif action == "csv_sp":
            candidates = ("split_csv", "main", "run", "execute", "split")
        elif action == "data_agg":
            candidates = ("run_data_agg", "main", "run", "execute")
        else:
            candidates = ("main", "run", "execute")
        found_name = ""
        for cand in candidates:
            if hasattr(mod, cand) and callable(getattr(mod, cand)):
                found_name = cand
                break
        if not found_name:
            avail = [
                n
                for n in dir(mod)
                if not n.startswith("_") and callable(getattr(mod, n, None))
            ]
            raise AttributeError(
                f"module '{mod_name}' has no attribute '{fn_name}' (action={action}); "
                f"also no compatible entrypoint found. Available callables: {avail}"
            )
        logger.warning(
            "[SVC_SERVER] handler fallback: action=%s expected=%s.%s resolved=%s.%s",
            action,
            mod_name,
            fn_name,
            mod_name,
            found_name,
        )
        fn = getattr(mod, found_name)
    if not callable(fn):
        raise TypeError("Handler not callable: %s.%s" % (mod_name, fn_name))
    _HANDLER_CACHE[action] = fn
    logger.info(
        "[SVC_SERVER] handler cached: action=%s -> %s.%s", action, mod_name, fn_name
    )
    return fn

def _get_warmup_actions() -> list[str]:
    """config/svc_warmup.json の warmup_actions を返す。無い・不正なら環境変数 HC_SVC_WARMUP_ACTIONS にフォールバック。どちらも無ければ []。"""
    actions: list[str] = []
    try:
        from core.core_cst import resolve_config_file_path

        cfg_path = resolve_config_file_path("svc_warmup.json")
        if cfg_path.is_file():
            import json
            with cfg_path.open("r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, dict):
                raw_list = data.get("warmup_actions")
                if isinstance(raw_list, list):
                    actions = [str(a).strip() for a in raw_list if a]
    except Exception as e:
        logger.debug("[SVC_SERVER] warmup config read skip: %s", e)
    if not actions:
        raw = os.environ.get("HC_SVC_WARMUP_ACTIONS", "").strip()
        if raw:
            actions = [a.strip() for a in raw.split(",") if a.strip()]
    return [a for a in actions if a in _ACTION_MAP]


def _run_warmup() -> None:
    """ウォームアップリストに従い、該当 action のハンドラを事前に _load_handler でキャッシュする（遅延C 軽減）。"""
    for action in _get_warmup_actions():
        try:
            _load_handler(action)
            logger.info("[SVC_SERVER] warmup: action=%s", action)
        except Exception as e:
            logger.warning("[SVC_SERVER] warmup skip action=%s: %s", action, e)


def _find_req_files(req_dir: Path) -> list[Path]:
    try:
        return sorted(req_dir.glob("svc_req_*.pkl"), key=lambda p: p.stat().st_mtime)
    except Exception:
        return []

_book_cache_lock = threading.Lock()
_book_cache_by_hwnd: dict[int, Any] = {}
_last_attached_hwnd: int = 0


def _is_com_broken(exc: BaseException) -> bool:
    from core.excel_com_session import is_com_session_error

    return is_com_session_error(exc)


def _clear_native_exception_state() -> None:
    """pywin32 com_error 捕捉後に C レベルで残る例外状態をクリアする（Nuitka 等）。"""
    try:
        import ctypes

        ctypes.pythonapi.PyErr_Clear()
    except Exception:
        pass


def _excel_hwnd_is_live(hwnd: int) -> bool:
    from core.core_w32 import is_window

    ph = int(hwnd or 0)
    return ph > 0 and is_window(ph)


def _book_label(book: Any) -> str:
    try:
        return str(getattr(book, "name", "") or "?")
    except Exception:
        return "?"


def _validate_book_alive(book: Any) -> bool:
    try:
        _ = str(getattr(book, "name", "") or "")
        return True
    except Exception:
        _clear_native_exception_state()
        return False


def _probe_app_com(app: Any) -> bool:
    """App の COM 接続が生きているか軽量に確認する。"""
    try:
        active = app.books.active
        if active is None:
            return False
        _ = str(getattr(active, "name", "") or "")
        return True
    except Exception:
        return False


def _recover_com_apartment() -> None:
    """xlwings/COM の STA を再初期化し、stale 参照を破棄する（best-effort）。"""
    try:
        import pythoncom

        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        pythoncom.CoInitialize()
    except Exception as ex:
        logger.warning("[SVC_SERVER] attach_book com_apartment_recover skipped: %r", ex)
    _hard_reset_all_excel_com_bindings()


def _hard_reset_all_excel_com_bindings() -> None:
    """全 Excel App シェルの COM バインドと Book キャッシュを破棄する。"""
    import xlwings as xw  # type: ignore

    reset = 0
    for app in xw.apps:
        impl = getattr(app, "impl", None)
        if impl is None:
            continue
        if getattr(impl, "_xl", None) is not None or int(getattr(impl, "_hwnd", 0) or 0) > 0:
            impl._xl = None
            impl._hwnd = 0
            reset += 1
    with _book_cache_lock:
        _book_cache_by_hwnd.clear()
    if reset:
        logger.info(
            "[SVC_SERVER] attach_book hard_reset_all_com_bindings count=%s",
            reset,
        )


def _reset_app_binding_for_hwnd(excel_hwnd: int) -> None:
    """HWND に紐づく xlwings App の COM バインドを破棄し、次回再接続させる。"""
    import xlwings as xw  # type: ignore

    ph = int(excel_hwnd or 0)
    if ph <= 0:
        return
    with _book_cache_lock:
        _book_cache_by_hwnd.pop(ph, None)
    for app in xw.apps:
        impl = getattr(app, "impl", None)
        if impl is None:
            continue
        if _app_impl_hwnd(impl) != ph:
            continue
        impl._xl = None
        impl._hwnd = ph


def _purge_dead_excel_app_shells() -> None:
    """終了済み Excel の HWND シェルを掃除し、stale COM 参照を残さない。"""
    import xlwings as xw  # type: ignore

    purged = 0
    for app in xw.apps:
        impl = getattr(app, "impl", None)
        if impl is None:
            continue
        hwnd = _app_impl_hwnd(impl)
        if hwnd <= 0 or _excel_hwnd_is_live(hwnd):
            continue
        impl._xl = None
        impl._hwnd = 0
        purged += 1
        with _book_cache_lock:
            _book_cache_by_hwnd.pop(hwnd, None)
    if purged:
        logger.info(
            "[SVC_SERVER] attach_book purged_dead_app_shells count=%s",
            purged,
        )


def _app_impl_hwnd(impl: Any) -> int:
    """WinApp impl の HWND（lazy 時は _hwnd のみ参照し COM を起こさない）。"""
    try:
        hwnd = int(getattr(impl, "_hwnd", 0) or 0)
        if hwnd > 0:
            return hwnd
    except Exception:
        pass
    try:
        xl = getattr(impl, "_xl", None)
        if xl is not None:
            return int(xl.Hwnd)  # type: ignore[attr-defined]
    except Exception:
        pass
    return 0


def _release_other_excel_com_bindings(keep_hwnd: int) -> None:
    """他 Excel インスタンスの COM バインドを外し、対象 HWND の接続だけを有効にする。

    常駐 svc_server で 2 つ目の Excel へ接続するとき、先に解決済みの _xl が残っていると
    get_xl_app_from_hwnd がハング／クラッシュすることがあるため。
    """
    import xlwings as xw  # type: ignore

    keep = int(keep_hwnd or 0)
    released = 0
    for app in xw.apps:
        impl = getattr(app, "impl", None)
        if impl is None:
            continue
        hwnd = _app_impl_hwnd(impl)
        if hwnd == keep:
            continue
        if hwnd > 0 and not _excel_hwnd_is_live(hwnd):
            impl._xl = None
            impl._hwnd = 0
            released += 1
            with _book_cache_lock:
                _book_cache_by_hwnd.pop(hwnd, None)
            continue
        if getattr(impl, "_xl", None) is None:
            continue
        try:
            if hwnd <= 0:
                hwnd = int(impl._xl.Hwnd)  # type: ignore[attr-defined]
        except Exception:
            hwnd = 0
        impl._xl = None
        if hwnd > 0:
            impl._hwnd = hwnd
        released += 1
    if released:
        logger.info(
            "[SVC_SERVER] attach_book released_other_com_bindings keep_hwnd=%s count=%s",
            keep,
            released,
        )
    with _book_cache_lock:
        for cached_hwnd in list(_book_cache_by_hwnd):
            if int(cached_hwnd) != keep:
                _book_cache_by_hwnd.pop(cached_hwnd, None)


def _find_or_create_app_for_hwnd(excel_hwnd: int, *, force_fresh: bool = False):
    """HWND に対応する xlwings App（xw.apps 走査優先、未登録時は lazy shell を追加）。"""
    import xlwings as xw  # type: ignore
    from xlwings._xlwindows import App as WinApp

    ph = int(excel_hwnd or 0)
    if ph <= 0:
        raise RuntimeError("excel_hwnd missing")
    if not _excel_hwnd_is_live(ph):
        raise RuntimeError(f"Excel window not found (hwnd={ph})")

    if not force_fresh:
        for app in xw.apps:
            impl = getattr(app, "impl", None)
            if impl is None or _app_impl_hwnd(impl) != ph:
                continue
            if _probe_app_com(app):
                logger.info(
                    "[SVC_SERVER] attach_book app_hit phase=apps_scan hwnd=%s",
                    ph,
                )
                return app
            logger.info(
                "[SVC_SERVER] attach_book app_stale phase=apps_scan hwnd=%s",
                ph,
            )
            _reset_app_binding_for_hwnd(ph)
            break

    logger.info(
        "[SVC_SERVER] attach_book app_create phase=%s hwnd=%s",
        "force_fresh" if force_fresh else "apps_miss_or_stale",
        ph,
    )
    return xw.App(impl=WinApp(xl=ph))


def _pick_book_from_app(target_app, *, book_fullname: str, book_name: str) -> Any:
    if book_fullname:
        for b in target_app.books:
            try:
                if str(getattr(b, "fullname", "")) == book_fullname:
                    return b
            except Exception:
                continue
    if book_name:
        for b in target_app.books:
            try:
                if str(getattr(b, "name", "")) == book_name:
                    return b
            except Exception:
                continue
    active = target_app.books.active
    if active is not None:
        return active
    raise RuntimeError(
        f"Workbook not found (fullname={book_fullname!r} name={book_name!r})"
    )


def _attach_book_via_xlc_context(
    excel_hwnd: int,
    *,
    book_fullname: str,
    book_name: str,
) -> Any | None:
    """short_runner と同じ get_excel_context_from_hwnd 経路で Book を取得する。"""
    from core.core_xlc import get_excel_context_from_hwnd

    ph = int(excel_hwnd or 0)
    if ph <= 0:
        return None
    ctx = get_excel_context_from_hwnd(ph, "")
    if ctx is None:
        return None
    _app, book, _sheet, _hwnd = ctx
    if book is None:
        return None
    if book_name:
        try:
            if str(getattr(book, "name", "")) != book_name:
                logger.info(
                    "[SVC_SERVER] attach_book xlc_ctx name_mismatch hwnd=%s expected=%r actual=%r",
                    ph,
                    book_name,
                    getattr(book, "name", "?"),
                )
        except Exception:
            pass
    if book_fullname:
        try:
            actual_full = str(getattr(book, "fullname", "") or "")
            if actual_full and actual_full != book_fullname:
                logger.info(
                    "[SVC_SERVER] attach_book xlc_ctx fullname_mismatch hwnd=%s expected=%r actual=%r",
                    ph,
                    book_fullname,
                    actual_full,
                )
        except Exception:
            pass
    if not _validate_book_alive(book):
        logger.info(
            "[SVC_SERVER] attach_book phase=xlc_ctx_stale hwnd=%s book=%s",
            ph,
            _book_label(book),
        )
        return None
    logger.info(
        "[SVC_SERVER] attach_book phase=xlc_ctx_ok hwnd=%s book=%s",
        ph,
        _book_label(book),
    )
    return book


def invalidate_attached_book_cache(excel_hwnd: int = 0) -> None:
    """Book キャッシュを破棄する（操作境界の再取得用）。excel_hwnd=0 は全 HWND。"""
    ph = int(excel_hwnd or 0)
    with _book_cache_lock:
        if ph <= 0:
            _book_cache_by_hwnd.clear()
        else:
            _book_cache_by_hwnd.pop(ph, None)
    if ph > 0:
        _reset_app_binding_for_hwnd(ph)


def resolve_fresh_book_after_ui_wait(
    book: Any,
    *,
    excel_hwnd: int = 0,
    book_fullname: str = "",
    book_name: str = "",
    parent_hwnd: int = 0,
    log_prefix: str = "SVC_SERVER",
) -> Any:
    """UI 待ち後に HWND から Book を再取得する（古い xlwings 参照で COM を触らない）。

    find_sheet_by_guid 等の前に呼び、COM_NG の根本回避とする。失敗時は book をそのまま返す。
    """
    hwnd = int(excel_hwnd or parent_hwnd or 0)
    bfn = str(book_fullname or "")
    bname = str(book_name or "")
    if hwnd <= 0 and not bfn and not bname and book is not None:
        try:
            hwnd = int(getattr(book.app, "hwnd", 0) or 0)
        except Exception:
            hwnd = 0
        try:
            bfn = str(getattr(book, "fullname", "") or "")
        except Exception:
            bfn = ""
        try:
            bname = str(getattr(book, "name", "") or "")
        except Exception:
            bname = ""
    if hwnd <= 0 and not bfn and not bname:
        return book
    try:
        fresh = _attach_book(
            excel_hwnd=hwnd,
            book_fullname=bfn,
            book_name=bname,
            skip_cache=True,
        )
        logger.info(
            "[%s] book_fresh_after_ui_wait hwnd=%s book=%s",
            log_prefix,
            hwnd,
            _book_label(fresh),
        )
        return fresh
    except Exception as ex:
        logger.warning(
            "[%s] book_fresh_after_ui_wait failed hwnd=%s ex=%r",
            log_prefix,
            hwnd,
            ex,
        )
        return book


def _store_attached_book(excel_hwnd: int, book: Any) -> Any:
    global _last_attached_hwnd

    ph = int(excel_hwnd or 0)
    with _book_cache_lock:
        _book_cache_by_hwnd[ph] = book
    _last_attached_hwnd = ph
    try:
        from core.excel_com_session import record_com_session_hwnd

        record_com_session_hwnd(ph)
    except Exception:
        pass
    return book


def _cached_book_if_alive(excel_hwnd: int) -> Any | None:
    ph = int(excel_hwnd or 0)
    if ph <= 0 or not _excel_hwnd_is_live(ph):
        with _book_cache_lock:
            _book_cache_by_hwnd.pop(ph, None)
        return None
    with _book_cache_lock:
        book = _book_cache_by_hwnd.get(ph)
    if book is None:
        return None
    if _validate_book_alive(book):
        return book
    with _book_cache_lock:
        _book_cache_by_hwnd.pop(ph, None)
    return None


def _resolve_attached_book(
    excel_hwnd: int,
    book_fullname: str,
    book_name: str,
) -> Any:
    """HWND から Book を解決しキャッシュへ格納する（cache_hit なし）。"""
    ph = int(excel_hwnd or 0)
    logger.info(
        "[SVC_SERVER] attach_book phase=resolve_enter hwnd=%s fullname=%r name=%r",
        ph,
        book_fullname or "",
        book_name or "",
    )

    book = _attach_book_via_xlc_context(
        ph, book_fullname=book_fullname, book_name=book_name
    )
    if book is not None:
        logger.info(
            "[SVC_SERVER] attach_book phase=resolve_ok hwnd=%s via=xlc_ctx book=%s",
            ph,
            _book_label(book),
        )
        return _store_attached_book(ph, book)

    logger.info("[SVC_SERVER] attach_book phase=xlc_ctx_miss hwnd=%s", ph)
    _release_other_excel_com_bindings(ph)
    target_app = _find_or_create_app_for_hwnd(ph, force_fresh=True)
    book = _pick_book_from_app(
        target_app, book_fullname=book_fullname, book_name=book_name
    )
    if not _validate_book_alive(book):
        raise RuntimeError(f"Workbook COM stale (hwnd={ph})")
    logger.info(
        "[SVC_SERVER] attach_book phase=resolve_ok hwnd=%s via=xlwings book=%s",
        ph,
        _book_label(book),
    )
    return _store_attached_book(ph, book)


def _attach_book(
    excel_hwnd: int,
    book_fullname: str,
    book_name: str,
    *,
    skip_cache: bool = False,
):
    """excel_hwnd と book識別子から xlwings.Book を取得する。

    Notes:
        - マルチ Excel は _book_cache_by_hwnd で HWND ごとに Book を保持する（B+ 常駐）。
        - skip_cache=True のときは cache_hit せず HWND から再取得する（UI 待ち action 用）。
        - 接続経路は short_runner と同じ get_excel_context_from_hwnd を優先する。
    """
    ph = int(excel_hwnd or 0)
    if ph <= 0:
        raise RuntimeError("excel_hwnd missing")
    if not _excel_hwnd_is_live(ph):
        raise RuntimeError(f"Excel window not found (hwnd={ph})")

    _purge_dead_excel_app_shells()

    if skip_cache:
        invalidate_attached_book_cache(ph)
    else:
        cached = _cached_book_if_alive(ph)
        if cached is not None:
            logger.info("[SVC_SERVER] attach_book phase=cache_hit hwnd=%s", ph)
            return cached

    try:
        return _resolve_attached_book(ph, book_fullname, book_name)
    except Exception as ex:
        if not _is_com_broken(ex):
            raise
        logger.warning(
            "[SVC_SERVER] attach_book com_retry hwnd=%s ex=%r",
            ph,
            ex,
        )
        _clear_native_exception_state()
        invalidate_attached_book_cache(ph)
        _recover_com_apartment()
        return _resolve_attached_book(ph, book_fullname, book_name)


_COM_MONITOR_POLL_SEC = 1.0
_com_monitor_started = False
_com_monitor_lock = threading.Lock()


def _schedule_com_recycle_after_op(*, reason: str = "com_session_error") -> None:
    """COM 汚染時にプロセスを終了し、次回リクエストでクリーンな COM を得る。"""
    try:
        _shutdown_flag().write_text(str(reason or "com_session_error"), encoding="utf-8")
        logger.info("[SVC_SERVER] com_recycle scheduled reason=%s", reason)
    except Exception:
        pass


def _sync_ipc_last_com_hwnd() -> None:
    """診断用 IPC（svc_last_com_hwnd.txt）をプロセス内の last_attached と同期する。"""
    try:
        from svc.svc_host import write_last_svc_com_hwnd

        last = int(_last_attached_hwnd or 0)
        if last > 0 and _excel_hwnd_is_live(last):
            write_last_svc_com_hwnd(last)
        else:
            write_last_svc_com_hwnd(0)
    except Exception:
        pass


def _prune_stale_hwnd_cache() -> int:
    """終了済み Excel の HWND を Book キャッシュから除去する（プロセスは維持）。"""
    global _last_attached_hwnd

    with _book_cache_lock:
        stale_hwnds = [
            int(h)
            for h in _book_cache_by_hwnd
            if not _excel_hwnd_is_live(int(h))
        ]
    pruned = 0
    for ph in stale_hwnds:
        with _book_cache_lock:
            if _book_cache_by_hwnd.pop(ph, None) is not None:
                pruned += 1
        _reset_app_binding_for_hwnd(ph)
    last = int(_last_attached_hwnd or 0)
    if last > 0 and not _excel_hwnd_is_live(last):
        _last_attached_hwnd = 0
        logger.info("[SVC_SERVER] com_monitor cleared stale last_attached_hwnd=%s", last)
    _purge_dead_excel_app_shells()
    _sync_ipc_last_com_hwnd()
    return pruned


def _com_hwnd_monitor_worker() -> None:
    """キャッシュ済み HWND のうち終了済み Excel を定期的に掃除する（常駐は維持）。"""
    while True:
        try:
            time.sleep(_COM_MONITOR_POLL_SEC)
            if _is_shutdown_requested():
                return
            pruned = _prune_stale_hwnd_cache()
            if pruned:
                logger.info(
                    "[SVC_SERVER] com_monitor pruned_dead_hwnds count=%s",
                    pruned,
                )
        except Exception:
            return


def _ensure_com_hwnd_monitor() -> None:
    global _com_monitor_started
    with _com_monitor_lock:
        if _com_monitor_started:
            return
        th = threading.Thread(
            target=_com_hwnd_monitor_worker,
            name="SvcComHwndMonitor",
            daemon=True,
        )
        th.start()
        _com_monitor_started = True
        logger.info(
            "[SVC_SERVER] com_monitor started poll_sec=%s",
            _COM_MONITOR_POLL_SEC,
        )

def _call_handler(
    handler, action: str, *, excel_hwnd: int, sheet_id: str, book, kwargs: dict
) -> None:
    """Call handler with best-effort argument compatibility.

    We do NOT assume a single signature across legacy svc modules.
    Strategy:
      - Prefer keyword arguments if accepted.
      - Otherwise, pass positional (book, sheet_id) when possible.
      - Filter extra kwargs unless handler accepts **kwargs.
    """
    import inspect

    sig = inspect.signature(handler)
    params = list(sig.parameters.values())
    accepts_var_kw = any(p.kind == inspect.Parameter.VAR_KEYWORD for p in params)

    # Candidate values (some modules use book_ptr/workbook/etc.)
    values = {
        "excel_hwnd": excel_hwnd,
        "target_hwnd": excel_hwnd,
        "hwnd": excel_hwnd,
        "sheet_id": sheet_id,
        "sheet_guid": sheet_id,
        "book": book,
        "wb": book,
        "workbook": book,
        "book_ptr": book,
        "book_pointer": book,
        "workbook_pointer": book,
    }

    # 1) Build keyword args that match parameters
    call_kwargs: dict = {}
    for name, val in kwargs.items():
        if accepts_var_kw or name in sig.parameters:
            call_kwargs[name] = val

    for k, v in values.items():
        if k in sig.parameters:
            call_kwargs[k] = v

    # 2) If we can satisfy all required params with keywords, do it
    missing_required = []
    for p in params:
        if p.kind in (inspect.Parameter.VAR_POSITIONAL, inspect.Parameter.VAR_KEYWORD):
            continue
        if p.default is not inspect._empty:
            continue
        if p.name not in call_kwargs:
            missing_required.append(p.name)
    if not missing_required:
        handler(**call_kwargs)
        return

    # 3) Try positional (book, sheet_id) + remaining keywords
    pos: list = []
    pos_params = [
        p
        for p in params
        if p.kind
        in (inspect.Parameter.POSITIONAL_ONLY, inspect.Parameter.POSITIONAL_OR_KEYWORD)
    ]
    if len(pos_params) >= 1:
        pos.append(book)
    if len(pos_params) >= 2:
        pos.append(sheet_id)

    # remove any conflicting names that would double-pass
    for p in pos_params[: len(pos)]:
        call_kwargs.pop(p.name, None)

    handler(*pos, **call_kwargs)

def _process_one(req_path: Path) -> None:
    from core.excel_com_session import (
        action_attach_book_fresh_resolve,
        action_uses_attach_book,
        should_schedule_com_recycle_after_handler,
    )

    res_path = _new_res_path(req_path)
    action = ""
    handler_ok = False
    excel_hwnd = 0
    handler_exc: Exception | None = None
    try:
        req = _read_pickle(req_path)
        if not isinstance(req, dict):
            raise TypeError("request pickle must be a dict")
        action = str(req.get("action", "")).strip()
        kwargs = req.get("kwargs", {})
        if not isinstance(kwargs, dict):
            kwargs = {}

        handler = _load_handler(action)

        # 共通パラメータ（hc_main から渡される）
        excel_hwnd = int(kwargs.pop("excel_hwnd", 0) or 0)
        book_fullname = str(kwargs.pop("book_fullname", "") or "")
        book_name = str(kwargs.pop("book_name", "") or "")
        sheet_id = str(kwargs.pop("sheet_id", "") or "")

        logger.info(
            "[SVC_SERVER] exec start action=%s req=%s pid=%s",
            action,
            req_path.name,
            os.getpid(),
        )
        t0 = time.perf_counter()

        # action ごとの呼び分け（既存svcモジュールのシグネチャを尊重）
        if action_uses_attach_book(action):
            fresh = action_attach_book_fresh_resolve(action)
            if fresh and excel_hwnd > 0:
                invalidate_attached_book_cache(excel_hwnd)
            book = _attach_book(
                excel_hwnd=excel_hwnd,
                book_fullname=book_fullname,
                book_name=book_name,
                skip_cache=fresh,
            )
            _call_handler(
                handler,
                action,
                excel_hwnd=excel_hwnd,
                sheet_id=sheet_id,
                book=book,
                kwargs=kwargs,
            )
        else:
            # 多くの svc は target_hwnd を受け取る
            _call_handler(
                handler,
                action,
                excel_hwnd=excel_hwnd,
                sheet_id=sheet_id,
                book=None,
                kwargs=kwargs,
            )

        ms = int((time.perf_counter() - t0) * 1000)
        _write_pickle(res_path, {"status": "OK", "ms": ms})
        handler_ok = True
        logger.info(
            "[SVC_SERVER] exec done action=%s ms=%s res=%s", action, ms, res_path.name
        )
    except Exception as ex:
        handler_exc = ex
        import traceback

        tb = traceback.format_exc()
        try:
            _write_pickle(
                res_path, {"status": "ERROR", "error": str(ex), "traceback": tb}
            )
        except Exception:
            pass
        logger.exception("[SVC_SERVER] exec failed req=%s ex=%s", req_path.name, ex)
    finally:
        try:
            from core.ribbon_public_to_svc import SVC_ACTIONS_NOTIFY_WAITFORM_AFTER_HANDLER

            if action in SVC_ACTIONS_NOTIFY_WAITFORM_AFTER_HANDLER:
                from core.core_cursor import notify_wait_form_ready

                notify_wait_form_ready(parent_hwnd=excel_hwnd)
            elif action == "data_agg" and not handler_ok:
                from core.core_cursor import notify_wait_form_ready

                notify_wait_form_ready(parent_hwnd=excel_hwnd)
        except Exception:
            pass
        if should_schedule_com_recycle_after_handler(
            action,
            handler_ok=handler_ok,
            exc=handler_exc,
        ):
            _schedule_com_recycle_after_op(reason="com_session_error")
        elif action_uses_attach_book(action) and excel_hwnd > 0:
            invalidate_attached_book_cache(excel_hwnd)
        try:
            req_path.unlink(missing_ok=True)
        except Exception:
            pass

def main() -> int:
    """svc 常駐サーバのエントリ: mutex・IPC ループでリクエストを処理し終了する。"""
    log_boot_origin(logger)
    _bootstrap_sys_path()

    mh_handle, mh_exists = _create_mutex(_MUTEX_NAME)
    if mh_exists:
        logger.info("[SVC_SERVER] already running (mutex exists). pid=%s", os.getpid())
        return 0

    try:
        from core.ipc_cleanup import run_svc_server_startup_sweeps

        run_svc_server_startup_sweeps(_ipc_root())
    except Exception:
        pass

    _clear_shutdown_flag()
    logger.info(
        "[SVC_SERVER] started pid=%s file=%s", os.getpid(), Path(__file__).resolve()
    )
    try:
        from core.excel_lifecycle_monitor import ensure_excel_lifecycle_monitor

        ensure_excel_lifecycle_monitor()
    except Exception:
        pass
    try:
        _ensure_com_hwnd_monitor()
    except Exception:
        pass
    _run_warmup()

    req_dir = _req_dir()
    idle_sleep = float(os.environ.get("HC_SVC_IDLE_POLL_SEC", "0.1"))

    try:
        while True:
            if _is_shutdown_requested():
                logger.info("[SVC_SERVER] shutdown requested (flag).")
                break

            reqs = _find_req_files(req_dir)
            if not reqs:
                time.sleep(idle_sleep)
                continue

            for p in reqs:
                _process_one(p)
    finally:
        try:
            _close_handle(mh_handle)
        except Exception:
            pass
        _close_handle(mh_handle)
        logger.info("[SVC_SERVER] exiting pid=%s", os.getpid())
    return 0

if __name__ == "__main__":
    raise SystemExit(main())

# Release: svc_server_clean_0.1.6
