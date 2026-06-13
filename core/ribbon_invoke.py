# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: core.ribbon_invoke
Created: 2025-11-28 (logic from ルート hc_main / hc_invoke)
Updated: 2026-06-07
Version: 1.12.2
Purpose:
    xlwings 短寿命プロセスからの invoke / register_book / clear_registry（司令塔）。
    フェーズ C で `core/` に移設。ルート `hc_main.py` は常駐ブリッジ専用。
    VBA からの公開入口は **core.excel_session** 経由。リボン tag と action は同一文字列。

History (latest 3):
  - 1.12.2 (2026-06-13): _ensure_book を HWND 直結（apps.active 廃止）。マルチ Excel 対応。
  - 1.12.1 (2026-06-07): xlwings 先行 import（起動時スレッド prewarm）を追加。
  - 1.12.0 (2026-04-11) `hc_invoke.py` から `core/ribbon_invoke.py` へ移設。診断ロガー `hc_csv_tool.diag.ribbon_invoke`。
  - 1.11.9 (2026-04-11) ルート hc_invoke.py。_invoke_simple_svc / _invoke_csv_family 集約。
  - 1.11.8 (2026-04-11) invoke finally の WaitForm 通知対象を ribbon_public_to_svc に集約。
"""

from __future__ import annotations

import os
import sys
import time
from pathlib import Path
import tempfile
import pickle
import threading
from typing import Any, Callable, Dict, Optional

# Bootstrap: project root on sys.path（本モジュールは core/ 配下）
_ribbon_invoke_path = Path(__file__).resolve()
path_commander_raw_v = str(_ribbon_invoke_path)
path_project_root_ptr = str(_ribbon_invoke_path.parent.parent)
if path_project_root_ptr not in sys.path:
    sys.path.insert(0, path_project_root_ptr)

from core import core_env
from core.core_log import get_diag_logger, get_logger, get_perf_logger
from core.ribbon_public_to_svc import (
    RIBBON_INVOKE_ACTION_KEYS,
    RIBBON_INVOKE_FINALLY_NOTIFY_WAITFORM,
    RIBBON_PUBLIC_TO_SVC_ACTION,
    RIBBON_TARGET_SVC_ACTION_KEYS,
)

__version__ = "1.12.2"

logger = get_logger(__name__)

# ==============================================================================
# svc_server action constants（値は core.ribbon_public_to_svc と svc_server._ACTION_MAP と一致）
# ==============================================================================
ACTION_CSV_MG = RIBBON_PUBLIC_TO_SVC_ACTION["merge_csv"]
ACTION_CSV_LD = RIBBON_PUBLIC_TO_SVC_ACTION["load_csv"]
ACTION_CSV_SV = RIBBON_PUBLIC_TO_SVC_ACTION["save_csv"]
ACTION_CSV_SP = RIBBON_PUBLIC_TO_SVC_ACTION["split_csv"]

_ALLOWED_SVC_ACTIONS: frozenset[str] = RIBBON_TARGET_SVC_ACTION_KEYS

# ==============================================================================
# svc_server wait policy
#   - Default timeout is controlled by env: HC_SVC_TIMEOUT_SEC (seconds)
#   - Per-action override: HC_SVC_TIMEOUT_<ACTION>_SEC (e.g. HC_SVC_TIMEOUT_CSV_MG_SEC)
#   - If value is "none"/"inf"/"infinite"/"0": treat as no-timeout (wait forever)
# ==============================================================================
_SVC_TIMEOUT_DEFAULT_SEC: float = float(os.environ.get("HC_SVC_TIMEOUT_SEC", "180"))

# ==============================================================================
# 早期復帰 (Excel 無応答対策)
#   - HC_RETURN_EARLY=1 (既定): 依頼を書き出したあと、core_env.return_early_wait_sec() 秒待ってから結果を待たずに return する。
#     これにより Python プロセスが早く終了し、VBA に実行権が戻るため Excel が OS から無応答とみなされにくい。
#   - HC_RETURN_EARLY=0: 従来どおり結果ファイル (res_*.pkl) ができるまで待つ。
# ==============================================================================
def _return_early_enabled() -> bool:
    try:
        return os.environ.get("HC_RETURN_EARLY", "1").strip() in ("1", "true", "yes", "on")
    except Exception:
        return True

# None = no timeout (wait forever)
_SVC_TIMEOUT_SEC_BY_ACTION: dict[str, float | None] = {
    ACTION_CSV_MG: None,  # 対話UI：ユーザーが放置してもタイムアウトしない
    ACTION_CSV_LD: None,  # ファイル選択UI：選択に時間がかかってもタイムアウトしない
}


def _svc_timeout_sec_for(action: str) -> float | None:
    """Return timeout seconds for the given svc action.

    Returns:
        float: timeout seconds
        None:  no-timeout (wait forever)
    """
    env_key = f"HC_SVC_TIMEOUT_{action.upper()}_SEC"
    raw = os.environ.get(env_key)
    if raw is not None:
        val = raw.strip().lower()
        if val in {"none", "inf", "infinite", "forever", "0"}:
            return None
        try:
            return float(val)
        except ValueError:
            # fall back to configured mapping/default if env is invalid
            pass

    if action in _SVC_TIMEOUT_SEC_BY_ACTION:
        return _SVC_TIMEOUT_SEC_BY_ACTION[action]

    return _SVC_TIMEOUT_DEFAULT_SEC

try:
    get_diag_logger("hc_csv_tool.diag.ribbon_invoke").info(
        "[MODULE_LOAD] %s version=%s pid=%s file=%s",
        __name__,
        __version__,
        os.getpid(),
        __file__,
    )
except Exception:
    pass

HC_DEBUG: bool = False


def _is_debug_enabled() -> bool:
    try:
        return HC_DEBUG or core_env.log_diag_enabled()
    except Exception:
        return HC_DEBUG


def _log_debug(msg: str) -> None:
    try:
        if _is_debug_enabled():
            get_diag_logger("hc_csv_tool.diag.ribbon_invoke").info("[DEBUG] %s", msg)
    except Exception:
        pass



# ==============================================================================
# svc_server IPC（別プロセス常駐svc）呼び出し
# ==============================================================================

# ==============================================================================
# svc_server IPC（別プロセス常駐svc）: hc_main -> svc/svc_server.py
# ==============================================================================
def _svc_get_ipc_root() -> Path:
    forced = core_env.ipc_dir_raw()
    if forced:
        d = Path(forced)
    else:
        d = Path(tempfile.gettempdir()) / "csv_tool"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _svc_req_dir() -> Path:
    d = _svc_get_ipc_root() / "svc_requests"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _svc_res_dir() -> Path:
    d = _svc_get_ipc_root() / "svc_results"
    d.mkdir(parents=True, exist_ok=True)
    return d


def _svc_atomic_write_bytes(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp_name = tempfile.mkstemp(prefix=path.name + ".", suffix=".tmp", dir=str(path.parent))
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


def _svc_write_pickle(path: Path, obj: object) -> None:
    _svc_atomic_write_bytes(path, pickle.dumps(obj, protocol=pickle.HIGHEST_PROTOCOL))


def _svc_read_pickle(path: Path) -> object:
    with path.open("rb") as f:
        return pickle.load(f)


def _svc_new_req_path() -> Path:
    ts_ms = int(time.time() * 1000)
    return _svc_req_dir() / f"svc_req_{ts_ms}_{os.getpid()}.pkl"


def _svc_res_path_for(req_path: Path) -> Path:
    stem = req_path.stem.replace("svc_req_", "svc_res_")
    return _svc_res_dir() / f"{stem}.pkl"


def _cleanup_old_res_files(max_age_sec: int = 3600) -> None:
    """HC_RETURN_EARLY 時に結果を読まないため残る res_*.pkl を、指定秒数より古いものだけ削除する。"""
    try:
        res_dir = _svc_res_dir()
        now = time.time()
        for p in res_dir.glob("svc_res_*.pkl"):
            try:
                if (now - p.stat().st_mtime) > max_age_sec:
                    p.unlink(missing_ok=True)
            except OSError:
                pass
    except Exception:
        pass


def _call_svc_server(action: str, book_ptr: Any, sheet_id: str = "", **kwargs: object) -> None:
    """svc_server に依頼し、結果を待つ（最小IPC）。

    HC_RETURN_EARLY=1（既定）のときは依頼を書き出したあと return_early_wait_sec() 秒待って return し、
    結果は待たない。これにより Python プロセスが早く終了し、VBA に実行権が戻るため
    Excel が OS から無応答とみなされにくい。実処理・完了通知は svc_server / UI 側で行う。

    Notes:
        - book はプロセス間で渡せないため、book_name/fullname と excel_hwnd を渡す。
        - 処理本体は svc_server 側で xlwings からブックを引き直して実行する。
    """
    from svc.svc_host import ensure_svc_server  # 遅延 import（RunPython 短寿命プロセスの起動コスト低減）

    plog = get_perf_logger(f"{__name__}.ribbon_invoke")
    t_cs = time.perf_counter()
    plog.info("call_svc phase=enter action=%s cumulative_ms=0", action)
    # Test Log Out
    logger.info(
        "[HC_MAIN->HOST] enter ensure_svc_server action=%s sheet_id=%s pid=%s exe=%s",
        action, (sheet_id or ""), os.getpid(), sys.executable
    )

    ensure_svc_server()
    plog.info(
        "call_svc phase=after_ensure_svc_server action=%s cumulative_ms=%d",
        action,
        int((time.perf_counter() - t_cs) * 1000),
    )

    # Test Log Out
    logger.info(
        "[HC_MAIN->HOST] leave ensure_svc_server action=%s sheet_id=%s pid=%s",
        action, (sheet_id or ""), os.getpid()
    )

    if action not in _ALLOWED_SVC_ACTIONS:
        raise ValueError(f"Unknown svc action: {action}")

    req_path = _svc_new_req_path()
    res_path = _svc_res_path_for(req_path)

    try:
        book_fullname = str(book_ptr.fullname)
    except Exception:
        book_fullname = ""
    try:
        book_name = str(book_ptr.name)
    except Exception:
        book_name = ""

    excel_hwnd = 0
    try:
        excel_hwnd = int(getattr(book_ptr.app, "hwnd", 0) or 0)
    except Exception:
        excel_hwnd = 0

    req = {
        "action": action,
        "args": [],
        "kwargs": {
            "excel_hwnd": excel_hwnd,
            "book_fullname": book_fullname,
            "book_name": book_name,
            "sheet_id": sheet_id or "",
            **kwargs,
        },
    }
    _svc_write_pickle(req_path, req)
    plog.info(
        "call_svc phase=after_write_req action=%s cumulative_ms=%d",
        action,
        int((time.perf_counter() - t_cs) * 1000),
    )

    if _return_early_enabled():
        plog.info(
            "call_svc phase=before_return_early_sleep action=%s cumulative_ms=%d",
            action,
            int((time.perf_counter() - t_cs) * 1000),
        )
        # 依頼だけ渡して return し、Python プロセスを終了させる。VBA に実行権が戻り Excel 無応答を防ぐ。
        # 待機秒は core_env（HC_RETURN_EARLY_WAIT_SEC、既定 1.0＝従来固定値）。短縮例: 0.5。コード既定を 0.5 に変える方針は取らない。
        _wait_env_key = "HC_RETURN_EARLY_WAIT_SEC"
        _wait_env_defined = _wait_env_key in os.environ
        _wait_env_raw = (os.environ.get(_wait_env_key) or "").strip()
        wait_sec = core_env.return_early_wait_sec()
        plog.info(
            "call_svc phase=return_early_wait_ready action=%s wait_sec_effective=%s "
            "env_defined=%s env_value=%r cumulative_ms=%d",
            action,
            wait_sec,
            _wait_env_defined,
            _wait_env_raw,
            int((time.perf_counter() - t_cs) * 1000),
        )
        t_sleep0 = time.perf_counter()
        if wait_sec > 0:
            time.sleep(wait_sec)
        sleep_actual_ms = int((time.perf_counter() - t_sleep0) * 1000)
        plog.info(
            "call_svc phase=after_return_early_sleep action=%s wait_sec_effective=%s actual_sleep_ms=%d cumulative_ms=%d",
            action,
            wait_sec,
            sleep_actual_ms,
            int((time.perf_counter() - t_cs) * 1000),
        )
        _cleanup_old_res_files(max_age_sec=3600)
        logger.info(
            "[HC_MAIN->HOST] return early (HC_RETURN_EARLY); result will be handled by svc_server/UI."
        )
        plog.info(
            "call_svc phase=after_return_early action=%s cumulative_ms=%d",
            action,
            int((time.perf_counter() - t_cs) * 1000),
        )
        return

    timeout_sec = _svc_timeout_sec_for(action)
    t0 = time.time()
    while True:
        if res_path.exists():
            break
        if timeout_sec is not None and (time.time() - t0) > timeout_sec:
            raise TimeoutError("svc_server timeout: action=%s req=%s" % (action, req_path.name))
        time.sleep(0.05)

    res = _svc_read_pickle(res_path)
    try:
        res_path.unlink(missing_ok=True)
    except Exception:
        pass

    if not isinstance(res, dict):
        raise RuntimeError("svc_server invalid response: %r" % (res,))
    if res.get("status") == "OK":
        plog.info(
            "call_svc phase=after_sync_ok action=%s cumulative_ms=%d",
            action,
            int((time.perf_counter() - t_cs) * 1000),
        )
        try:
            from core.core_cursor import notify_wait_form_ready

            notify_wait_form_ready(parent_hwnd=excel_hwnd)
        except Exception:
            pass
        return

    raise RuntimeError("svc_server ERROR: %s\n%s" % (res.get("error", ""), res.get("traceback", "")))

# =============================================================================
# 内部参照レジストリ管理
# =============================================================================
# 変数: アクティブなブックインスタンスを HWND をキーに物理保持する辞書。
_active_books: Dict[int, Any] = {}
# 変数: レジストリ操作時のスレッド排他制御用ロックオブジェクト。
_registry_lock: threading.Lock = threading.Lock()


# =============================================================================
# xlwings 先行 import（起動時: 子プロセス spawn 待ちと並行）
# =============================================================================
_xlwings_prewarm_lock = threading.Lock()
_xlwings_prewarm_started = False
_xlwings_prewarm_done = threading.Event()


def start_xlwings_import_prewarm() -> None:
    """import xlwings のみをバックグラウンドで実行（COM 操作は行わない）。"""
    global _xlwings_prewarm_started
    with _xlwings_prewarm_lock:
        if _xlwings_prewarm_started:
            return
        _xlwings_prewarm_started = True

        def _worker() -> None:
            try:
                import xlwings as xw  # noqa: F401
            except Exception as ex:
                logger.debug("[XLWINGS_PREWARM] import failed: %r", ex)
            finally:
                _xlwings_prewarm_done.set()

        threading.Thread(
            target=_worker,
            name="xlwings-import-prewarm",
            daemon=True,
        ).start()


def wait_xlwings_import_prewarm(timeout_sec: float = 120.0) -> None:
    """prewarm 完了を待つ。未開始なら同期的 import にフォールバック。"""
    if not _xlwings_prewarm_started:
        import xlwings as xw  # noqa: F401
        _xlwings_prewarm_done.set()
        return
    if _xlwings_prewarm_done.wait(timeout=max(0.0, float(timeout_sec))):
        return
    logger.warning("[XLWINGS_PREWARM] wait timeout; sync import fallback")
    import xlwings as xw  # noqa: F401
    _xlwings_prewarm_done.set()


# =============================================================================
# 内部補助：自己修復登録ロジック
# =============================================================================
def _ensure_book(target_hwnd: Optional[int]) -> Optional[Any]:
    """
    Method Name : _ensure_book
    Arguments   : target_hwnd (Optional[int]) : 対象のウィンドウハンドル
    Return      : xlwings Book または None（型は実行時のみ xlwings で解決）
    概要: 内部レジストリを確認し、無効であればその場で物理再登録を試行してブックを返却する。
    """
    # 判定: ハンドルが無効な場合（通常は発生しない）。
    if target_hwnd is None:
        # 命令分離。
        return None

    wait_xlwings_import_prewarm()
    import xlwings as xw  # 遅延 import（prewarm 済みなら即 return）

    plog = get_perf_logger(f"{__name__}.ribbon_invoke")
    t_ensure = time.perf_counter()
    plog.info("ensure_book phase=enter hwnd=%s", target_hwnd)
    try:
        # 【目的】スレッドセーフにレジストリを確認し、不整合（COM切断等）を物理的に解消するため。
        with _registry_lock:
            # 変数: 既存の登録を確認。
            book_inst = _active_books.get(target_hwnd)

            # 判定: 未登録、またはブック参照が物理的に失われている場合。
            if book_inst is None:
                try:
                    from core.core_xlc import get_excel_context_from_hwnd

                    ctx = get_excel_context_from_hwnd(int(target_hwnd), "")
                    if ctx is not None:
                        _app, book_inst, _sheet, _hwnd = ctx
                    else:
                        from xlwings._xlwindows import App as WinApp

                        app_bound = xw.App(impl=WinApp(xl=int(target_hwnd)))
                        book_inst = app_bound.books.active
                    if book_inst is None:
                        raise RuntimeError(
                            f"No active workbook for HWND: {target_hwnd}"
                        )
                    _active_books[target_hwnd] = book_inst
                    # 命令分離: 正常登録のログ記録。
                    logger.info(
                        f"Self-healing registration executed for HWND: {target_hwnd}"
                    )
                except Exception as ex_repair:
                    # 命令分離: 修復失敗の詳細なログ記録（スタックトレース付）。
                    logger.error(
                        f"Self-healing registration FAILED: {ex_repair}", exc_info=True
                    )
                    # 命令分離。
                    return None

            # 戻り値: 特定されたブックオブジェクト。
            return book_inst
    finally:
        plog.info(
            "ensure_book phase=leave hwnd=%s elapsed_ms=%d",
            target_hwnd,
            int((time.perf_counter() - t_ensure) * 1000),
        )


# =============================================================================
# invoke 一元ディスパッチ（VBA: hc_main.invoke(action=..., target_hwnd=..., sheet_id=..., **kwargs)）
# action は _INVOKE_HANDLER_MAP のキーのみ。getattr による任意実行は行わない。
# リボン customUI の tag と同一文字列であること（CSV_Tool_xml.txt）。
# =============================================================================


def _notify_wait_form_best_effort(*, parent_hwnd: int = 0) -> None:
    try:
        from core.core_cursor import notify_wait_form_ready

        notify_wait_form_ready(parent_hwnd=parent_hwnd)
    except Exception:
        pass


def _invoke_simple_svc(
    public_action: str,
    svc_action: str,
    target_hwnd: Optional[int],
    sheet_id: str,
    *,
    svc_extra: Optional[Dict[str, object]] = None,
) -> None:
    """ブック解決後に _call_svc_server へ委譲する invoke 実装の共通部（READY_UI 系以外）。"""
    book_ptr = _ensure_book(target_hwnd)
    if not book_ptr:
        return
    try:
        if svc_extra:
            _call_svc_server(svc_action, book_ptr, sheet_id, **svc_extra)
        else:
            _call_svc_server(svc_action, book_ptr, sheet_id)
    except Exception as ex:
        logger.error(
            "Module load/exec error (%s): %s", public_action, ex, exc_info=True
        )


def _invoke_csv_family(
    public_action: str,
    svc_action: str,
    target_hwnd: Optional[int],
    sheet_id: str,
    *,
    merge_enter_log: bool = False,
) -> None:
    """load/save/merge/split: 失敗・ブック欠落時は WaitForm をベストエフォートで閉じる。"""
    if merge_enter_log:
        logger.info(
            "[HC_MAIN_ENTER] action=merge_csv pid=%s ppid=%s exe=%s cwd=%s argv=%s hwnd=%s sheet_id=%s",
            os.getpid(),
            getattr(os, "getppid", lambda: -1)(),
            sys.executable,
            os.getcwd(),
            sys.argv,
            target_hwnd,
            sheet_id,
        )
    book_ptr = _ensure_book(target_hwnd)
    if book_ptr:
        try:
            _call_svc_server(svc_action, book_ptr, sheet_id)
        except Exception as ex:
            logger.error(
                "Module load/exec error (%s): %s", public_action, ex, exc_info=True
            )
            _notify_wait_form_best_effort(parent_hwnd=int(target_hwnd or 0))
        return
    if public_action == "load_csv":
        logger.error("Service ABORTED: Book context missing for HWND: %s", target_hwnd)
    else:
        logger.error(
            "Service ABORTED (%s): Book context error for HWND: %s",
            public_action,
            target_hwnd,
        )
    _notify_wait_form_best_effort(parent_hwnd=int(target_hwnd or 0))


def _invoke_impl_load_csv(target_hwnd: Optional[int], sheet_id: str, **_kw: object) -> None:
    _invoke_csv_family("load_csv", ACTION_CSV_LD, target_hwnd, sheet_id)


def _invoke_impl_save_csv(target_hwnd: Optional[int], sheet_id: str, **_kw: object) -> None:
    _invoke_csv_family("save_csv", ACTION_CSV_SV, target_hwnd, sheet_id)


def _invoke_impl_merge_csv(target_hwnd: Optional[int], sheet_id: str, **_kw: object) -> None:
    _invoke_csv_family(
        "merge_csv", ACTION_CSV_MG, target_hwnd, sheet_id, merge_enter_log=True
    )


def _invoke_impl_split_csv(target_hwnd: Optional[int], sheet_id: str, **_kw: object) -> None:
    _invoke_csv_family("split_csv", ACTION_CSV_SP, target_hwnd, sheet_id)


def _invoke_impl_check_duplicates(
    target_hwnd: Optional[int], sheet_id: str, **_kw: object
) -> None:
    # リボン本体は bridge JSON（selection_areas 含む）経路。短寿命 invoke に selection_areas は通常付かない。
    _invoke_simple_svc("check_duplicates", "dupli", target_hwnd, sheet_id)


def _invoke_impl_delete_empty_rows(
    target_hwnd: Optional[int], sheet_id: str, **_kw: object
) -> None:
    _invoke_simple_svc("delete_empty_rows", "row_dl", target_hwnd, sheet_id)


def _invoke_impl_delete_empty_cols(
    target_hwnd: Optional[int], sheet_id: str, **_kw: object
) -> None:
    _invoke_simple_svc("delete_empty_cols", "col_dl", target_hwnd, sheet_id)


def _invoke_impl_convert_date_ymd(
    target_hwnd: Optional[int], sheet_id: str, **_kw: object
) -> None:
    _invoke_simple_svc("convert_date_ymd", "dt_ymd", target_hwnd, sheet_id)


def _invoke_impl_convert_date_ymd_hm(
    target_hwnd: Optional[int], sheet_id: str, **_kw: object
) -> None:
    _invoke_simple_svc("convert_date_ymd_hm", "dt_hm", target_hwnd, sheet_id)


def _invoke_impl_trim_spaces(target_hwnd: Optional[int], sheet_id: str, **_kw: object) -> None:
    _invoke_simple_svc("trim_spaces", "trm_ex", target_hwnd, sheet_id)


def _invoke_impl_normalize_header(
    target_hwnd: Optional[int], sheet_id: str, **_kw: object
) -> None:
    _invoke_simple_svc("normalize_header", "hd_nr", target_hwnd, sheet_id)


def _invoke_impl_insert_shuka_header(
    target_hwnd: Optional[int], sheet_id: str, **_kw: object
) -> None:
    _invoke_simple_svc("insert_shuka_header", "hd_in", target_hwnd, sheet_id)


def _invoke_impl_undo_last_action(
    target_hwnd: Optional[int], sheet_id: str, **_kw: object
) -> None:
    _invoke_simple_svc("undo_last_action", "undo", target_hwnd, sheet_id)


def _invoke_impl_show_help(target_hwnd: Optional[int], sheet_id: str, **_kw: object) -> None:
    _invoke_simple_svc("show_help", "help", target_hwnd, sheet_id)


def _invoke_impl_check_for_updates(
    target_hwnd: Optional[int], sheet_id: str, **_kw: object
) -> None:
    _invoke_simple_svc("check_for_updates", "update_check", target_hwnd, sheet_id)


def _invoke_impl_run_data_agg(
    target_hwnd: Optional[int], sheet_id: str, **kwargs: object
) -> None:
    pl = kwargs.get("payload")
    extra: Optional[Dict[str, object]] = (
        {"payload": pl} if pl is not None else None
    )
    _invoke_simple_svc("run_data_agg", "data_agg", target_hwnd, sheet_id, svc_extra=extra)


_INVOKE_HANDLER_MAP: dict[str, Callable[..., None]] = {
    "load_csv": _invoke_impl_load_csv,
    "save_csv": _invoke_impl_save_csv,
    "merge_csv": _invoke_impl_merge_csv,
    "split_csv": _invoke_impl_split_csv,
    "check_duplicates": _invoke_impl_check_duplicates,
    "delete_empty_rows": _invoke_impl_delete_empty_rows,
    "delete_empty_cols": _invoke_impl_delete_empty_cols,
    "convert_date_ymd": _invoke_impl_convert_date_ymd,
    "convert_date_ymd_hm": _invoke_impl_convert_date_ymd_hm,
    "trim_spaces": _invoke_impl_trim_spaces,
    "normalize_header": _invoke_impl_normalize_header,
    "insert_shuka_header": _invoke_impl_insert_shuka_header,
    "undo_last_action": _invoke_impl_undo_last_action,
    "show_help": _invoke_impl_show_help,
    "check_for_updates": _invoke_impl_check_for_updates,
    "run_data_agg": _invoke_impl_run_data_agg,
}

if frozenset(_INVOKE_HANDLER_MAP) != RIBBON_INVOKE_ACTION_KEYS:
    raise RuntimeError(
        "ribbon_invoke _INVOKE_HANDLER_MAP keys must match core.ribbon_public_to_svc.RIBBON_INVOKE_ACTION_KEYS"
    )

INVOKE_ACTIONS: frozenset[str] = RIBBON_INVOKE_ACTION_KEYS

# invoke の finally で notify_wait_form_ready を呼ぶ action（契約は core.ribbon_public_to_svc）。
_INVOKE_NOTIFY_WAITFORM_ACTIONS: frozenset[str] = RIBBON_INVOKE_FINALLY_NOTIFY_WAITFORM


def _notify_wait_form_after_sync_invoke(action: str, *, parent_hwnd: int = 0) -> None:
    if action not in _INVOKE_NOTIFY_WAITFORM_ACTIONS:
        return
    try:
        from core.core_cursor import notify_wait_form_ready

        notify_wait_form_ready(parent_hwnd=parent_hwnd)
    except Exception:
        pass


def invoke(
    action: str,
    target_hwnd: Optional[int] = None,
    sheet_id: str = "",
    **kwargs: object,
) -> None:
    """VBA xlwings からの単一受け口。許可された action のみ実行する。

    run_data_agg では kwargs に payload を渡せる（UI 一括実行など）。
    """
    plog = get_perf_logger(f"{__name__}.ribbon_invoke")
    t0 = time.perf_counter()
    a = (action or "").strip()
    plog.info(
        "invoke phase=enter action=%r cumulative_ms=0 target_hwnd=%s sheet_id=%s",
        a,
        target_hwnd,
        sheet_id or "",
    )
    fn = _INVOKE_HANDLER_MAP.get(a)
    if fn is None:
        try:
            from core.core_cursor import notify_wait_form_ready

            notify_wait_form_ready(parent_hwnd=int(target_hwnd or 0))
        except Exception:
            pass
        allowed = ", ".join(sorted(_INVOKE_HANDLER_MAP.keys()))
        raise ValueError("ribbon_invoke.invoke: unknown action %r (allowed: %s)" % (a, allowed))
    plog.info(
        "invoke phase=before_handler action=%r cumulative_ms=%d",
        a,
        int((time.perf_counter() - t0) * 1000),
    )
    try:
        fn(target_hwnd=target_hwnd, sheet_id=sheet_id, **kwargs)
    finally:
        _notify_wait_form_after_sync_invoke(a, parent_hwnd=int(target_hwnd or 0))
        plog.info(
            "invoke phase=after_handler action=%r cumulative_ms=%d",
            a,
            int((time.perf_counter() - t0) * 1000),
        )


# =============================================================================
# サーバー管理関数 (VBA からの起動・終了制御)
# =============================================================================


def register_book(target_hwnd: Optional[int] = None) -> None:
    """
    Method Name : register_book
    Arguments   : target_hwnd (Optional[int]) : 登録対象のウィンドウハンドル
    Return      : None
    概要: アドイン起動時に VBA から呼び出される初期登録エントリ。
    """
    # 命令分離: 共通自己修復ロジックへ処理を委譲。
    # 【目的】起動のタイミングにより Excel が Ready でない場合でも、遅延登録させるため。
    _ensure_book(target_hwnd)


def clear_registry() -> None:
    """
    Method Name : clear_registry
    Arguments   : None
    Return      : None
    概要: すべてのブック参照を内部レジストリから物理抹消し、メモリ資源を解放する。
    """
    # 【目的】プロセス終了に伴い、Python 側に残った COM 参照を原子的にクリアするため。
    with _registry_lock:
        try:
            # 命令分離: 辞書の全内容を物理破棄。
            _active_books.clear()
            # 命令分離: 破棄完了のログ記録。
            logger.info("Internal reference registry cleared.")
        except Exception as ex_clear:
            # 命令分離: 異常ログ。
            logger.error(f"clear_registry error: {ex_clear}", exc_info=True)


# ==============================================================================
# メインエントリポイント (物理パス検証用)
# ==============================================================================
if __name__ == "__main__":
    # 【目的】単体起動時に Python 側のパス構成が物理的に成立しているか確認するため。
    # 命令分離: 司令塔の絶対パス。
    print(f"Commander established: {path_commander_raw_v}")
    # 命令分離: ルートディレクトリの絶対パス。
    print(f"Project root identified: {path_project_root_ptr}")