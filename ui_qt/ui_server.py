# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: ui_qt/ui_server.py
Created: 2026-02-09
Updated: 2026-05-03
Version: 1.4.57
Purpose:
  svc層からの要求(req_*.pkl)を監視し、Qtダイアログを生成・実行して結果(res_*.pkl)を返す。
  - IPC: ui_qt.ipc_file を使用（req/res/ready/shutdown）
  - Resilience: 例外でプロセスを落とさず ERROR 結果で返す
  - Qt: QApplication を必ず生成してから UI を起動する
  - Claim: pop_next_request() が既に .work.pkl を返す前提で、二重claimを防ぐ
  - Result: dlg.get_result() があれば必ず返し、files=0 を防ぐ
  - csv_ld ファイル選択OK時: 無表示1秒未満のため同一プロセスで進捗を即表示
  - 診断: HC_CSV_SP_CONFLICT_HWND_DIAG=1 で HWND ログ（GetClassName / GetWindowText / GetParent / GetAncestor）を hc_csv_diag.log へ。HC_CSV_SP_CONFLICT_HWND_DIAG_TREE=1 併用で Excel 配下 HWND 列挙 [CONFLICT_EXCEL_DESC]。
  - csv_sp_conflict: exec 直後に pump＋_close_stale_csv_sp_conflict_if_any＋再 pump でネイティブ枠を進捗前と同等以上に片付け。
  - shutdown: QTimer で shutdown.flag をポーリングし、ネスト QEventLoop（csv_mg 結合メイン）と dlg.exec 中でもループ終了＋トップレベル close。clear_shutdown_flag は mutex 取得成功後のみ（二重起動時にフラグを消さない）。

History (latest 3):
  - 1.4.60 (2026-06-06) 旧 COM WaitForm 解除（install_ribbon_startup_wait_dismiss）を削除。.ready 合図のみ。
  - 1.4.59 (2026-06-06) create_dialog 成功時に write_waitform_ready_signal（VBA DoEvents 待ち合図）。
  - 1.4.57 (2026-05-03) ui_qt.ui_help help_show: dlg.exec 直前に bump_front_follow_deferred_ensure_generation と HELP_BEFORE_MODAL_EXEC の QEventLoop 待ち（重複ジャンプ後のヘルプ前面化のため。ui_help.json の TOPMOST/FOLLOW は変更なし）。
  - 1.4.56 (2026-04-18) Windows: `ensure_ui_server_windows_dll_search_paths` に ``app`` ・ ``PySide6\\lib`` と ``PATH`` 先頭付与を追加（Shiboken.pyd の DLL 解決）。
  - 1.4.55 (2026-04-18) Windows: `shared_dll_bootstrap.ensure_ui_server_windows_dll_search_paths`（shared + EXE 直下 + shiboken6/PySide6）を PySide6 より前に実行。
  - 1.4.54 (2026-04-18) Windows: `app\\shared` への `add_dll_directory` を **PySide6 より前**に移動（Nuitka+shiboken が shared 配下にある場合の ImportError 回避）。
  - 1.4.53 (2026-04-13) create_dialog 成功後に install_ribbon_startup_wait_dismiss_on_first_show。設定エラー QMessageBox 前に notify_wait_form_ready。
  - 1.4.52 (2026-04-12) Windows: main 冒頭で FreeConsole（コンソール付き起動時のモーダル終了後フォーカスが CMD に戻る抑止）。HC_UI_KEEP_CONSOLE=1 で無効。
  - 1.4.51 (2026-04-10) progress かつ module=ui_qt.ui_dupli でも sheet_id 空を許可（重複 hlclr 進捗の dispatch 拒否を防ぐ）。
  - 1.4.50 (2026-04-10) mutex 取得成功直後に `core.ipc_cleanup.run_ui_server_startup_sweeps`（`requests` TTL・`control` の古い `*_starting.flag`）。二重起動時はスキップ。
  - 1.4.49 (2026-04-09) csv_mg 結合メイン: dlg.finished → QEventLoop.quit を QueuedConnection にし、done()/close 完了後に exec が戻る（キャンセル時の空枠ゴースト抑止）。exec 直後に processEvents を追加。
  - 1.4.48 (2026-04-09) shutdown 応答改善: 200ms ポーリングで nested QEventLoop を quit、全トップレベルを close。clear_shutdown_flag を mutex 取得後へ移動。
  - 1.4.47 (2026-04-09) csv_mg 結合メイン: QDialog.exec の代わりに show + ローカル QEventLoop（finished で終了）。IPC／戻り値契約は従来どおり。親なし exec 由来の初回位置チラつき抑止（A案）。
  - 1.4.46 (2026-04-09) csv_mg 結合メイン: exec 直前に prepare_dialog_excel_center_before_show（excel_rect）を実行。初回表示のモニタ中央→Excel 中央のチラつきを抑止。CENTER_ON_EXCEL 時のみ（_hc_csv_mg_center_on_excel）。
  - 1.4.45 (2026-04-09) csv_sp_conflict teardown: stale close を即時実行＋競合専用 pump 上限。stale 内は processEvents×4 を _pump_deferred_deletes に統一。
  - 1.4.44 (2026-04-09) CONFLICT_HWND_DIAG: Win32 クラス名・ウィンドウテキスト・親/ルート HWND。任意で DIAG_TREE で Excel 子孫列挙。
  - 1.4.43 (2026-04-09) CONFLICT_HWND_DIAG 出力先を get_diag_logger（hc_csv_diag.log）へ変更。
  - 1.4.42 (2026-04-09) HC_CSV_SP_CONFLICT_HWND_DIAG: csv_sp_conflict／進捗前後の HWND 状態ログ（ゴースト枠の切り分け用）。
  - 1.4.41 (2026-04-09) csv_sp_conflict: 結合メインと同様にアクティブ参照を保持し、csv_sp progress show 直前に hide/close/flush（_close_stale_csv_sp_conflict_if_any）。キャンセル時は参照のみ解放。
  - 1.4.40 (2026-04-09) csv_sp progress: show 直前に pump_deferred_deletes＋QTimer.singleShot(~70ms) で prepare/show を遅延（conflict 空枠 HWND の消化余地を確保）。
  - 1.4.39 (2026-04-09) csv_sp_conflict: exec 復帰後に get_result を先に確定し、その後 hide/close/deleteLater。DeferredDelete 消化のため processEvents を上限付きで追加（進捗前の空枠ゴースト低減）。
  - 1.4.38 (2026-04-09) import 時の Global\\HC_UI_SERVER を廃止。単一インスタンスは main() 内の Global\\HC_QT_UI_SERVER のみ（is_ui_server_running と整合。二重 mutex による誤起動抑止）。
  - 1.4.37 (2026-04-09) csv_sp モーダル・csv_sp 進捗: exec/show 前に prepare_dialog_excel_center_before_show。csv_sp_conflict 終了後 hide+close+deleteLater。
  - 1.4.36 (2026-04-09) csv_sp_conflict: UI_TRACE に modal_exec_ms / post_exec_teardown_ms を追加（exec 復帰と deleteLater 後の切り分け）。
  - 1.4.35 (2026-04-09) csv_sp_conflict のみ: exec 後は deleteLater＋1x processEvents（hide/close/三重 flush 省略）。枠ゴースト低減。csv_sp 分割・csv_mg は従来どおり。
  - 1.4.34 (2026-04-09) csv_sp_conflict のみ: exec＋後処理完了までの経過 ms を UI_TRACE に記録（ゴースト枠調査用）。csv_mg 無変更。
  - 1.4.33 (2026-04-08) csv_sp: 分割メイン参照を保持し、進捗 show 直前に flush（csv_mg と同様の外枠残留対策）。
  - 1.4.32 (2026-04-08) csv_mg 進捗: ui_server 側の二重 ensure_front（0/120ms）を廃止。位置調整は ProgressDialog.showEvent に一本化。
  - 1.4.31 (2026-04-05) csv_mg: 進捗表示直前に結合メイン参照で再 hide/flush（deleteLater 遅延で外枠が進捗と同居するのを防ぐ）。
"""

import os
import sys
from pathlib import Path

# ---------------------------------------------------------------------
# Bootstrap: sys.path + DLL search BEFORE PySide6/shiboken (Nuitka + app\bin)
# ---------------------------------------------------------------------
_PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(_PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(_PROJECT_ROOT))

if os.name == "nt":
    from core.shared_dll_bootstrap import ensure_ui_server_windows_dll_search_paths

    ensure_ui_server_windows_dll_search_paths()

import time
import traceback
import ctypes
from ctypes import wintypes

from dataclasses import dataclass
from typing import Any

from PySide6.QtCore import QEventLoop, Qt, QTimer
from PySide6.QtWidgets import QApplication, QMessageBox

__version__ = "1.4.59"

from ui_qt import ipc_file  # noqa: E402

from core.core_log import append_text_with_cap, get_diag_logger, get_logger  # noqa: E402

logger = get_logger("ui_qt.ui_server")
_ui_trace = get_diag_logger("hc_csv_tool.diag.ui_server")
logger.info(
    "[MODULE_LOAD] ui_qt.ui_server version=%s pid=%s file=%s",
    __version__,
    os.getpid(),
    __file__,
)

# ---------------------------------------------------------------------
# BOOT TRACE (Windows diagnostics; single-instance mutex は main() 内 HC_QT_UI_SERVER のみ)
# ---------------------------------------------------------------------
_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

_kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
_kernel32.OpenProcess.restype = wintypes.HANDLE
_kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
_kernel32.CloseHandle.restype = wintypes.BOOL
_kernel32.QueryFullProcessImageNameW.argtypes = [
    wintypes.HANDLE,
    wintypes.DWORD,
    wintypes.LPWSTR,
    ctypes.POINTER(wintypes.DWORD),
]
_kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL


def _get_process_image_path(pid: int) -> str:
    if os.name != "nt" or pid <= 0:
        return ""
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    h = _kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid)
    if not h:
        return ""
    try:
        buf = ctypes.create_unicode_buffer(32768)
        size = wintypes.DWORD(len(buf))
        ok = _kernel32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size))
        return buf.value if ok else ""
    finally:
        _kernel32.CloseHandle(h)


def _log_boot_origin() -> None:
    pid = os.getpid()
    ppid = os.getppid()
    self_img = _get_process_image_path(pid)
    parent_img = _get_process_image_path(ppid)
    logger.info(
        "[BOOT] pid=%s ppid=%s self_img=%s parent_img=%s sys.executable=%s argv=%s spawn_id=%s",
        pid,
        ppid,
        self_img,
        parent_img,
        sys.executable,
        sys.argv,
        os.environ.get("CSV_TOOL_SPAWN_ID", ""),
    )


_log_boot_origin()


def _try_detach_from_console() -> None:
    """CMD 上の python.exe で ui_server を起動したとき、モーダル終了後にコンソールが前面に出やすい。プロセスをコンソールから切り離す。"""
    if os.name != "nt":
        return
    try:
        from core import core_env

        if core_env.ui_keep_console_enabled():
            return
    except Exception:
        pass
    try:
        k32 = ctypes.windll.kernel32
        if not int(k32.GetConsoleWindow() or 0):
            return
        k32.FreeConsole()
    except Exception:
        pass


# ---------------------------------------------------------------------
# Qt bootstrap
# ---------------------------------------------------------------------
_QAPP: QApplication | None = None

# csv_mg: 結合メインは OK 後に deleteLater され、svc が即座に進捗 req を投げると外枠が残ることがある。
# 進捗 show 直前にこの参照で hide/close し、イベントを捌いてから進捗を出す。
_CSV_MG_ACTIVE_MERGE: Any = None

# csv_sp: 分割メイン（action=csv_sp）を進捗 IPC 経由で flush するための参照（結合と同様）。
_CSV_SP_ACTIVE_SPLIT: Any = None

# csv_sp: 同名確認（csv_sp_conflict）を進捗 show 直前に結合メインと同パターンで flush するための参照。
_CSV_SP_ACTIVE_CONFLICT: Any = None

# csv_sp_conflict: DeferredDelete 消化の上限（進捗経路の stale close と揃え、HWND ゴースト低減）。
_CSV_SP_CONFLICT_PUMP_ROUNDS = 60
_CSV_SP_CONFLICT_PUMP_SEC = 0.45

# shutdown.flag 検知: メイン while ループが _dispatch 内でブロック中でも終了できるようにする
_NESTED_EVENT_LOOPS: list[QEventLoop] = []
_SHUTDOWN_POLL_TIMER: QTimer | None = None
_SHUTDOWN_POLL_BUSY: bool = False


def _register_nested_loop(loop: QEventLoop) -> None:
    if loop not in _NESTED_EVENT_LOOPS:
        _NESTED_EVENT_LOOPS.append(loop)


def _unregister_nested_loop(loop: QEventLoop) -> None:
    try:
        _NESTED_EVENT_LOOPS.remove(loop)
    except ValueError:
        pass


def _quit_nested_event_loops() -> None:
    for loop in list(_NESTED_EVENT_LOOPS):
        try:
            loop.quit()
        except Exception:
            pass


def _close_all_toplevel_widgets() -> None:
    inst = QApplication.instance()
    if inst is None:
        return
    for w in list(inst.topLevelWidgets()):
        try:
            w.close()
        except Exception:
            pass


def _on_shutdown_poll() -> None:
    """shutdown.flag 成立時、ブロック中のダイアログ／ネストループを抜ける。"""
    global _SHUTDOWN_POLL_BUSY
    if _SHUTDOWN_POLL_BUSY:
        return
    if not _shutdown_requested():
        return
    _SHUTDOWN_POLL_BUSY = True
    try:
        _log(
            "INFO",
            "[ui_server] shutdown flag seen: quit nested loops + close top-level widgets",
        )
        try:
            logger.info("[UI_SERVER] shutdown flag: teardown UI (nested loops + top-level)")
        except Exception:
            pass
        _quit_nested_event_loops()
        _close_all_toplevel_widgets()
        inst = QApplication.instance()
        if inst is not None:
            try:
                inst.processEvents()
            except Exception:
                pass
    finally:
        _SHUTDOWN_POLL_BUSY = False


def _start_shutdown_poll_timer() -> None:
    global _SHUTDOWN_POLL_TIMER
    inst = QApplication.instance()
    if inst is None or _SHUTDOWN_POLL_TIMER is not None:
        return
    t = QTimer(inst)
    t.setInterval(200)
    t.timeout.connect(_on_shutdown_poll)
    t.start()
    _SHUTDOWN_POLL_TIMER = t


def _close_stale_csv_mg_merge_if_any() -> None:
    global _CSV_MG_ACTIVE_MERGE
    w = _CSV_MG_ACTIVE_MERGE
    _CSV_MG_ACTIVE_MERGE = None
    if w is None:
        return
    try:
        if hasattr(w, "hide"):
            w.hide()
        if hasattr(w, "close"):
            w.close()
    except RuntimeError:
        pass
    try:
        app = QApplication.instance()
        if app is not None:
            for _ in range(4):
                app.processEvents()
    except Exception:
        pass


def _close_stale_csv_sp_split_if_any() -> None:
    global _CSV_SP_ACTIVE_SPLIT
    w = _CSV_SP_ACTIVE_SPLIT
    _CSV_SP_ACTIVE_SPLIT = None
    if w is None:
        return
    try:
        if hasattr(w, "hide"):
            w.hide()
        if hasattr(w, "close"):
            w.close()
    except RuntimeError:
        pass
    try:
        app = QApplication.instance()
        if app is not None:
            for _ in range(4):
                app.processEvents()
    except Exception:
        pass


def _close_stale_csv_sp_conflict_if_any() -> None:
    """同名確認の保留参照を解放し hide/close/deleteLater 後、DeferredDelete を pump（二重呼び可）。"""
    global _CSV_SP_ACTIVE_CONFLICT
    w = _CSV_SP_ACTIVE_CONFLICT
    _CSV_SP_ACTIVE_CONFLICT = None
    if w is None:
        return
    try:
        if hasattr(w, "hide"):
            w.hide()
        if hasattr(w, "close"):
            w.close()
        if hasattr(w, "deleteLater"):
            w.deleteLater()
    except RuntimeError:
        pass
    try:
        app = QApplication.instance()
        if app is not None:
            _pump_deferred_deletes(
                app,
                max_rounds=_CSV_SP_CONFLICT_PUMP_ROUNDS,
                max_sec=_CSV_SP_CONFLICT_PUMP_SEC,
            )
    except Exception:
        pass


def _ensure_qapp() -> QApplication:
    """QApplication を必ず1つだけ用意する。"""
    global _QAPP
    inst = QApplication.instance()
    if inst is not None:
        _QAPP = inst
        return inst
    if _QAPP is None:
        _QAPP = QApplication(sys.argv)
    return _QAPP


# ---------------------------------------------------------------------
# Logging (best-effort)
# ---------------------------------------------------------------------
def _log(level: str, msg: str) -> None:
    try:
        log_path = ipc_file.get_server_log_path()
        stamp = time.strftime("%Y-%m-%d %H:%M:%S")
        append_text_with_cap(log_path, f"{stamp} [{level}] {msg}\n")
    except Exception:
        # ログ失敗で落とさない
        pass


# ---------------------------------------------------------------------
# IPC helpers
# ---------------------------------------------------------------------
@dataclass(frozen=True)
class _ReqPaths:
    req_path: Path
    result_path: Path
    ready_path: Path | None


def _shutdown_requested() -> bool:
    return ipc_file.get_shutdown_flag_path().exists()


def _claim_request(req_path: Path) -> Path:
    """req を追加claimする（冪等）。

    ipc_file.pop_next_request() は既に claim（.work.pkl）済みで返す前提。
    ここでは .work.pkl の二重化（.work.work.pkl）を防ぐ。
    """
    try:
        if req_path.suffix.lower() != ".pkl":
            return req_path
        if req_path.name.lower().endswith(".work.pkl"):
            return req_path
        work = req_path.with_suffix(".work.pkl")
        req_path.replace(work)
        return work
    except Exception:
        return req_path


def _extract_paths(req_path: Path, payload: dict[str, Any]) -> _ReqPaths:
    result_path = Path(str(payload["result_path"]))
    ready_raw = str(payload.get("ready_path", "")).strip()
    ready_path = Path(ready_raw) if ready_raw else None
    return _ReqPaths(req_path=req_path, result_path=result_path, ready_path=ready_path)


def _read_request_payload_with_retry(
    req_path: Path,
    *,
    max_attempts: int = 3,
) -> dict[str, Any]:
    """要求 pickle 読み込みの上位再試行。瞬間的なファイルロック揺らぎを吸収する。"""
    last_exc: Exception | None = None
    for i in range(max(1, int(max_attempts))):
        try:
            payload = ipc_file.read_pickle(req_path)
            if not isinstance(payload, dict):
                raise TypeError("payload is not dict")
            return payload
        except (PermissionError, EOFError, FileNotFoundError) as exc:
            last_exc = exc
            if i + 1 >= max_attempts:
                break
            time.sleep(min(0.08, 0.02 * (i + 1)))
    if last_exc is not None:
        raise last_exc
    raise RuntimeError("request payload read failed")


def _quarantine_bad_request(req_path: Path) -> None:
    """読込不能な要求を隔離。即削除せず診断用に残す。"""
    try:
        failed_dir = ipc_file.get_request_dir() / "_failed"
        failed_dir.mkdir(parents=True, exist_ok=True)
        ts_ms = int(time.time() * 1000)
        dst = failed_dir / f"{req_path.stem}_{ts_ms}.bad.pkl"
        req_path.replace(dst)
        try:
            logger.warning("[UI_SERVER] moved unreadable req to failed dir: %s", dst)
        except Exception:
            pass
        try:
            max_keep = 200
            raw = (os.environ.get("HC_IPC_FAILED_REQ_MAX_KEEP") or "").strip()
            if raw:
                max_keep = max(0, int(raw))
            trimmed = ipc_file.cap_failed_requests(max_keep=max_keep)
            if trimmed > 0:
                logger.info(
                    "[UI_SERVER] failed req cap trimmed=%s max_keep=%s",
                    trimmed,
                    max_keep,
                )
        except Exception:
            pass
        return
    except Exception:
        pass
    try:
        req_path.unlink(missing_ok=True)
    except Exception:
        pass


def _write_ready(path: Path) -> None:
    try:
        ipc_file.write_pickle(path, {"status": "READY_UI"})
    except Exception:
        msg = f"failed to write ready: {path}\n{traceback.format_exc()}"
        _log("ERROR", msg)
        try:
            logger.error("[UI_SERVER] %s", msg.replace("\n", " "))
        except Exception:
            pass


def _write_result(path: Path, obj: dict[str, Any]) -> None:
    try:
        ipc_file.write_pickle(path, obj)
    except Exception:
        msg = f"failed to write result: {path}\n{traceback.format_exc()}"
        _log("ERROR", msg)
        try:
            logger.error("[UI_SERVER] %s", msg.replace("\n", " "))
        except Exception:
            pass


def _get_window_rect(hwnd: int) -> list[int] | None:
    """指定 HWND の画面矩形 [left, top, right, bottom] を取得。進捗ダイアログの中央配置用。"""
    if not int(hwnd or 0) or os.name != "nt":
        return None
    try:
        r = wintypes.RECT()
        if ctypes.windll.user32.GetWindowRect(int(hwnd), ctypes.byref(r)):
            return [int(r.left), int(r.top), int(r.right), int(r.bottom)]
    except Exception:
        pass
    return None


def _error_result(message: str, exc: BaseException | None = None) -> dict[str, Any]:
    tb = ""
    if exc is not None:
        tb = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
    return {"status": "ERROR", "message": message, "traceback": tb}


def _pump_deferred_deletes(app: QApplication, *, max_rounds: int = 40, max_sec: float = 0.25) -> None:
    """deleteLater 等の遅延削除をイベントループで消化する（回数・時間の上限付き）。"""
    t0 = time.perf_counter()
    for _ in range(max_rounds):
        if time.perf_counter() - t0 > max_sec:
            break
        try:
            app.processEvents()
        except Exception:
            break


_GW_OWNER = 4
_GA_ROOT = 2
_ENUM_CHILD_PROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

_conflict_hwnd_diag_logger: Any = None


def _conflict_hwnd_diag_enabled() -> bool:
    v = (os.environ.get("HC_CSV_SP_CONFLICT_HWND_DIAG") or "").strip().lower()
    return v in ("1", "true", "yes", "on", "y")


def _conflict_hwnd_diag_tree_enabled() -> bool:
    """Excel 配下の子孫 HWND を列挙する [CONFLICT_EXCEL_DESC]（HC_CSV_SP_CONFLICT_HWND_DIAG 併用）。"""
    if not _conflict_hwnd_diag_enabled():
        return False
    v = (os.environ.get("HC_CSV_SP_CONFLICT_HWND_DIAG_TREE") or "").strip().lower()
    return v in ("1", "true", "yes", "on", "y")


def _conflict_hwnd_log_sanitize(s: str, *, max_len: int = 80) -> str:
    t = (s or "").replace("\r", " ").replace("\n", "_").strip()
    if len(t) > max_len:
        t = t[: max_len - 3] + "..."
    return t.replace(" ", "_")


def _enum_direct_child_hwnds(parent_hwnd: int) -> list[int]:
    acc: list[int] = []

    def _py_cb(child: Any, _lp: Any) -> bool:
        acc.append(int(child))
        return True

    cb = _ENUM_CHILD_PROC(_py_cb)
    try:
        ctypes.windll.user32.EnumChildWindows(wintypes.HWND(int(parent_hwnd)), cb, 0)
    except Exception:
        pass
    return acc


def _collect_excel_descendants(excel_hwnd: int, max_nodes: int) -> list[int]:
    ex = int(excel_hwnd or 0)
    if ex <= 0 or max_nodes <= 0:
        return []
    acc: list[int] = []

    def walk(parent: int) -> None:
        if len(acc) >= max_nodes:
            return
        for ch in _enum_direct_child_hwnds(parent):
            if len(acc) >= max_nodes:
                return
            acc.append(ch)
            walk(ch)

    walk(ex)
    return acc


def _win32_hwnd_probe(qhwnd: int) -> dict[str, Any]:
    """Qt winId に対応するネイティブ HWND の生存・可視・矩形・オーナー・クラス名等を調べる。"""
    out: dict[str, Any] = {
        "qhwnd": int(qhwnd or 0),
        "is_window": False,
        "is_visible": False,
        "rect": None,
        "gw_owner": 0,
        "class_name": "",
        "win_text": "",
        "parent_hwnd": 0,
        "root_hwnd": 0,
    }
    if os.name != "nt" or not out["qhwnd"]:
        return out
    try:
        user32 = ctypes.windll.user32
        h = wintypes.HWND(out["qhwnd"])
        out["is_window"] = bool(user32.IsWindow(h))
        if out["is_window"]:
            out["is_visible"] = bool(user32.IsWindowVisible(h))
            try:
                own = user32.GetWindow(h, _GW_OWNER)
                out["gw_owner"] = int(own) if own else 0
            except Exception:
                out["gw_owner"] = 0
            r = wintypes.RECT()
            if user32.GetWindowRect(h, ctypes.byref(r)):
                out["rect"] = [int(r.left), int(r.top), int(r.right), int(r.bottom)]
            buf_cls = ctypes.create_unicode_buffer(260)
            if user32.GetClassNameW(h, buf_cls, 260):
                out["class_name"] = str(buf_cls.value or "")
            buf_txt = ctypes.create_unicode_buffer(512)
            user32.GetWindowTextW(h, buf_txt, 512)
            out["win_text"] = str(buf_txt.value or "")
            try:
                p = user32.GetParent(h)
                out["parent_hwnd"] = int(p) if p else 0
            except Exception:
                out["parent_hwnd"] = 0
            try:
                rt = user32.GetAncestor(h, _GA_ROOT)
                out["root_hwnd"] = int(rt) if rt else 0
            except Exception:
                out["root_hwnd"] = 0
    except Exception as ex:
        out["probe_error"] = str(ex)
    return out


def _log_excel_descendants_snap(
    phase: str,
    excel_hwnd: int,
    *,
    sheet_id: str = "",
    source_req: str = "",
) -> None:
    """HC_CSV_SP_CONFLICT_HWND_DIAG_TREE=1 のとき、Excel 配下の子孫 HWND を 1 行で列挙（最大ノード数は上限）。"""
    global _conflict_hwnd_diag_logger
    if not _conflict_hwnd_diag_tree_enabled():
        return
    if _conflict_hwnd_diag_logger is None:
        _conflict_hwnd_diag_logger = get_diag_logger("hc_csv_tool.diag.conflict_hwnd")
    ex = int(excel_hwnd or 0)
    if not ex:
        return
    max_nodes = 48
    nodes = _collect_excel_descendants(ex, max_nodes)
    tokens: list[str] = []
    for h in nodes:
        st = _win32_hwnd_probe(h)
        cls = _conflict_hwnd_log_sanitize(st.get("class_name") or "?", max_len=40)
        vis = "1" if st.get("is_visible") else "0"
        ok = "1" if st.get("is_window") else "0"
        tokens.append(f"{h},{cls},{vis},{ok}")
    payload = ";".join(tokens)
    if len(payload) > 3800:
        payload = payload[:3790] + "...trunc"
    try:
        _conflict_hwnd_diag_logger.info(
            "[CONFLICT_EXCEL_DESC] phase=%s excel_hwnd=%s n=%s cap=%s sheet_id=%s source_req=%s items=%s",
            phase,
            ex,
            len(nodes),
            max_nodes,
            sheet_id,
            source_req or "-",
            payload,
        )
    except Exception:
        pass


def _log_conflict_hwnd_diag(
    phase: str,
    w: Any | None = None,
    *,
    qhwnd_override: int | None = None,
    sheet_id: str = "",
    source_req: str = "",
    excel_hwnd: int = 0,
    rc: int | None = None,
    extra: str = "",
) -> None:
    """HC_CSV_SP_CONFLICT_HWND_DIAG=1 のときのみ hc_csv_diag.log に 1 行出力。"""
    global _conflict_hwnd_diag_logger
    if not _conflict_hwnd_diag_enabled():
        return
    if _conflict_hwnd_diag_logger is None:
        _conflict_hwnd_diag_logger = get_diag_logger("hc_csv_tool.diag.conflict_hwnd")
    qh = int(qhwnd_override or 0)
    if qh == 0 and w is not None:
        try:
            if hasattr(w, "winId"):
                qh = int(w.winId())
        except Exception:
            qh = 0
    st = _win32_hwnd_probe(qh)
    parts = [
        f"phase={phase}",
        f"qt_winId={st.get('qhwnd')}",
        f"is_window={st.get('is_window')}",
        f"is_visible={st.get('is_visible')}",
        f"gw_owner={st.get('gw_owner')}",
        f"rect={st.get('rect')}",
    ]
    if st.get("is_window"):
        _cn = st.get("class_name") or ""
        if _cn:
            parts.append(f"cls={_conflict_hwnd_log_sanitize(_cn, max_len=56)}")
        _wt = (st.get("win_text") or "").strip()
        if _wt:
            parts.append(f"text={_conflict_hwnd_log_sanitize(_wt, max_len=72)}")
        parts.append(f"parent_hwnd={int(st.get('parent_hwnd') or 0)}")
        parts.append(f"root_hwnd={int(st.get('root_hwnd') or 0)}")
    if sheet_id:
        parts.append(f"sheet_id={sheet_id}")
    if source_req:
        parts.append(f"source_req={source_req}")
    if int(excel_hwnd or 0):
        parts.append(f"excel_hwnd={int(excel_hwnd)}")
    if rc is not None:
        parts.append(f"rc={int(rc)}")
    if st.get("probe_error"):
        parts.append(f"probe_error={st.get('probe_error')}")
    if extra:
        parts.append(extra)
    try:
        _conflict_hwnd_diag_logger.info("[CONFLICT_HWND_DIAG] %s", " ".join(parts))
    except Exception:
        pass


def _dispatch(payload: dict[str, Any], *, source_req: str = "") -> dict[str, Any]:
    """要求を処理し、結果dictを返す。"""
    global _CSV_MG_ACTIVE_MERGE, _CSV_SP_ACTIVE_SPLIT, _CSV_SP_ACTIVE_CONFLICT
    t_dispatch0 = time.perf_counter()
    req_dict_early = payload.get("req_dict")
    if not isinstance(req_dict_early, dict):
        req_dict_early = payload
    action = str(payload.get("action", "") or req_dict_early.get("action", "") or "ui_ipc").strip()
    module_name = str(payload.get("module", "")).strip() or "ui_qt.ui_csv_mg"

    parent_hwnd = int(payload.get("parent_hwnd", 0) or 0)
    sheet_id = str(payload.get("sheet_id", "")).strip()
    logger.info(
        "[UI_DISPATCH] source_req=%s action=%s module=%s parent_hwnd=%s sheet_id=%s",
        source_req or "-",
        action,
        module_name,
        parent_hwnd,
        sheet_id,
    )

    _rd_action = str(req_dict_early.get("action", "")).strip().lower()
    _ui_mod = str(payload.get("module", "")).strip()
    # データ集約の一括進捗: シート GUID 未設定でも sheet_id が空になり得る。進捗だけは許可する。
    _progress_sid_ok = _rd_action == "progress" and _ui_mod in (
        "ui_qt.ui_data_agg",
        "ui_qt.ui_dupli",
        "ui_qt.ui_row_dl",
        "ui_qt.ui_col_dl",
        "ui_qt.ui_undo",
    )

    if parent_hwnd == 0:
        logger.warning("[UI_DISPATCH] payload.parent_hwnd is empty -> Excel親子関係・タスクバー非表示不可")
        return _error_result("payload.parent_hwnd is empty")
    if sheet_id == "" and not _progress_sid_ok:
        return _error_result("payload.sheet_id is empty")
    if sheet_id == "":
        sheet_id = "_"

    try:
        mod = __import__(module_name, fromlist=["*"])
        logger.debug(
            "[UI_LOAD] module=%s file=%s has_create_dialog=%s",
            module_name,
            getattr(mod, "__file__", ""),
            hasattr(mod, "create_dialog"),
        )
    except Exception as exc:
        return _error_result(f"failed to import module: {module_name}", exc)

    if not hasattr(mod, "create_dialog"):
        return _error_result(f"module has no create_dialog: {module_name}")

    req_dict = payload.get("req_dict")
    if not isinstance(req_dict, dict):
        req_dict = payload

    try:
        _ensure_qapp()
        try:
            dlg = mod.create_dialog(req_dict, parent_hwnd, sheet_id)
        except Exception as config_exc:
            # 設定ファイル読込エラー等: エラー種別を表示して終了（救済しない）
            err_msg = str(config_exc)
            try:
                from core.core_cst import UiConfigLoadError
                if isinstance(config_exc, UiConfigLoadError):
                    err_msg = f"画面設定の読み込みに失敗しました。\n\n{config_exc}"
            except Exception:
                pass
            logger.error("[UI_DISPATCH] create_dialog failed: %s", err_msg)
            try:
                from ui_qt.ipc_file import write_waitform_ready_signal

                write_waitform_ready_signal(int(parent_hwnd or 0))
            except Exception:
                pass
            try:
                from ui_qt.ui_common import show_error_notice

                show_error_notice(None, "設定エラー", err_msg)
            except Exception:
                pass
            return _error_result(err_msg, config_exc)

        try:
            from ui_qt.ipc_file import write_waitform_ready_signal

            write_waitform_ready_signal(int(parent_hwnd or 0))
        except Exception:
            pass

        _dip_ms = int((time.perf_counter() - t_dispatch0) * 1000)
        _req_action = str(req_dict.get("action", "")).strip().lower()
        if module_name == "ui_qt.ui_csv_mg":
            if _req_action == "done_then_merge":
                _close_stale_csv_mg_merge_if_any()
            elif _req_action not in ("progress", "done_then_merge"):
                _CSV_MG_ACTIVE_MERGE = dlg
        if module_name == "ui_qt.ui_csv_sp" and _req_action == "csv_sp":
            _CSV_SP_ACTIVE_SPLIT = dlg
        if module_name == "ui_qt.ui_csv_sp" and _req_action == "csv_sp_conflict":
            _CSV_SP_ACTIVE_CONFLICT = dlg
        _modeless_hint = _req_action == "progress" or bool(req_dict.get("modeless", False))
        if module_name == "ui_qt.ui_data_agg":
            logger.info(
                "[UI_DATA_AGG] create_dialog ok source_req=%s req_action=%s sheet_id=%s hwnd=%s elapsed_ms=%d modeless_hint=%s",
                source_req or "-",
                _req_action,
                sheet_id,
                parent_hwnd,
                _dip_ms,
                _modeless_hint,
            )
            try:
                _ui_trace.info(
                    "[UI_TRACE] data_agg create_dialog ok source_req=%s req_action=%s sheet_id=%s hwnd=%s "
                    "elapsed_ms=%d modeless_hint=%s ui_pid=%s wall_perf_s=%.6f",
                    source_req or "-",
                    _req_action,
                    sheet_id,
                    parent_hwnd,
                    _dip_ms,
                    _modeless_hint,
                    os.getpid(),
                    time.perf_counter(),
                )
            except Exception:
                pass

        if module_name == "ui_qt.ui_csv_ld":
            logger.info(
                "[UI_CSV_LD] create_dialog ok source_req=%s req_action=%s sheet_id=%s hwnd=%s elapsed_ms=%d",
                source_req or "-",
                _req_action,
                sheet_id,
                parent_hwnd,
                _dip_ms,
            )
            try:
                _ui_trace.info(
                    "[UI_TRACE] csv_ld create_dialog ok source_req=%s req_action=%s sheet_id=%s hwnd=%s "
                    "elapsed_ms=%d ui_pid=%s wall_perf_s=%.6f",
                    source_req or "-",
                    _req_action,
                    sheet_id,
                    parent_hwnd,
                    _dip_ms,
                    os.getpid(),
                    time.perf_counter(),
                )
            except Exception:
                pass

        if module_name == "ui_qt.ui_csv_sv":
            logger.info(
                "[UI_CSV_SV] create_dialog ok source_req=%s req_action=%s sheet_id=%s hwnd=%s elapsed_ms=%d",
                source_req or "-",
                _req_action,
                sheet_id,
                parent_hwnd,
                _dip_ms,
            )
            try:
                _ui_trace.info(
                    "[UI_TRACE] csv_sv create_dialog ok source_req=%s req_action=%s sheet_id=%s hwnd=%s "
                    "elapsed_ms=%d ui_pid=%s wall_perf_s=%.6f",
                    source_req or "-",
                    _req_action,
                    sheet_id,
                    parent_hwnd,
                    _dip_ms,
                    os.getpid(),
                    time.perf_counter(),
                )
            except Exception:
                pass

        if module_name == "ui_qt.ui_csv_mg":
            logger.info(
                "[UI_CSV_MG] create_dialog ok source_req=%s req_action=%s sheet_id=%s hwnd=%s elapsed_ms=%d",
                source_req or "-",
                _req_action,
                sheet_id,
                parent_hwnd,
                _dip_ms,
            )
            try:
                _ui_trace.info(
                    "[UI_TRACE] csv_mg create_dialog ok source_req=%s req_action=%s sheet_id=%s hwnd=%s "
                    "elapsed_ms=%d ui_pid=%s wall_perf_s=%.6f",
                    source_req or "-",
                    _req_action,
                    sheet_id,
                    parent_hwnd,
                    _dip_ms,
                    os.getpid(),
                    time.perf_counter(),
                )
            except Exception:
                pass

        if module_name == "ui_qt.ui_csv_sp":
            logger.info(
                "[UI_CSV_SP] create_dialog ok source_req=%s req_action=%s sheet_id=%s hwnd=%s elapsed_ms=%d",
                source_req or "-",
                _req_action,
                sheet_id,
                parent_hwnd,
                _dip_ms,
            )
            try:
                _ui_trace.info(
                    "[UI_TRACE] csv_sp create_dialog ok source_req=%s req_action=%s sheet_id=%s hwnd=%s "
                    "elapsed_ms=%d ui_pid=%s wall_perf_s=%.6f",
                    source_req or "-",
                    _req_action,
                    sheet_id,
                    parent_hwnd,
                    _dip_ms,
                    os.getpid(),
                    time.perf_counter(),
                )
            except Exception:
                pass

        # modeless support (e.g., progress window)
        is_modeless = False
        try:
            a = str(req_dict.get("action", "")).strip().lower()
            # 'progress' は結果を待たない（ui_server をブロックしない）
            if a == "progress":
                is_modeless = True
                logger.debug("[CSV_LD_FLOW] ui_server: progress request received, creating dialog t=%.3f", time.time())
            # 明示指定も許可
            if bool(req_dict.get("modeless", False)):
                is_modeless = True
        except Exception:
            is_modeless = False

        _csv_sp_conflict_got: dict[str, Any] | None = None

        if is_modeless:
            # show() できる場合は表示して即返す
            try:
                if _req_action == "progress" and module_name == "ui_qt.ui_csv_mg":
                    _close_stale_csv_mg_merge_if_any()
                if _req_action == "progress" and module_name == "ui_qt.ui_csv_sp":
                    _conflict_qhwnd_snapshot = 0
                    if _conflict_hwnd_diag_enabled():
                        try:
                            _cref = _CSV_SP_ACTIVE_CONFLICT
                            _log_conflict_hwnd_diag(
                                "progress_before_conflict_flush",
                                _cref,
                                sheet_id=sheet_id,
                                source_req=source_req or "",
                                excel_hwnd=int(parent_hwnd or 0),
                            )
                            if _cref is not None and hasattr(_cref, "winId"):
                                _conflict_qhwnd_snapshot = int(_cref.winId())
                        except Exception:
                            _conflict_qhwnd_snapshot = 0
                    _log_excel_descendants_snap(
                        "progress_before_conflict_flush_tree",
                        int(parent_hwnd or 0),
                        sheet_id=sheet_id,
                        source_req=source_req or "",
                    )
                    _close_stale_csv_sp_conflict_if_any()
                    if _conflict_hwnd_diag_enabled():
                        if _conflict_qhwnd_snapshot:
                            _log_conflict_hwnd_diag(
                                "progress_after_conflict_flush",
                                None,
                                qhwnd_override=_conflict_qhwnd_snapshot,
                                sheet_id=sheet_id,
                                source_req=source_req or "",
                                excel_hwnd=int(parent_hwnd or 0),
                                extra="reprobe_same_qt_winId",
                            )
                        else:
                            _log_conflict_hwnd_diag(
                                "progress_after_conflict_flush",
                                None,
                                qhwnd_override=0,
                                sheet_id=sheet_id,
                                source_req=source_req or "",
                                excel_hwnd=int(parent_hwnd or 0),
                                extra="no_conflict_ref_before_flush",
                            )
                    _log_excel_descendants_snap(
                        "progress_after_conflict_flush_tree",
                        int(parent_hwnd or 0),
                        sheet_id=sheet_id,
                        source_req=source_req or "",
                    )
                    _close_stale_csv_sp_split_if_any()
            except Exception:
                pass
            try:
                _csv_sp_progress_deferred = (
                    _req_action == "progress" and module_name == "ui_qt.ui_csv_sp"
                )
                if _csv_sp_progress_deferred:
                    _app_sp = QApplication.instance()
                    if _app_sp is not None:
                        _pump_deferred_deletes(_app_sp)
                    from PySide6.QtCore import QTimer

                    def _do_csv_sp_progress_show() -> None:
                        try:
                            from ui_qt.ui_common import (
                                excel_rect_tuple_from_req,
                                prepare_dialog_excel_center_before_show,
                            )

                            prepare_dialog_excel_center_before_show(
                                dlg,
                                parent_hwnd,
                                excel_rect_tuple_from_req(req_dict),
                                getattr(dlg, "_hc_prepare_window_cfg", None),
                            )
                        except Exception:
                            pass
                        try:
                            if hasattr(dlg, "show"):
                                dlg.show()
                                logger.debug(
                                    "[CSV_LD_FLOW] ui_server: csv_sp progress show() deferred t=%.3f",
                                    time.time(),
                                )
                        except Exception:
                            pass
                        if _conflict_hwnd_diag_enabled():
                            _log_conflict_hwnd_diag(
                                "progress_after_deferred_show",
                                dlg,
                                sheet_id=sheet_id,
                                source_req=source_req or "",
                                excel_hwnd=int(parent_hwnd or 0),
                                extra="modeless_csv_sp_progress",
                            )
                        _log_excel_descendants_snap(
                            "progress_after_deferred_show_tree",
                            int(parent_hwnd or 0),
                            sheet_id=sheet_id,
                            source_req=source_req or "",
                        )

                    QTimer.singleShot(70, _do_csv_sp_progress_show)
                else:
                    if hasattr(dlg, "show"):
                        dlg.show()
                        if a == "progress":
                            logger.debug(
                                "[CSV_LD_FLOW] ui_server: progress dialog show() called t=%.3f",
                                time.time(),
                            )
            except Exception:
                pass
            rc = 0
        else:
            _sp_conflict_lifecycle_t0 = None
            if module_name == "ui_qt.ui_csv_sp" and _req_action == "csv_sp_conflict":
                _sp_conflict_lifecycle_t0 = time.perf_counter()
            if module_name == "ui_qt.ui_csv_sp" and _req_action in ("csv_sp", "csv_sp_conflict"):
                try:
                    from ui_qt.ui_common import (
                        excel_rect_tuple_from_req,
                        prepare_dialog_excel_center_before_show,
                    )

                    prepare_dialog_excel_center_before_show(
                        dlg,
                        parent_hwnd,
                        excel_rect_tuple_from_req(req_dict),
                        getattr(dlg, "_hc_prepare_window_cfg", None),
                    )
                except Exception:
                    pass
            if (
                module_name == "ui_qt.ui_csv_mg"
                and _req_action not in ("progress", "done_then_merge")
            ):
                try:
                    _co = dlg.property("_hc_csv_mg_center_on_excel")
                    if bool(_co):
                        from ui_qt.ui_common import (
                            excel_rect_tuple_from_req,
                            prepare_dialog_excel_center_before_show,
                        )

                        prepare_dialog_excel_center_before_show(
                            dlg,
                            parent_hwnd,
                            excel_rect_tuple_from_req(req_dict),
                            getattr(dlg, "_hc_prepare_window_cfg", None),
                        )
                except Exception:
                    pass
            if module_name == "ui_qt.ui_csv_sp" and _req_action == "csv_sp_conflict":
                _log_conflict_hwnd_diag(
                    "excel_parent_probe",
                    None,
                    qhwnd_override=int(parent_hwnd or 0),
                    sheet_id=sheet_id,
                    source_req=source_req or "",
                    excel_hwnd=int(parent_hwnd or 0),
                    extra="note=ipc_parent_hwnd",
                )
                _log_conflict_hwnd_diag(
                    "conflict_before_exec",
                    dlg,
                    sheet_id=sheet_id,
                    source_req=source_req or "",
                    excel_hwnd=int(parent_hwnd or 0),
                )
            _csv_mg_merge_show_loop = (
                module_name == "ui_qt.ui_csv_mg"
                and _req_action not in ("progress", "done_then_merge")
                and hasattr(dlg, "show")
                and hasattr(dlg, "finished")
            )
            if _csv_mg_merge_show_loop:
                _mg_loop = QEventLoop()
                _register_nested_loop(_mg_loop)

                def _csv_mg_on_finished(*_args: object) -> None:
                    _mg_loop.quit()

                try:
                    try:
                        if hasattr(dlg, "setModal"):
                            dlg.setModal(True)
                    except Exception:
                        pass
                    # DirectConnection だと done()/close 途中で quit が走り、teardown と競って空枠が残る
                    dlg.finished.connect(
                        _csv_mg_on_finished,
                        Qt.ConnectionType.QueuedConnection,
                    )
                    dlg.show()
                    _mg_loop.exec()
                finally:
                    _unregister_nested_loop(_mg_loop)
                    try:
                        dlg.finished.disconnect(_csv_mg_on_finished)
                    except Exception:
                        pass
                try:
                    _app_mg = QApplication.instance()
                    if _app_mg is not None:
                        for _ in range(12):
                            _app_mg.processEvents()
                except Exception:
                    pass
                try:
                    rc = int(dlg.result())
                except Exception:
                    rc = 0
            elif hasattr(dlg, "exec"):
                if module_name == "ui_qt.ui_help" and _req_action == "help_show":
                    try:
                        from core.ui_window_timing import get_ui_window_timings

                        _help_pre_ms = int(get_ui_window_timings().help_before_modal_exec_delay_ms)
                        if _help_pre_ms > 0:
                            _help_pre_loop = QEventLoop()
                            QTimer.singleShot(_help_pre_ms, _help_pre_loop.quit)
                            _help_pre_loop.exec()
                        else:
                            _app_hpre = QApplication.instance()
                            if _app_hpre is not None:
                                _app_hpre.processEvents()
                    except Exception:
                        pass
                rc = dlg.exec()
            elif hasattr(dlg, "exec_"):
                rc = dlg.exec_()
            else:
                rc = 0
            _sp_conflict_t_after_exec = time.perf_counter()
            if module_name == "ui_qt.ui_csv_sp" and _req_action == "csv_sp_conflict":
                _log_conflict_hwnd_diag(
                    "conflict_after_exec",
                    dlg,
                    sheet_id=sheet_id,
                    source_req=source_req or "",
                    excel_hwnd=int(parent_hwnd or 0),
                    rc=int(rc),
                )
            # モーダル終了後はダイアログを閉じる。csv_mg/csv_sp は外枠残留を防ぐため hide＋イベント処理を先に行う。
            # csv_sp_conflict: get_result を先に確定してから teardown（C++ 側削除後の get_result 回避）。
            # deleteLater 後は DeferredDelete 用に processEvents を十分に回し、進捗 show 前に空枠 HWND を減らす。
            try:
                if module_name == "ui_qt.ui_csv_sp" and _req_action == "csv_sp_conflict":
                    if hasattr(dlg, "get_result"):
                        try:
                            _gr_cf = dlg.get_result()
                            if isinstance(_gr_cf, dict):
                                _csv_sp_conflict_got = dict(_gr_cf)
                        except Exception:
                            _csv_sp_conflict_got = None
                    try:
                        if hasattr(dlg, "hide"):
                            dlg.hide()
                    except Exception:
                        pass
                    try:
                        if hasattr(dlg, "close"):
                            dlg.close()
                    except Exception:
                        pass
                    if hasattr(dlg, "deleteLater"):
                        dlg.deleteLater()
                    _app_cf = QApplication.instance()
                    if _app_cf is not None:
                        _pump_deferred_deletes(
                            _app_cf,
                            max_rounds=_CSV_SP_CONFLICT_PUMP_ROUNDS,
                            max_sec=_CSV_SP_CONFLICT_PUMP_SEC,
                        )
                    # 進捗 show 直前の _close_stale_csv_sp_conflict_if_any と同等以上を、診断ログより前に実行（二重呼び可）。
                    _close_stale_csv_sp_conflict_if_any()
                    if _app_cf is not None:
                        _pump_deferred_deletes(
                            _app_cf,
                            max_rounds=_CSV_SP_CONFLICT_PUMP_ROUNDS,
                            max_sec=_CSV_SP_CONFLICT_PUMP_SEC,
                        )
                    _log_conflict_hwnd_diag(
                        "conflict_after_teardown_pump",
                        dlg,
                        sheet_id=sheet_id,
                        source_req=source_req or "",
                        excel_hwnd=int(parent_hwnd or 0),
                        rc=int(rc),
                    )
                    _log_excel_descendants_snap(
                        "conflict_after_teardown_pump_tree",
                        int(parent_hwnd or 0),
                        sheet_id=sheet_id,
                        source_req=source_req or "",
                    )
                else:
                    if module_name in ("ui_qt.ui_csv_mg", "ui_qt.ui_csv_sp"):
                        if hasattr(dlg, "hide"):
                            dlg.hide()
                        _app = QApplication.instance()
                        if _app is not None:
                            _app.processEvents()
                    if hasattr(dlg, "close"):
                        dlg.close()
                    if hasattr(dlg, "deleteLater"):
                        dlg.deleteLater()
                    if module_name in ("ui_qt.ui_csv_mg", "ui_qt.ui_csv_sp"):
                        _app2 = QApplication.instance()
                        if _app2 is not None:
                            for _ in range(3):
                                _app2.processEvents()
                            if module_name == "ui_qt.ui_csv_mg":
                                _pump_deferred_deletes(
                                    _app2,
                                    max_rounds=32,
                                    max_sec=0.25,
                                )
            except Exception:
                pass
            _sp_conflict_t_after_teardown = time.perf_counter()
            if module_name == "ui_qt.ui_csv_mg" and int(rc) != 1:
                _CSV_MG_ACTIVE_MERGE = None
            if module_name == "ui_qt.ui_csv_sp":
                if _req_action == "csv_sp" or int(rc) != 1:
                    _CSV_SP_ACTIVE_SPLIT = None
                if _req_action == "csv_sp_conflict" and int(rc) != 1:
                    _CSV_SP_ACTIVE_CONFLICT = None
            if _sp_conflict_lifecycle_t0 is not None:
                try:
                    _elapsed_cf = int((_sp_conflict_t_after_teardown - _sp_conflict_lifecycle_t0) * 1000)
                    _modal_ms = int((_sp_conflict_t_after_exec - _sp_conflict_lifecycle_t0) * 1000)
                    _teardown_ms = int((_sp_conflict_t_after_teardown - _sp_conflict_t_after_exec) * 1000)
                    _ui_trace.info(
                        "[UI_TRACE] csv_sp_conflict lifecycle_end rc=%s exec_plus_teardown_ms=%s "
                        "modal_exec_ms=%s post_exec_teardown_ms=%s "
                        "sheet_id=%s hwnd=%s source_req=%s ui_pid=%s wall_perf_s=%.6f",
                        int(rc),
                        _elapsed_cf,
                        _modal_ms,
                        _teardown_ms,
                        sheet_id,
                        parent_hwnd,
                        source_req or "-",
                        os.getpid(),
                        time.perf_counter(),
                    )
                except Exception:
                    pass

        # UIが結果辞書を保持している場合は、必ず返す（files=0対策）
        # キャンセル時(rc!=Accepted): 結合処理を実行せず、status=CANCEL と files=[] で返す
        accepted = 1  # QDialog.DialogCode.Accepted
        if _csv_sp_conflict_got is not None:
            try:
                got = _csv_sp_conflict_got
                got.setdefault("rc", int(rc))
                if int(rc) != accepted:
                    got["status"] = "CANCEL"
                    got["files"] = []
                else:
                    got.setdefault("status", "OK")
                return got
            except Exception:
                pass
        if hasattr(dlg, "get_result"):
            try:
                got = dlg.get_result()
                if isinstance(got, dict):
                    got.setdefault("rc", int(rc))
                    if int(rc) != accepted:
                        got["status"] = "CANCEL"
                        got["files"] = []
                    else:
                        got.setdefault("status", "OK")
                    # csv_ld でファイル選択OK時: 無表示時間を1秒未満にするため、同一プロセスで進捗を即表示
                    if action == "csv_ld" and got.get("status") == "OK":
                        try:
                            root = Path(str(ipc_file.get_ipc_root()))
                            progress_path = root / "progress" / f"progress_ld_{sheet_id}.pkl"
                            progress_path.parent.mkdir(parents=True, exist_ok=True)
                            ipc_file.write_pickle(
                                progress_path,
                                {
                                    "status": "RUN",
                                    "phase_i": 0,
                                    "phase": "準備中...",
                                    "done": 0,
                                    "total": 0,
                                    "pct": 0,
                                    "current_file": "",
                                    "seq": 0,
                                },
                            )
                            progress_req_dict = {
                                "action": "progress",
                                "progress_path": str(progress_path),
                                "phase_total": 4,
                                "excel_lock": True,
                                "no_native_window": True,
                                "progress_poll_ms": 40,
                                "progress_bar_creep_pct": 2,
                            }
                            excel_rect = _get_window_rect(parent_hwnd)
                            if excel_rect is not None:
                                progress_req_dict["excel_rect"] = excel_rect
                            progress_dlg = mod.create_dialog(
                                progress_req_dict, parent_hwnd, sheet_id
                            )
                            if hasattr(progress_dlg, "show"):
                                progress_dlg.show()
                                logger.debug(
                                    "[CSV_LD_FLOW] ui_server: progress shown immediately after file picker OK t=%.3f",
                                    time.time(),
                                )
                        except Exception as prog_exc:
                            logger.warning(
                                "[CSV_LD_FLOW] ui_server: immediate progress show failed: %s",
                                prog_exc,
                            )
                    return got
            except Exception:
                pass

        if int(rc) != accepted:
            return {"status": "CANCEL", "rc": int(rc), "files": []}
        return {"status": "OK", "rc": int(rc)}
    except Exception as exc:
        return _error_result("create_dialog failed", exc)


def main() -> int:
    """Qt UI 常駐プロセスのエントリ: IPC リクエストを受け取りダイアログで処理して結果を返す。"""
    global _SHUTDOWN_POLL_TIMER
    _try_detach_from_console()
    # best-effort diagnostics
    try:
        ipc_file.log_module_loaded(__name__, __file__, __version__)
    except Exception:
        pass

    _log(
        "INFO",
        f"[ui_server] started version={__version__} pid={os.getpid()} ppid={os.getppid()} file={__file__} cwd={os.getcwd()}",
    )

    # multi-instance guard（二重起動時は shutdown.flag を消さない → 生存中プロセスが終了要求を取りこぼさない）
    mutex_handle: int | None = None
    try:
        handle, already = ipc_file.create_single_instance_mutex()
        mutex_handle = handle
        if already:
            _log("INFO", "already running (mutex exists) -> exit")
            return 0
        try:
            from core.ipc_cleanup import run_ui_server_startup_sweeps

            run_ui_server_startup_sweeps(ipc_file.get_ipc_root())
        except Exception:
            pass
    except Exception:
        msg = "failed to create mutex\n" + traceback.format_exc()
        _log("ERROR", msg)
        try:
            logger.error("[UI_SERVER] %s", msg.replace("\n", " "))
        except Exception:
            pass

    try:
        ipc_file.clear_shutdown_flag()
    except Exception:
        pass

    try:
        _ensure_qapp()
        _start_shutdown_poll_timer()

        inst = QApplication.instance()
        next_failed_cleanup_at = time.monotonic()

        while True:
            # Pump Qt events so modeless windows (show) can paint and timers fire.
            if inst is not None:
                try:
                    inst.processEvents()
                except Exception:
                    pass

            if _shutdown_requested():
                _log("INFO", "shutdown requested")
                break

            req = ipc_file.pop_next_request()
            if req is None:
                now_mono = time.monotonic()
                if now_mono >= next_failed_cleanup_at:
                    try:
                        n_removed = ipc_file.cleanup_failed_requests(
                            ttl_sec=24 * 60 * 60,
                            max_remove=10,
                        )
                        if n_removed > 0:
                            logger.info(
                                "[UI_SERVER] cleaned failed requests count=%s",
                                n_removed,
                            )
                    except Exception:
                        pass
                    next_failed_cleanup_at = now_mono + 60.0
                if inst is not None:
                    try:
                        inst.processEvents()
                    except Exception:
                        pass
                time.sleep(0.02)
                continue

            req_path = _claim_request(Path(str(req)))

            try:
                payload = _read_request_payload_with_retry(req_path)
                paths = _extract_paths(req_path, payload)
            except Exception:
                msg = f"failed to read req: {req_path}\n{traceback.format_exc()}"
                _log("ERROR", msg)
                try:
                    logger.error("[UI_SERVER] %s", msg.replace("\n", " "))
                except Exception:
                    pass
                _quarantine_bad_request(req_path)
                continue

            if paths.ready_path is not None:
                _write_ready(paths.ready_path)

            res = _dispatch(payload, source_req=req_path.name)
            _write_result(paths.result_path, res)

            if inst is not None:
                try:
                    inst.processEvents()
                except Exception:
                    pass

            try:
                req_path.unlink(missing_ok=True)
            except Exception:
                pass

    finally:
        try:
            if _SHUTDOWN_POLL_TIMER is not None:
                _SHUTDOWN_POLL_TIMER.stop()
                _SHUTDOWN_POLL_TIMER = None
        except Exception:
            pass
        try:
            if mutex_handle is not None:
                ipc_file.release_mutex(mutex_handle)
        except Exception:
            pass

    _log("INFO", "[ui_server] exited")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
