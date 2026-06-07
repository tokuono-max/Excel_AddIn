# -*- coding: utf-8 -*-
"""
Python: 3.10+
Module: svc/svc_undo
Created: 2026-03-05
Version: 1.7.8
Purpose:
  元に戻す（Undo）：Pickle キャッシュから直前状態を復元する（新方式: svc_server + book 渡し）。
  共通仕様（docs/共通仕様_機能.md）に準拠。復元中は ui_server 経由で進捗（config/ui_undo.json PROGRESS）。
  復元できない場合は既存のステータス通知に加え ui_qt.ui_undo の専用ダイアログで画面表示する。

  Undo情報の扱い（1回だけ戻す）:
  - キー: _make_undo_key(hwnd, ブック名, シート名) で同一ブック・同一シートで1スナップショット。
  - 上書き: 新しいUndo対象の機能（例: hd_nrで「開始」）が実行されると save_undo_snapshot が呼ばれ、その時点の状態で上書き保存される。
  - 削除: exec_undo で元に戻した直後に CacheManager.delete(str_undo_key) で当該キーを削除する。戻した後は同じ操作で再度Undoできない。
  - まとめ: 「直前に実行した1つの機能」分だけ戻せる。戻すとスナップショットは消える。別の機能を実行するとスナップショットは新しい状態で上書きされる。

History (latest 3):
  - 1.7.8 (2026-06-06) 進捗表示中は excel_lock=True（復元開始から Excel 操作無効）。完了時 teardown で解除。
  - 1.7.7 exec_undo finally: 成功時も常に w32.bring_to_front(hwnd) を実行（1.7.5 のスキップを撤去）。TOPMOST 進捗後の前景を他機能（例 svc_dt_ymd）と揃え、続くモーダル前面の切り分け用。再発時はリバートまたは進捗側の TOPMOST 解除等を検討。
  - 1.7.6 Undo 成功: キャッシュ削除を進捗フェーズ（PHASE_UNDO_CACHE_DELETE）に含め、削除完了後に _undo_progress_done。進捗内の早期 DONE 呼び出しを廃止。_undo_progress_done の done_delay_ms / sleep を短縮。
  - 1.7.5 exec_undo finally: 復元成功かつ _show_undo_done_dialog を表示した直後は w32.bring_to_front(hwnd) をスキップ（完了 OK 後に Excel を強制前面にすると続く UI が背後に回る事象の緩和）。失敗・例外経路では従来どおり bring_to_front。（1.7.7 で撤去）
  - 1.7.4 Undo 進捗 IPC: excel_lock を False に変更（CSV 読込等の進捗と同様。ProgressDialog の enable_excel_window(False) を踏まない）。復元中の操作抑止は Interactive=False 等に依存。
  - 1.7.3 データ復元: スナップショットが空リスト（使用範囲なし）でも復元可能に（出荷履歴項目追加の Undo 等）。
  - 1.7.2 データ復元: 進捗表示直後（with 突入前）から Interactive=False。進捗～Excel 操作無効の隙間を解消。
  - 1.7.1 Undo 進捗: excel_lock=True（復元中は enable_excel_window 無効化）。構造復元も Interactive=False。終了時は進捗 close＋finally で有効化。
  - 1.7.0 exec_undo: ui_server 経由の進捗ダイアログ（config/ui_undo.json PROGRESS）。データ復元・構造復元の両方。キャンセルなし。
  - 1.6.1 HC_LOG_PERF: [UNDO_PERF]。診断: [UNDO_TRACE]（exec_undo）。
  - 1.6.0 表示更新を xlc.suspend_sheet_updates で一括制御。復元成功時に終了通知を表示（config/ui_undo.json SCREENS.UNDO_DONE）。
  - 1.5.3 枠固定をスナップショットに含め復元時に解除→データ復元後に再設定。
  - 1.5.2 復元後、不要列の列幅をシートの標準幅に戻す。
"""
from __future__ import annotations

import os
import re
import sys
import time
from contextlib import nullcontext
from pathlib import Path
from typing import Any, List, Optional

_path_here = os.path.abspath(os.path.dirname(__file__))
_path_root = os.path.dirname(_path_here)
if _path_root not in sys.path:
    sys.path.insert(0, _path_root)

from core.core_log import get_diag_logger, get_logger, get_perf_logger

logger = get_logger(__name__)
_undo_diag = get_diag_logger("hc_csv_tool.diag.undo")
_undo_perf = get_perf_logger("svc.svc_undo.perf")
__version__ = "1.7.8"


def _elapsed_ms(since: float) -> int:
    return max(0, int((time.perf_counter() - since) * 1000))


def _undo_trace(phase: str, t0: float, **kv: object) -> None:
    try:
        if kv:
            _undo_diag.info(
                "[UNDO_TRACE] phase=%s cumulative_ms=%d %s",
                phase,
                _elapsed_ms(t0),
                " ".join("%s=%s" % (k, v) for k, v in kv.items()),
            )
        else:
            _undo_diag.info("[UNDO_TRACE] phase=%s cumulative_ms=%d", phase, _elapsed_ms(t0))
    except Exception:
        pass


def _log_undo_perf(phase: str, t0: float, **kv: object) -> None:
    try:
        if kv:
            _undo_perf.info(
                "[UNDO_PERF] phase=%s cumulative_ms=%d %s",
                phase,
                _elapsed_ms(t0),
                " ".join("%s=%s" % (k, v) for k, v in kv.items()),
            )
        else:
            _undo_perf.info("[UNDO_PERF] phase=%s cumulative_ms=%d", phase, _elapsed_ms(t0))
    except Exception:
        pass

try:
    from ui_qt import ipc_file
    from svc.svc_host import ensure_ui_server
except ImportError:
    ipc_file = None  # type: ignore[assignment]
    ensure_ui_server = None  # type: ignore[assignment]

try:
    from core import core_cst as cst
    from core import core_xlc as xlc
    from core import core_stat
    from core import core_sys as hsys
    from core import core_w32 as w32
except ImportError:
    cst = None  # type: ignore[assignment]
    xlc = None  # type: ignore[assignment]
    core_stat = None  # type: ignore[assignment]
    hsys = None  # type: ignore[assignment]
    w32 = None  # type: ignore[assignment]


def _get_window_rect(hwnd: int) -> tuple[int, int, int, int] | None:
    """Excel HWND の GetWindowRect。UI の excel_rect 用。"""
    if not int(hwnd or 0) or os.name != "nt":
        return None
    try:
        import ctypes
        from ctypes import wintypes

        r = wintypes.RECT()
        if ctypes.windll.user32.GetWindowRect(int(hwnd), ctypes.byref(r)):
            return (int(r.left), int(r.top), int(r.right), int(r.bottom))
    except Exception:
        pass
    return None


def _undo_progress_file_path() -> Path:
    base = Path(os.environ.get("TEMP", ".")) / "csv_tool" / "progress"
    if ipc_file is not None:
        try:
            base = Path(str(ipc_file.get_ipc_root())) / "progress"
        except Exception:
            pass
    base.mkdir(parents=True, exist_ok=True)
    return base / f"progress_undo_{os.getpid()}_{int(time.time() * 1000)}.pkl"


def _undo_progress_write(path: Path, obj: dict[str, Any]) -> None:
    try:
        if ipc_file is None:
            return
        obj = dict(obj)
        if "seq" not in obj:
            try:
                prev = ipc_file.read_pickle(path)
                nxt = int(prev.get("seq", -1)) + 1 if isinstance(prev, dict) else 0
            except Exception:
                nxt = 0
            obj["seq"] = nxt
        ipc_file.write_pickle(path, obj)
    except Exception:
        pass


def _undo_progress_phase(
    path: Path,
    seq: list[int],
    pct: int,
    phase: str,
    *,
    cur: str = "",
) -> None:
    seq[0] += 1
    p = max(0, min(100, int(pct)))
    _undo_progress_write(
        path,
        {
            "status": "RUN",
            "phase": phase,
            "msg": phase,
            "current_file": cur or phase,
            "pct": p,
            "done": p,
            "total": 100,
            "seq": seq[0],
        },
    )


def _undo_progress_done(path: Path | None) -> None:
    if path is None:
        return
    try:
        _undo_progress_write(
            path,
            {
                "status": "DONE",
                "phase": "復元処理が完了しました",
                "msg": "復元処理が完了しました",
                "pct": 100,
                "done": 100,
                "total": 100,
                "seq": 999,
                "done_delay_ms": 200,
            },
        )
        time.sleep(0.12)
    except Exception:
        pass


def _undo_msg(key: str, default: str) -> str:
    try:
        if cst is None:
            return default
        cfg = cst.get_ui_config_from_file_required("undo")
        raw = (cfg.get("MESSAGES") or {}).get(key)
        return str(raw).strip() if raw else default
    except Exception:
        return default


def _submit_undo_progress_ui(
    parent_hwnd: int,
    sheet_id: str,
    progress_path: Path,
    phase_total: int,
) -> None:
    if ipc_file is None or ensure_ui_server is None:
        return
    try:
        ensure_ui_server()
    except Exception as e:
        logger.warning("[UNDO] ensure_ui_server 失敗: %s", e)
        return
    try:
        root = Path(str(ipc_file.get_ipc_root()))
        res_dir = root / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        ts_ms = int(time.time() * 1000)
        result_path = str(res_dir / f"res_undo_progress_{ts_ms}_{os.getpid()}.pkl")
        sid = (sheet_id or "").strip() or "undo"
        er = _get_window_rect(int(parent_hwnd or 0))
        req_dict: dict[str, Any] = {
            "action": "progress",
            "progress_path": str(progress_path),
            "phase_total": int(phase_total),
            "excel_lock": True,
            "no_native_window": True,
        }
        if er is not None:
            req_dict["excel_rect"] = list(er)
        payload = {
            "parent_hwnd": int(parent_hwnd),
            "result_path": result_path,
            "ready_path": "",
            "sheet_id": sid,
            "action": "progress",
            "module": "ui_qt.ui_undo",
            "req_dict": req_dict,
        }
        req_path = ipc_file.get_request_dir() / f"req_undo_progress_{ts_ms}_{os.getpid()}.pkl"
        ipc_file.write_pickle(req_path, payload)
        logger.info("[UNDO] progress UI request: %s", req_path)
    except Exception as exc:
        logger.warning("[UNDO] progress UI request failed: %s", exc)


def _make_undo_key(hwnd: int, wb_name: str, sh_name: str) -> str:
    """
    exec_undo と save_undo_snapshot で共通のキャッシュキーを生成する。
    PID は含めないため、別プロセス（svc_server 再起動後など）からも同一キーで参照可能。

    Args:
        hwnd: Excel ウィンドウのハンドル（同一インスタンス識別用）。
        wb_name: ブック名。
        sh_name: シート名。

    Returns:
        ファイルシステムで使用可能な形式に正規化したキー文字列。
    """
    raw = f"{hwnd}_{wb_name}_{sh_name}"
    return re.sub(r'[\\/:*?"<>|]', "_", raw)


def _get_sheet(book: Any, sheet_id: str) -> Any:
    """
    指定ブックから sheet_id に対応するシート、またはアクティブシートを返す。

    Args:
        book: xlwings の Book オブジェクト。
        sheet_id: シートの GUID（空の場合はアクティブシートを使用）。

    Returns:
        シートオブジェクト。取得できない場合は None。
    """
    if sheet_id and xlc:
        sh = xlc.find_sheet_by_guid(book, sheet_id)
        if sh is not None:
            return sh
    try:
        return book.sheets.active
    except Exception:
        return None


def _show_undo_failed_dialog(hwnd: int, sheet_id: str, message: str) -> None:
    """
    復元できない場合に ui_qt.ui_undo の専用ダイアログでメッセージを表示する。
    UI サーバへ IPC でリクエストを投入し、結果ファイルが書き込まれるまでポーリングして待つ。

    Args:
        hwnd: Excel ウィンドウの HWND（ダイアログの親・中央配置用）。
        sheet_id: シート識別子（リクエスト識別用）。
        message: ユーザーに表示するエラーメッセージ（detail_text として渡す）。
    """
    if not message or not int(hwnd or 0):
        return
    if ipc_file is None or ensure_ui_server is None:
        return
    try:
        ensure_ui_server()
    except Exception as e:
        logger.warning("[UNDO] ensure_ui_server 失敗: %s", e)
        return
    req_dir = ipc_file.get_request_dir()
    res_dir = Path(ipc_file.get_ipc_root()) / "results"
    res_dir.mkdir(parents=True, exist_ok=True)
    ts_ms = int(time.time() * 1000)
    result_path = res_dir / f"res_undo_failed_{ts_ms}_{os.getpid()}.pkl"
    sheet_id_for_req = (sheet_id or "").strip() or "undo"
    # UI サーバが create_dialog で undo_failed を処理し、config/ui_undo.json の UNDO_FAILED で表示
    ufr: dict[str, Any] = {
        "action": "undo_failed",
        "detail_text": message,
        "items": [],
    }
    er_f = _get_window_rect(int(hwnd or 0))
    if er_f is not None:
        ufr["excel_rect"] = list(er_f)
    payload = {
        "parent_hwnd": int(hwnd),
        "result_path": str(result_path),
        "ready_path": "",
        "sheet_id": sheet_id_for_req,
        "action": "undo_failed",
        "module": "ui_qt.ui_undo",
        "req_dict": ufr,
    }
    req_path = req_dir / f"req_undo_failed_{ts_ms}_{os.getpid()}.pkl"
    try:
        ipc_file.write_pickle(req_path, payload)
    except Exception as e:
        logger.warning("[UNDO] 完了ダイアログ リクエスト失敗: %s", e)
        return
    # ダイアログが閉じられるまで結果ファイルの出現をポーリング（最大 60 秒）
    t0 = time.time()
    while (time.time() - t0) < 60.0:
        if result_path.exists() and result_path.stat().st_size > 0:
            return
        time.sleep(0.05)


def _show_undo_done_dialog(hwnd: int, sheet_id: str, sheet_name: str = "") -> None:
    """
    Undo 復元成功時に終了通知を ui_qt.ui_undo の専用ダイアログで表示する。
    表示定義は config/ui_undo.json の SCREENS.UNDO_DONE。他モジュールと同様に Excel 中央配置・前面表示を行う。

    Args:
        hwnd: Excel ウィンドウの HWND（ダイアログの親・中央配置用）。
        sheet_id: シート識別子（リクエスト識別用）。
        sheet_name: 復元したシート名（本文メッセージに含める。空の場合は「元に戻しました。」のみ）。
    """
    if not int(hwnd or 0):
        return
    if ipc_file is None or ensure_ui_server is None:
        return
    try:
        ensure_ui_server()
    except Exception as e:
        logger.warning("[UNDO] ensure_ui_server 失敗: %s", e)
        return
    req_dir = ipc_file.get_request_dir()
    res_dir = Path(ipc_file.get_ipc_root()) / "results"
    res_dir.mkdir(parents=True, exist_ok=True)
    ts_ms = int(time.time() * 1000)
    result_path = res_dir / f"res_undo_done_{ts_ms}_{os.getpid()}.pkl"
    sheet_id_for_req = (sheet_id or "").strip() or "undo"
    detail = f"シート「{sheet_name}」を元に戻しました。" if sheet_name else "元に戻しました。"
    # UI サーバが create_dialog で undo_done を処理し、config/ui_undo.json の UNDO_DONE で表示
    udr: dict[str, Any] = {
        "action": "undo_done",
        "detail_text": detail,
        "items": [],
    }
    er_u = _get_window_rect(int(hwnd or 0))
    if er_u is not None:
        udr["excel_rect"] = list(er_u)
    payload = {
        "parent_hwnd": int(hwnd),
        "result_path": str(result_path),
        "ready_path": "",
        "sheet_id": sheet_id_for_req,
        "action": "undo_done",
        "module": "ui_qt.ui_undo",
        "req_dict": udr,
    }
    req_path = req_dir / f"req_undo_done_{ts_ms}_{os.getpid()}.pkl"
    try:
        ipc_file.write_pickle(req_path, payload)
    except Exception as e:
        logger.warning("[UNDO] 復元完了ダイアログ リクエスト失敗: %s", e)
        return
    # ダイアログが閉じられるまで結果ファイルの出現をポーリング（最大 60 秒）
    t0 = time.time()
    while (time.time() - t0) < 60.0:
        if result_path.exists() and result_path.stat().st_size > 0:
            return
        time.sleep(0.05)


def save_undo_snapshot(
    book: Any,
    sheet_id: str = "",
    target_hwnd: Optional[int] = None,
    excel_hwnd: Optional[int] = None,
    **kwargs: Any,
) -> bool:
    """
    チェックOK時などに、元に戻す用にシートの有効データを保存する。
    保存先は core_sys.CacheManager（Windows Tmp + core_cst.CACHE_FILE_NAME）。
    キーは _make_undo_key(hwnd, ブック名, シート名) で一意。exec_undo で同一キーから復元する。

    Args:
        book: xlwings の Book オブジェクト。
        sheet_id: シートの GUID（空の場合はアクティブシート）。
        target_hwnd: Excel ウィンドウの HWND（キー生成用。省略時は excel_hwnd を使用）。
        excel_hwnd: 上記の代替。どちらか一方を渡す。
        **kwargs: 将来の拡張用（未使用）。

    Returns:
        True: 保存成功。False: シート取得失敗・キャッシュ不可・使用範囲読取失敗・save 失敗など。
    """
    hwnd = int(target_hwnd or excel_hwnd or 0)
    logger.info("[UNDO] スナップショット保存 開始 book=%s sheet_id=%s", getattr(book, "name", None), sheet_id)
    if book is None:
        logger.warning("[UNDO] スナップショット保存: 対象ブックなし")
        return False
    ptr_s = _get_sheet(book, sheet_id)
    if ptr_s is None:
        logger.warning("[UNDO] スナップショット保存: 対象シートなし sheet_id=%s", sheet_id)
        return False
    if hsys is None:
        logger.warning("[UNDO] スナップショット保存: CacheManager 利用不可")
        return False

    wb_name = getattr(book, "name", "") or ""
    sh_name = getattr(ptr_s, "name", "") or ""
    str_undo_key = _make_undo_key(hwnd, wb_name, sh_name)

    # UsedRange を (1,1) 起点で 2 次元リストに読み取り。単一セル・1行・複数行いずれも正規化
    list_2d: List[List[Any]] = []
    try:
        ur = getattr(ptr_s, "used_range", None)
        if ur is None:
            list_2d = []  # 使用範囲が無い場合は空を保存（復元時にクリア可能）
        else:
            nr = getattr(ur, "rows", None)
            nc = getattr(ur, "columns", None)
            if nr is None or nc is None:
                list_2d = []
            else:
                last_row = int(nr.count)
                ncols = int(nc.count)
                if last_row < 1 or ncols < 1:
                    list_2d = []
                else:
                    raw = ptr_s.range((1, 1), (last_row, ncols)).value
                    if raw is None:
                        list_2d = []
                    elif isinstance(raw, (list, tuple)):
                        # xlwings は 1 行のとき 1 次元で返す場合があるため行リストに正規化
                        if len(raw) > 0 and not isinstance(raw[0], (list, tuple)):
                            list_2d = [list(raw)]
                        else:
                            list_2d = [list(row) if isinstance(row, (list, tuple)) else [row] for row in raw]
                    else:
                        list_2d = [[raw]]
    except Exception as e:
        logger.warning("[UNDO] スナップショット保存: 使用範囲読込失敗: %s", e)
        return False

    # 枠固定状態を取得（復元時に再現するため）
    freeze_split_row, freeze_split_col = 0, 0
    try:
        app_api = getattr(getattr(book, "app", None), "api", None)
        if app_api is not None:
            aw = getattr(app_api, "ActiveWindow", None)
            if aw is not None and getattr(aw, "FreezePanes", False):
                freeze_split_row = int(getattr(aw, "SplitRow", 0) or 0)
                freeze_split_col = int(getattr(aw, "SplitColumn", 0) or 0)
    except Exception:
        pass

    # 復元時に write_chunk と枠固定再設定で使う payload
    payload = {
        "data": list_2d,
        "book_name": wb_name,
        "sheet_name": sh_name,
        "freeze_split_row": freeze_split_row,
        "freeze_split_col": freeze_split_col,
    }
    try:
        hsys.CacheManager.save(str_undo_key, payload)
        logger.info("[UNDO] スナップショット保存 完了 key=%s 行数=%s", str_undo_key, len(list_2d))
        return True
    except Exception as e:
        logger.warning("[UNDO] スナップショット保存: CacheManager.save 失敗: %s", e)
        return False


def exec_undo(
    book: Any,
    sheet_id: str = "",
    target_hwnd: Optional[int] = None,
    excel_hwnd: Optional[int] = None,
    **kwargs: Any,
) -> None:
    """
    キャッシュから直前の状態を復元する。
    save_undo_snapshot で保存した payload を _make_undo_key 同一キーで読み、シートに書き戻す。
    復元後は当該キーを CacheManager.delete で削除し、成功時は _show_undo_done_dialog で通知する。

    Args:
        book: xlwings の Book オブジェクト。
        sheet_id: シートの GUID（空の場合はアクティブシート）。
        target_hwnd: Excel ウィンドウの HWND（キー生成・ダイアログ親用）。
        excel_hwnd: 上記の代替。
        **kwargs: 将来の拡張用（未使用）。

    共通仕様:
        core_stat.set_status_info でステータス報告。失敗時は _show_undo_failed_dialog でモーダル表示。
    """
    t_flow = time.perf_counter()
    _log_undo_perf("enter", t_flow)
    _undo_trace("enter", t_flow)

    logger.info("[UNDO] 復元 開始 book=%s sheet_id=%s", getattr(book, "name", None), sheet_id)
    hwnd = int(target_hwnd or excel_hwnd or 0)
    if book is None:
        logger.warning("[UNDO] 対象ブックなし")
        _log_undo_perf("abort_no_book", t_flow)
        _undo_trace("abort_no_book", t_flow)
        _log_undo_perf("flow_end", t_flow)
        _undo_trace("flow_end", t_flow)
        return
    ptr_s = _get_sheet(book, sheet_id)
    if ptr_s is None:
        msg = "ERROR: シートを特定できませんでした。"
        if core_stat:
            try:
                core_stat.set_status_info(book.sheets.active, msg)
            except Exception:
                pass
        logger.warning("[UNDO] 対象シートなし sheet_id=%s", sheet_id)
        _show_undo_failed_dialog(hwnd, sheet_id, msg)
        _log_undo_perf("abort_no_sheet", t_flow)
        _undo_trace("abort_no_sheet", t_flow)
        _log_undo_perf("flow_end", t_flow)
        _undo_trace("flow_end", t_flow)
        return

    ptr_a = getattr(book, "app", None)
    if ptr_a is None:
        msg = "ERROR: アプリケーションを取得できませんでした。"
        logger.warning("[UNDO] book.app なし")
        _show_undo_failed_dialog(hwnd, sheet_id, msg)
        _log_undo_perf("abort_no_app", t_flow)
        _undo_trace("abort_no_app", t_flow)
        _log_undo_perf("flow_end", t_flow)
        _undo_trace("flow_end", t_flow)
        return

    wb_name = getattr(book, "name", "") or ""
    sh_name = getattr(ptr_s, "name", "") or ""
    str_undo_key = _make_undo_key(hwnd, wb_name, sh_name)

    if hsys is None:
        msg = "ERROR: キャッシュモジュールを利用できません。"
        if core_stat:
            core_stat.set_status_info(ptr_s, msg)
        _show_undo_failed_dialog(hwnd, sheet_id, msg)
        _log_undo_perf("abort_no_hsys", t_flow)
        _undo_trace("abort_no_hsys", t_flow)
        _log_undo_perf("flow_end", t_flow)
        _undo_trace("flow_end", t_flow)
        return

    dict_undo_payload = hsys.CacheManager.load(str_undo_key)
    if dict_undo_payload is None:
        msg = "元に戻すための物理キャッシュ情報が見つかりませんでした。"
        if core_stat:
            core_stat.set_status_info(ptr_s, msg)
        logger.info("[UNDO] キャッシュなし")
        _show_undo_failed_dialog(hwnd, sheet_id, msg)
        _log_undo_perf("abort_no_cache", t_flow)
        _undo_trace("abort_no_cache", t_flow)
        _log_undo_perf("flow_end", t_flow)
        _undo_trace("flow_end", t_flow)
        return

    _log_undo_perf("after_payload_load", t_flow, key=str_undo_key)
    _undo_trace("after_payload_load", t_flow, key=str_undo_key)

    # 構造復元（ヘッダ行削除等）とデータ復元の分岐。num_rows があれば構造復元として hc_hd_rs に委譲
    is_structure_undo = "num_rows" in dict_undo_payload

    prog_path_undo: Path | None = None
    # 構造／データ復元で共有する進捗 seq（キャッシュ削除フェーズまで同一リストを使う）
    undo_seq: list[int] = [0]
    try:
        if is_structure_undo:
            # 共通仕様: hc_* に依存しない。構造復元は別モジュールがあれば委譲、なければエラー表示。
            _log_undo_perf("branch_structure", t_flow)
            _undo_trace("branch_structure", t_flow)
            if ipc_file is not None and ensure_ui_server is not None:
                prog_path_undo = _undo_progress_file_path()
                _undo_progress_write(
                    prog_path_undo,
                    {
                        "status": "RUN",
                        "phase": _undo_msg("PHASE_UNDO_BOOT", "復元を開始しています..."),
                        "msg": _undo_msg("PHASE_UNDO_BOOT", "復元を開始しています..."),
                        "pct": 0,
                        "done": 0,
                        "total": 100,
                        "seq": 0,
                    },
                )
                _submit_undo_progress_ui(hwnd, sheet_id, prog_path_undo, 4)
            api_st = getattr(ptr_a, "api", None) or ptr_a
            try:
                try:
                    api_st.Interactive = False
                except Exception:
                    pass
                try:
                    from svc import hc_hd_rs  # noqa: F401

                    if prog_path_undo is not None:
                        _undo_progress_phase(
                            prog_path_undo,
                            undo_seq,
                            30,
                            _undo_msg("PHASE_UNDO_STRUCTURE", "構造を復元しています..."),
                        )
                    hc_hd_rs.restore_header_logic(target_hwnd=hwnd)
                    if core_stat:
                        core_stat.set_status_info(ptr_s, "UI")
                    _log_undo_perf("after_structure_restore", t_flow)
                    _undo_trace("after_structure_restore", t_flow)
                    if prog_path_undo is not None:
                        _undo_progress_phase(
                            prog_path_undo,
                            undo_seq,
                            88,
                            _undo_msg("PHASE_UNDO_STRUCTURE_FINISH", "復元後の処理を行っています..."),
                        )
                except ImportError:
                    _undo_progress_done(prog_path_undo)
                    msg = "ERROR: 構造復元モジュールが利用できません。"
                    logger.warning("[UNDO] 構造復元モジュールなし")
                    if core_stat:
                        core_stat.set_status_info(ptr_s, msg)
                    _show_undo_failed_dialog(hwnd, sheet_id, msg)
                    _log_undo_perf("abort_structure_import", t_flow)
                    _undo_trace("abort_structure_import", t_flow)
                    return
                except Exception as e:
                    _undo_progress_done(prog_path_undo)
                    msg = f"ERROR: ヘッダ復元失敗 Detail: {e}"
                    logger.exception("[UNDO] ヘッダ復元失敗: %s", e)
                    if core_stat:
                        core_stat.set_status_info(ptr_s, msg)
                    _show_undo_failed_dialog(hwnd, sheet_id, msg)
                    _log_undo_perf("abort_structure_error", t_flow)
                    _undo_trace("abort_structure_error", t_flow)
                    return
            finally:
                try:
                    api_st.Interactive = True
                except Exception:
                    pass
        else:
            # データ復元: payload.data を UsedRange に書き戻し、枠固定・A1 選択まで実施
            _log_undo_perf("branch_data", t_flow)
            _undo_trace("branch_data", t_flow)
            list_data = dict_undo_payload.get("data")
            if list_data is None:
                msg = "復元するデータがありません。"
                _show_undo_failed_dialog(hwnd, sheet_id, msg)
                _log_undo_perf("abort_empty_data", t_flow)
                _undo_trace("abort_empty_data", t_flow)
                return
            if not isinstance(list_data, list):
                msg = "ERROR: 復元データの形式が不正です。"
                _show_undo_failed_dialog(hwnd, sheet_id, msg)
                _log_undo_perf("abort_bad_data_type", t_flow)
                _undo_trace("abort_bad_data_type", t_flow)
                return

            if ipc_file is not None and ensure_ui_server is not None:
                prog_path_undo = _undo_progress_file_path()
                _undo_progress_write(
                    prog_path_undo,
                    {
                        "status": "RUN",
                        "phase": _undo_msg("PHASE_UNDO_BOOT", "復元を開始しています..."),
                        "msg": _undo_msg("PHASE_UNDO_BOOT", "復元を開始しています..."),
                        "pct": 0,
                        "done": 0,
                        "total": 100,
                        "seq": 0,
                    },
                )
                _submit_undo_progress_ui(hwnd, sheet_id, prog_path_undo, 9)
            # シート更新・再描画を抑止。with 内で ClearContents → write_chunk → 整形 → 枠固定 → A1 選択
            api = getattr(ptr_a, "api", None) or ptr_a
            try:
                api.Interactive = False
            except Exception:
                pass
            prev_calc = None
            prev_events = None
            ctx = (xlc.suspend_sheet_updates(ptr_s) if xlc else nullcontext())
            with ctx:
                try:
                    try:
                        prev_calc = api.Calculation
                        api.Calculation = -4135  # xlCalculationManual
                    except Exception:
                        pass
                    try:
                        prev_events = api.EnableEvents
                        api.EnableEvents = False
                    except Exception:
                        pass
                except Exception:
                    pass

                if prog_path_undo is not None:
                    _undo_progress_phase(
                        prog_path_undo,
                        undo_seq,
                        10,
                        _undo_msg("PHASE_UNDO_PREP", "復元を準備しています..."),
                    )

                try:
                    # クリア前に UsedRange の範囲を取得（復元後に有効領域を拡大させないため）
                    old_rows, old_cols = 0, 0
                    try:
                        ur = getattr(ptr_s, "used_range", None)
                        if ur is not None:
                            uapi = getattr(ur, "api", None)
                            if uapi is not None:
                                old_rows = int(getattr(getattr(uapi, "Rows", None), "Count", 0) or 0)
                                old_cols = int(getattr(getattr(uapi, "Columns", None), "Count", 0) or 0)
                    except Exception:
                        pass

                    if prog_path_undo is not None:
                        _undo_progress_phase(
                            prog_path_undo,
                            undo_seq,
                            20,
                            _undo_msg("PHASE_UNDO_MEASURE", "使用中の範囲を確認しています..."),
                        )

                    # 復元前に現状の UsedRange をクリア（ClearContents を優先して表示更新を抑える）
                    try:
                        ur = getattr(ptr_s, "used_range", None)
                        if ur is not None:
                            clear_fn = getattr(ur, "clearcontents", None) or getattr(ur, "ClearContents", None)
                            if callable(clear_fn):
                                clear_fn()
                            else:
                                try:
                                    ur.value = None
                                except Exception:
                                    pass
                    except Exception:
                        pass

                    if prog_path_undo is not None:
                        _undo_progress_phase(
                            prog_path_undo,
                            undo_seq,
                            35,
                            _undo_msg("PHASE_UNDO_CLEAR", "シートをクリアしています..."),
                        )

                    if prog_path_undo is not None:
                        _undo_progress_phase(
                            prog_path_undo,
                            undo_seq,
                            48,
                            _undo_msg("PHASE_UNDO_WRITE", "データを書き戻しています..."),
                        )

                    if xlc:
                        xlc.write_chunk(ptr_s, 1, 1, list_data)

                    saved_rows = len(list_data)
                    saved_cols = len(list_data[0]) if list_data else 0
                    # 空スナップショット（直前に使用範囲が無かった／空シート）: 行挿入後の UsedRange を丸ごと戻す
                    if saved_rows == 0 and xlc:
                        try:
                            xlc.clear_sheet_used_range(ptr_s)
                        except Exception:
                            pass

                    if prog_path_undo is not None:
                        _undo_progress_phase(
                            prog_path_undo,
                            undo_seq,
                            58,
                            _undo_msg("PHASE_UNDO_TRIM", "表示範囲を調整しています..."),
                        )

                    # 有効データ領域の拡大を防ぐ: clear_used_range_overflow で整形＋列幅リセット
                    if xlc and saved_rows > 0 and saved_cols > 0:
                        try:
                            xlc.clear_used_range_overflow(ptr_s, saved_rows, saved_cols)
                        except Exception:
                            pass
                    # 保存時より大きかった範囲の余白クリアと列幅を標準に戻す
                    if (old_rows > saved_rows or old_cols > saved_cols) and saved_rows > 0 and saved_cols > 0:
                        try:
                            if old_cols > saved_cols:
                                rng_right = ptr_s.range((1, saved_cols + 1), (max(old_rows, saved_rows), old_cols))
                                rng_right.value = None
                                try:
                                    sheet_api = getattr(ptr_s, "api", None)
                                    if sheet_api is not None:
                                        std_width = float(getattr(sheet_api, "StandardWidth", 8.43) or 8.43)
                                        for col in range(saved_cols + 1, old_cols + 1):
                                            try:
                                                col_rng = sheet_api.Columns(col)
                                                col_rng.ColumnWidth = std_width
                                            except Exception:
                                                pass
                                except Exception:
                                    pass
                            if old_rows > saved_rows:
                                rng_bottom = ptr_s.range((saved_rows + 1, 1), (old_rows, old_cols))
                                rng_bottom.value = None
                        except Exception:
                            pass

                    if prog_path_undo is not None:
                        _undo_progress_phase(
                            prog_path_undo,
                            undo_seq,
                            72,
                            _undo_msg("PHASE_UNDO_AUTOFIT", "列幅を整えています..."),
                        )
                    # オートフィット（行数が閾値以下だけ）
                    try:
                        max_af_rows = int(getattr(cst, "AUTOFIT_MAX_ROWS", 100000) or 100000)
                        if saved_rows <= max_af_rows:
                            ur = getattr(ptr_s, "used_range", None)
                            if ur is not None:
                                cols = getattr(ur, "columns", None)
                                if cols is not None:
                                    af = getattr(cols, "autofit", None) or getattr(cols, "AutoFit", None)
                                    if callable(af):
                                        af()
                    except Exception:
                        pass
                    if prog_path_undo is not None:
                        _undo_progress_phase(
                            prog_path_undo,
                            undo_seq,
                            84,
                            _undo_msg("PHASE_UNDO_FREEZE", "表示枠とスクロール位置を復元しています..."),
                        )
                    # 枠固定を復元
                    try:
                        ptr_s.activate()
                        app_api = getattr(ptr_a, "api", None) or ptr_a
                        aw = getattr(app_api, "ActiveWindow", None)
                        if aw is not None:
                            aw.FreezePanes = False
                            sr = int(dict_undo_payload.get("freeze_split_row") or 0)
                            sc = int(dict_undo_payload.get("freeze_split_col") or 0)
                            if sr > 0 or sc > 0:
                                aw.SplitRow = sr
                                aw.SplitColumn = sc
                                aw.FreezePanes = True
                    except Exception:
                        pass
                    # A1 に移動
                    try:
                        ptr_s.range("A1").select()
                    except Exception:
                        try:
                            rng = ptr_s.range("A1")
                            if getattr(rng, "api", None) is not None:
                                rng.api.Select()
                        except Exception:
                            pass
                    if prog_path_undo is not None:
                        _undo_progress_phase(
                            prog_path_undo,
                            undo_seq,
                            93,
                            _undo_msg("PHASE_UNDO_FINALIZE", "終了処理中..."),
                        )
                    if core_stat:
                        core_stat.set_status_info(ptr_s, "加工直前の物理状態へ正常に復元されました。")
                    _log_undo_perf(
                        "after_data_restore",
                        t_flow,
                        rows=saved_rows,
                        cols=saved_cols,
                    )
                    _undo_trace(
                        "after_data_restore",
                        t_flow,
                        rows=saved_rows,
                        cols=saved_cols,
                    )
                finally:
                    try:
                        api.Calculation = prev_calc if prev_calc is not None else -4105  # xlCalculationAutomatic
                    except Exception:
                        pass
                    try:
                        api.EnableEvents = True
                    except Exception:
                        pass
                    try:
                        api.Interactive = True
                    except Exception:
                        pass
                    # ScreenUpdating は suspend_sheet_updates の exit で True に戻る

        # 1 回だけ戻す: 進捗でキャッシュ削除フェーズ→削除→DONE で進捗を閉じてから終了通知
        if prog_path_undo is not None:
            _undo_progress_phase(
                prog_path_undo,
                undo_seq,
                96,
                _undo_msg("PHASE_UNDO_CACHE_DELETE", "復元情報を整理しています..."),
            )
        _log_undo_perf("before_cache_delete", t_flow)
        _undo_trace("before_cache_delete", t_flow)
        hsys.CacheManager.delete(str_undo_key)
        _log_undo_perf("after_cache_delete", t_flow)
        _undo_trace("after_cache_delete", t_flow)
        _undo_progress_done(prog_path_undo)
        logger.info("[UNDO] 復元 完了 シート=%s", sh_name)
        _show_undo_done_dialog(hwnd, sheet_id, sh_name)
        _log_undo_perf("after_done_ui", t_flow)
        _undo_trace("after_done_ui", t_flow)

    except Exception as ex_undo:
        _undo_progress_done(prog_path_undo)
        err_msg = f"ERROR: 元に戻す処理中に例外。 Detail: {ex_undo}"
        if core_stat:
            try:
                core_stat.set_status_info(ptr_s, err_msg)
            except Exception:
                pass
        logger.exception("[UNDO] 致命的エラー: %s", err_msg)
        _show_undo_failed_dialog(hwnd, sheet_id, err_msg)
        _log_undo_perf("abort_exception", t_flow)
        _undo_trace("abort_exception", t_flow)
        return

    finally:
        # 例外・正常どちらでも Excel の Interactive / ScreenUpdating を復帰し、ステータスバーを復元。
        # bring_to_front は常に実行（svc_dt_ymd 等と同様）。1.7.5 の成功時スキップは TOPMOST 進捗との相互作用の切り分けのため撤去（1.7.7）。
        try:
            api = getattr(ptr_a, "api", None) or ptr_a
            api.Interactive = True
            api.ScreenUpdating = True
        except Exception:
            pass
        if w32:
            try:
                w32.bring_to_front(hwnd)
            except Exception:
                pass
        if core_stat:
            try:
                info = core_stat.get_status_info(ptr_s)
                if info and hasattr(book, "app"):
                    app_api = getattr(book.app, "api", None)
                    if app_api is not None:
                        app_api.StatusBar = info
            except Exception:
                pass
        _log_undo_perf("flow_end", t_flow)
        _undo_trace("flow_end", t_flow)


# hc_main が undo_last_action で呼ぶ場合の互換エイリアス（後方互換）
undo_last_action = exec_undo
