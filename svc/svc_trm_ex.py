# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: svc/svc_trm_ex.py
Created: 2026-03-19
Version: 1.2.4
Purpose:
  文頭・文末トリム（選択範囲の空白削除）。選択範囲または有効データ領域を走査し、
  文頭/文末の件数表示後に「文頭削除」「文末削除」「全削除」のいずれかを実行。Undo 対応。
  画面は ui_qt.ui_trm_ex + config/ui_trm_ex.json。
History (latest 1):
  - 1.2.4 (2026-05-05) progress n/m の定義を変更。m=前後空白ありセル数（choice に応じた変換対象数）、n=変換中セル数。
  - 1.2.3 (2026-05-05) progress n/m をセル数ベースへ変更。更新頻度はセル進捗ステップ＋時間間引きで制御。
  - 1.2.2 (2026-05-05) 選択範囲を Areas 対応（歯抜け選択対応）。進捗 done/total を常時送信し n/m 表示を 0/0 にならないよう修正。
  - 1.2.1 (2026-05-05) トリム実行時に progress を表示。choice 確定後に進捗開始し、DONE 書込＋クローズ ACK 待ち後に完了通知を表示。
  - 1.2.0 (2026-04-13) 選択ダイアログ表示中、文頭・文末検出セルを sidecar＋viewport 追従で着色（ui_dupli と同系）。閉じる／確定時に全矩形クリア。HIGHLIGHT_LEADING / HIGHLIGHT_TRAILING。
  - 1.1.0 HC_LOG_PERF: [TRM_EX_PERF]。診断: [TRM_EX_TRACE]。
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path
from typing import Any, Optional

_path_svc = os.path.dirname(os.path.abspath(__file__))
_root = os.path.dirname(_path_svc)
if _root not in sys.path:
    sys.path.insert(0, _root)

from core.core_log import get_diag_logger, get_logger, get_perf_logger  # noqa: E402
from ui_qt.ipc_file import get_ipc_root, get_request_dir, read_pickle, write_pickle  # noqa: E402

_TRM_HL_SIDECAR_GLOB = "trm_ex_hl_rects_*.pkl"

logger = get_logger(__name__)
_trm_diag = get_diag_logger("hc_csv_tool.diag.trm_ex")
_perf = get_perf_logger("svc.svc_trm_ex.perf")
__version__ = "1.2.4"

_PROGRESS_CLOSE_ACK_TIMEOUT_SEC = 3.0
_PROGRESS_CLOSE_ACK_POLL_SEC = 0.03


def _elapsed_ms(since: float) -> int:
    return max(0, int((time.perf_counter() - since) * 1000))


def _trm_trace(phase: str, t0: float, **kv: object) -> None:
    try:
        if kv:
            _trm_diag.info(
                "[TRM_EX_TRACE] phase=%s cumulative_ms=%d %s",
                phase,
                _elapsed_ms(t0),
                " ".join("%s=%s" % (k, v) for k, v in kv.items()),
            )
        else:
            _trm_diag.info("[TRM_EX_TRACE] phase=%s cumulative_ms=%d", phase, _elapsed_ms(t0))
    except Exception:
        pass


def _perf_trm(phase: str, t0: float, **kv: object) -> None:
    try:
        if kv:
            _perf.info(
                "[TRM_EX_PERF] phase=%s cumulative_ms=%d %s",
                phase,
                _elapsed_ms(t0),
                " ".join("%s=%s" % (k, v) for k, v in kv.items()),
            )
        else:
            _perf.info("[TRM_EX_PERF] phase=%s cumulative_ms=%d", phase, _elapsed_ms(t0))
    except Exception:
        pass

try:
    from core import core_cst as cst
except Exception:
    cst = None  # type: ignore

try:
    from core import core_xlc as core_xlc_mod
except Exception:
    core_xlc_mod = None  # type: ignore


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


def _status_bar_save(book: Any) -> str:
    """
    現在の Excel ステータスバー文言を退避する。
    処理後に _status_bar_restore で復元するために使用する。
    """
    try:
        return str(book.app.api.StatusBar or "")
    except Exception:
        return ""


def _status_bar_set(book: Any, msg: str) -> None:
    """
    Excel のステータスバーに指定メッセージを表示する。
    処理中・完了・エラー時のユーザーへのフィードバック用。
    """
    try:
        book.app.api.DisplayStatusBar = True
        book.app.api.StatusBar = str(msg)
    except Exception:
        pass


def _status_bar_restore(book: Any, saved: str) -> None:
    """
    ステータスバーを _status_bar_save で退避した文言に戻す。
    処理終了時（正常・異常問わず）に必ず呼ぶ。
    """
    try:
        book.app.api.StatusBar = saved
    except Exception:
        pass


def _cfg() -> dict[str, Any]:
    """
    文頭・文末トリム用の画面・メッセージ設定を config/ui_trm_ex.json から読み込む。
    読込失敗時は UiConfigLoadError が発生する（救済なし）。
    """
    if cst is None:
        return {}
    return cst.get_ui_config_from_file_required("trm_ex")


def _msg(cfg: dict[str, Any], key: str, **fmt: Any) -> str:
    """
    設定の MESSAGES からキーに対応する文言を取得し、任意でフォーマットする。
    **fmt で {key} プレースホルダを置換する（例: n_leading=5, n_trailing=3）。
    """
    m = (cfg.get("MESSAGES") or {}).get(key) or key
    try:
        return str(m).format(**fmt)
    except Exception:
        return str(m)


def _normalize_2d(raw: Any, yn: int, xn: int) -> list[list[Any]]:
    """
    xlwings の Range.value で得た値を、yn×xn の 2 次元リストに正規化する。
    単一セル・1 行・複数行のいずれでも、欠けている要素は None で埋める。
    """
    if yn <= 0 or xn <= 0:
        return []
    if yn == 1 and xn == 1:
        return [[raw]]
    if yn == 1:
        row = raw if isinstance(raw, list) else [raw]
        return [row[:xn] + [None] * max(0, xn - len(row))]
    out: list[list[Any]] = []
    if not isinstance(raw, list):
        return [[None] * xn for _ in range(yn)]
    for r in range(yn):
        row = raw[r] if r < len(raw) else None
        if row is None:
            out.append([None] * xn)
        elif isinstance(row, list):
            out.append((row + [None] * xn)[:xn])
        else:
            out.append([row] + [None] * (xn - 1) if xn > 1 else [row])
    return out


def _rgb_to_excel_bgr_int(r: int, g: int, b: int) -> int:
    """Excel / xlwings Range.color 用の BGR パック整数（dupli._highlight_bgr と同式）。"""
    return int(r) + int(g) * 256 + int(b) * 65536


def _trm_rgb_default(which: str) -> tuple[int, int, int]:
    if which == "LEADING":
        return 200, 240, 255  # うすい水色
    return 175, 215, 255  # うすい青（文末）


def _trm_highlight_bgr_from_cfg(cfg: dict[str, Any], which: str) -> int:
    key = "HIGHLIGHT_LEADING" if which == "LEADING" else "HIGHLIGHT_TRAILING"
    block = cfg.get(key) or {}
    rgb = block.get("RGB")
    if isinstance(rgb, (list, tuple)) and len(rgb) >= 3:
        r, g, b = int(rgb[0]), int(rgb[1]), int(rgb[2])
    else:
        r, g, b = _trm_rgb_default(which)
    return _rgb_to_excel_bgr_int(r, g, b)


def _sheet_name_for_highlight(ptr_s: Any) -> str:
    try:
        return str(ptr_s.name or "").strip()
    except Exception:
        return ""


def _book_name_for_highlight(ptr_s: Any) -> str:
    try:
        return str(ptr_s.book.name or "").strip()
    except Exception:
        return ""


def _cleanup_stale_trm_hl_sidecars(*, max_age_sec: float = 86400.0) -> None:
    d = Path(get_ipc_root()) / "progress"
    if not d.is_dir():
        return
    now = time.time()
    for p in d.glob(_TRM_HL_SIDECAR_GLOB):
        try:
            if now - p.stat().st_mtime > max_age_sec:
                p.unlink(missing_ok=True)
        except OSError:
            pass


def _write_trm_ex_hl_sidecar(
    sheet_id: str,
    leading: list[list[int]],
    trailing: list[list[int]],
    fill_bgr_leading: int,
    fill_bgr_trailing: int,
) -> Path:
    d = Path(get_ipc_root()) / "progress"
    d.mkdir(parents=True, exist_ok=True)
    sid = str(sheet_id or "").strip() or "_"
    ts_ms = int(time.time() * 1000)
    path = d / f"trm_ex_hl_rects_{sid}_{ts_ms}_{os.getpid()}.pkl"
    write_pickle(
        path,
        {
            "v": 2,
            "leading": leading,
            "trailing": trailing,
            "fill_bgr_leading": int(fill_bgr_leading),
            "fill_bgr_trailing": int(fill_bgr_trailing),
        },
    )
    return path


def _progress_path(sheet_id: str) -> Path:
    d = Path(get_ipc_root()) / "progress"
    d.mkdir(parents=True, exist_ok=True)
    return d / f"progress_trm_ex_{sheet_id}.pkl"


def _progress_closed_ack_path(sheet_id: str) -> Path:
    d = Path(get_ipc_root()) / "progress"
    d.mkdir(parents=True, exist_ok=True)
    return d / f"progress_trm_ex_closed_{sheet_id}.pkl"


def _wait_progress_closed_ack(path: Optional[Path], timeout_sec: float = _PROGRESS_CLOSE_ACK_TIMEOUT_SEC) -> None:
    if path is None:
        return
    p = path
    t0 = time.perf_counter()
    while True:
        try:
            if p.exists():
                return
        except Exception:
            return
        if (time.perf_counter() - t0) >= max(0.05, float(timeout_sec)):
            logger.info("[TRM_EX] progress close ack timeout: %s", str(p))
            return
        time.sleep(_PROGRESS_CLOSE_ACK_POLL_SEC)


def _progress_write(path: Path, obj: dict[str, Any]) -> None:
    try:
        d = dict(obj)
        if "seq" not in d:
            try:
                prev = read_pickle(path)
                seq = int(prev.get("seq", -1)) + 1 if isinstance(prev, dict) else 0
            except Exception:
                seq = 0
            d["seq"] = seq
        write_pickle(path, d)
    except Exception:
        pass


def _submit_progress_ui(
    parent_hwnd: int,
    sheet_id: str,
    progress_path: Path,
    phase_total: int,
    *,
    progress_closed_path: Path | None = None,
) -> None:
    try:
        from svc.svc_host import ensure_ui_server

        ensure_ui_server()
        excel_rect = _get_window_rect(int(parent_hwnd or 0))
        res_dir = Path(get_ipc_root()) / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        ts_ms = int(time.time() * 1000)
        result_path = str(res_dir / f"res_progress_trm_ex_{ts_ms}_{os.getpid()}.pkl")
        req_dict: dict[str, Any] = {
            "action": "progress",
            "progress_path": str(progress_path),
            "phase_total": int(phase_total),
            "excel_lock": False,
            "no_native_window": True,
        }
        if progress_closed_path is not None:
            cp = str(progress_closed_path).strip()
            if cp:
                req_dict["progress_closed_path"] = cp
        if excel_rect is not None:
            req_dict["excel_rect"] = list(excel_rect)
        payload: dict[str, Any] = {
            "parent_hwnd": int(parent_hwnd),
            "result_path": result_path,
            "ready_path": "",
            "sheet_id": str(sheet_id),
            "log_path": "",
            "action": "progress",
            "module": "ui_qt.ui_trm_ex",
            "req_dict": req_dict,
        }
        req_path = get_request_dir() / f"req_progress_trm_ex_{ts_ms}_{os.getpid()}.pkl"
        write_pickle(req_path, payload)
        logger.info("[TRM_EX] progress UI request: %s", req_path)
    except Exception as exc:
        logger.warning("[TRM_EX] progress UI request failed: %s", exc)


def _submit_choice_ui(
    parent_hwnd: int,
    sheet_id: str,
    result_path: Path,
    n_leading: int,
    n_trailing: int,
    cfg: dict[str, Any],
    highlight_trm: dict[str, Any] | None = None,
) -> None:
    """
    文頭・文末の件数（n_leading, n_trailing）を表示し、4 ボタンで選択させる
    モーダルダイアログの表示を ui_server に依頼する。
    ユーザーの選択結果は result_path に Pickle で書き出されるため、呼び出し元でポーリングする。
    """
    try:
        from svc.svc_host import ensure_ui_server

        ensure_ui_server()
        res_dir = Path(get_ipc_root()) / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        req_dict: dict[str, Any] = {
            "action": "trm_ex_choice",
            "n_leading": int(n_leading),
            "n_trailing": int(n_trailing),
            "modeless": False,
        }
        if highlight_trm:
            req_dict["highlight_trm"] = dict(highlight_trm)
        er_c = _get_window_rect(int(parent_hwnd or 0))
        if er_c is not None:
            req_dict["excel_rect"] = list(er_c)
        payload: dict[str, Any] = {
            "parent_hwnd": int(parent_hwnd),
            "result_path": str(result_path),
            "ready_path": "",
            "sheet_id": str(sheet_id),
            "log_path": "",
            "action": "trm_ex",
            "module": "ui_qt.ui_trm_ex",
            "req_dict": req_dict,
        }
        ts_ms = int(time.time() * 1000)
        req_path = get_request_dir() / f"req_trm_ex_choice_{ts_ms}_{os.getpid()}.pkl"
        write_pickle(req_path, payload)
        logger.info("[TRM_EX] choice UI request: %s", req_path)
    except Exception as exc:
        logger.warning("[TRM_EX] choice UI request failed: %s", exc)


def _submit_done_ui(
    parent_hwnd: int,
    sheet_id: str,
    message: str,
    title: str = "文頭・文末トリム",
    cfg: dict[str, Any] | None = None,
) -> None:
    """
    トリム実行後の完了通知をモーダルで表示するため ui_server に依頼する。
    SCREENS.DONE の設定に従い、アイコン・中央表示・OK ボタンで閉じる。
    message には STATUS_DONE をフォーマットした縦並び文言（\\n 含む）を渡す。
    """
    try:
        from svc.svc_host import ensure_ui_server

        ensure_ui_server()
        ts_ms = int(time.time() * 1000)
        res_dir = Path(get_ipc_root()) / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        result_path = str(res_dir / f"res_trm_ex_done_{ts_ms}_{os.getpid()}.pkl")
        done_cfg = ((cfg or {}).get("SCREENS") or {}).get("DONE") or {}
        req_dict: dict[str, Any] = {
            "action": "trm_ex_done",
            "modeless": False,
            "title": str(title),
            "message": str(message),
        }
        er_d = _get_window_rect(int(parent_hwnd or 0))
        if er_d is not None:
            req_dict["excel_rect"] = list(er_d)
        payload = {
            "parent_hwnd": int(parent_hwnd),
            "result_path": result_path,
            "ready_path": "",
            "sheet_id": str(sheet_id),
            "log_path": "",
            "action": "trm_ex",
            "module": "ui_qt.ui_trm_ex",
            "req_dict": req_dict,
        }
        req_path = get_request_dir() / f"req_trm_ex_done_{ts_ms}_{os.getpid()}.pkl"
        write_pickle(req_path, payload)
        logger.info("[TRM_EX] done UI request: %s", req_path)
    except Exception as exc:
        logger.warning("[TRM_EX] done UI request failed: %s", exc)


def _submit_no_target_ui(
    parent_hwnd: int,
    sheet_id: str,
    message: str,
    cfg: dict[str, Any],
) -> None:
    """
    削除対象が一件もない場合、または有効データ領域が空の場合の通知を
    モーダルで表示するため ui_server に依頼する。SCREENS.NO_TARGET の設定を使用する。
    """
    try:
        from svc.svc_host import ensure_ui_server

        ensure_ui_server()
        ts_ms = int(time.time() * 1000)
        res_dir = Path(get_ipc_root()) / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        result_path = str(res_dir / f"res_trm_ex_notarget_{ts_ms}_{os.getpid()}.pkl")
        no_cfg = (cfg.get("SCREENS") or {}).get("NO_TARGET") or {}
        req_dict: dict[str, Any] = {
            "action": "trm_ex_no_target",
            "modeless": False,
            "title": str(no_cfg.get("TITLE") or "文頭・文末トリム"),
            "message": str(message),
        }
        er_n = _get_window_rect(int(parent_hwnd or 0))
        if er_n is not None:
            req_dict["excel_rect"] = list(er_n)
        payload = {
            "parent_hwnd": int(parent_hwnd),
            "result_path": result_path,
            "ready_path": "",
            "sheet_id": str(sheet_id),
            "log_path": "",
            "action": "trm_ex",
            "module": "ui_qt.ui_trm_ex",
            "req_dict": req_dict,
        }
        req_path = get_request_dir() / f"req_trm_ex_notarget_{ts_ms}_{os.getpid()}.pkl"
        write_pickle(req_path, payload)
        logger.info("[TRM_EX] no_target UI request: %s", req_path)
    except Exception as exc:
        logger.warning("[TRM_EX] no_target UI request failed: %s", exc)


def _poll_result(result_path: Path, timeout_sec: float = 120.0) -> dict[str, Any] | None:
    """
    選択ダイアログの結果が result_path に書き出されるまでポーリングする。
    取得した辞書には "choice"（leading / trailing / all / cancel）が含まれる。
    タイムアウト時は None を返す。
    """
    t0 = time.time()
    while (time.time() - t0) < timeout_sec:
        if result_path.exists() and result_path.stat().st_size > 0:
            try:
                return read_pickle(result_path)
            except Exception:
                pass
        time.sleep(0.05)
    return None


def trim_cells(target_hwnd: Optional[int] = None, sheet_id: str = "") -> None:
    """
    選択範囲（または選択なし時は有効データ領域）の文頭・文末空白を走査し、
    選択ダイアログで「文頭削除」「文末削除」「全削除」のいずれかを実行する。Undo 対応。

    手順:
      1. Excel コンテキスト取得。選択ありならその範囲、なければ UsedRange を対象とする。
      2. 対象範囲を読込み、文字列セルの文頭・文末の空白件数を集計。
      3. 件数が 0 なら NO_TARGET 通知で終了。
      4. 選択 UI を表示し、ユーザーの choice（leading / trailing / all / cancel）を取得。
      5. cancel なら終了。それ以外なら配列にトリムを適用し、Undo スナップショット保存後に書込。
      6. 完了通知（削除文頭数・削除文末数を縦並び表示）を表示し、ステータスバー復元。
    """
    t_flow = time.perf_counter()
    _perf_trm("enter", t_flow)
    _trm_trace("enter", t_flow)

    if core_xlc_mod is None:
        logger.error("[TRM_EX] core_xlc not available")
        _perf_trm("abort_no_core_xlc", t_flow)
        _trm_trace("abort_no_core_xlc", t_flow)
        return
    ctx = core_xlc_mod.get_excel_context_from_hwnd(int(target_hwnd or 0), sheet_id)
    if ctx is None:
        logger.error("[TRM_EX] Excel context not available (xlwings + HWND)")
        _perf_trm("abort_no_context", t_flow)
        _trm_trace("abort_no_context", t_flow)
        return

    ptr_a, ptr_w, ptr_s, ph = ctx
    sid = str(sheet_id or "").strip() or f"trm_ex_{abs(id(ptr_s))}"
    progress_path = _progress_path(sid)
    progress_closed_path = _progress_closed_ack_path(sid)
    progress_close_waited = False
    logger.info("[TRM_EX] 開始")
    _perf_trm("after_context", t_flow, hwnd=ph)
    _trm_trace("after_context", t_flow, hwnd=ph)
    try:
        _cleanup_stale_trm_hl_sidecars(max_age_sec=3600.0)
    except Exception:
        pass
    cfg = _cfg()
    saved_status = _status_bar_save(ptr_w)

    try:
        # 対象範囲の決定: 選択ありなら Areas（歯抜け選択対応）、なければ used_range(1 area)
        ptr_sel = None
        try:
            ptr_sel = ptr_a.selection
        except Exception:
            pass

        areas: list[dict[str, Any]] = []
        range_src = "used_range"
        total_rows = 0
        total_cols = 0

        # 選択範囲が有効なら Areas へ展開（連続/歯抜けの両対応）
        if ptr_sel is not None:
            try:
                api_areas = ptr_sel.api.Areas
                area_count = int(api_areas.Count)
                for i in range(1, area_count + 1):
                    ar = api_areas.Item(i)
                    yn_i = int(ar.Rows.Count)
                    xn_i = int(ar.Columns.Count)
                    if yn_i < 1 or xn_i < 1:
                        continue
                    y1_i = int(ar.Row)
                    x1_i = int(ar.Column)
                    areas.append({"y1": y1_i, "x1": x1_i, "yn": yn_i, "xn": xn_i})
                    total_rows += yn_i
                    total_cols += xn_i
                if areas:
                    range_src = "selection_areas"
                else:
                    ptr_sel = None
            except Exception:
                # API Areas が取得できない環境向けフォールバック（単一矩形）
                try:
                    val_yn = int(ptr_sel.rows.count)
                    val_xn = int(ptr_sel.columns.count)
                    if val_yn >= 1 and val_xn >= 1:
                        val_y1 = int(ptr_sel.row)
                        val_x1 = int(ptr_sel.column)
                        areas.append({"y1": val_y1, "x1": val_x1, "yn": val_yn, "xn": val_xn})
                        total_rows += val_yn
                        total_cols += val_xn
                        range_src = "selection"
                    else:
                        ptr_sel = None
                except Exception:
                    ptr_sel = None

        # 選択なし or 無効な選択 → UsedRange を 1 area として扱う
        if not areas:
            try:
                ur = ptr_s.used_range
                val_yn = int(ur.rows.count)
                val_xn = int(ur.columns.count)
                if val_yn < 1 or val_xn < 1:
                    _status_bar_set(ptr_w, _msg(cfg, "NO_DATA"))
                    no_cfg = (cfg.get("SCREENS") or {}).get("NO_TARGET") or {}
                    _submit_no_target_ui(ph, sheet_id, _msg(cfg, "NO_DATA"), cfg)
                    _perf_trm("early_no_data", t_flow)
                    _trm_trace("early_no_data", t_flow)
                    return
                val_y1 = int(ur.row)
                val_x1 = int(ur.column)
                areas = [{"y1": val_y1, "x1": val_x1, "yn": val_yn, "xn": val_xn}]
                total_rows = val_yn
                total_cols = val_xn
                range_src = "used_range"
            except Exception as e:
                logger.warning("[TRM_EX] used_range 取得失敗: %s", e)
                _status_bar_set(ptr_w, _msg(cfg, "ERROR_PREFIX"))
                _perf_trm("abort_used_range_failed", t_flow)
                _trm_trace("abort_used_range_failed", t_flow)
                return

        areas_total = max(1, int(len(areas)))
        _perf_trm(
            "after_range_resolve",
            t_flow,
            source=range_src,
            areas=areas_total,
            yn=total_rows,
            xn=total_cols,
        )
        _trm_trace(
            "after_range_resolve",
            t_flow,
            source=range_src,
            areas=areas_total,
            yn=total_rows,
            xn=total_cols,
        )

        # 各 area を読込み 2 次元化して保持
        area_payloads: list[dict[str, Any]] = []
        for a_idx, a in enumerate(areas, start=1):
            y1_i = int(a["y1"])
            x1_i = int(a["x1"])
            yn_i = int(a["yn"])
            xn_i = int(a["xn"])
            try:
                rng = ptr_s.range((y1_i, x1_i), (y1_i + yn_i - 1, x1_i + xn_i - 1))
                raw = rng.value
                arr_i = _normalize_2d(raw, yn_i, xn_i)
            except Exception as e:
                logger.warning("[TRM_EX] 読込失敗 area=%s: %s", a_idx, e)
                _status_bar_set(ptr_w, _msg(cfg, "ERROR_PREFIX"))
                _perf_trm("abort_matrix_read_failed", t_flow, area_i=a_idx, areas=areas_total)
                _trm_trace("abort_matrix_read_failed", t_flow, area_i=a_idx, areas=areas_total)
                return
            if not arr_i or len(arr_i) != yn_i:
                _status_bar_set(ptr_w, _msg(cfg, "ERROR_PREFIX"))
                _perf_trm("abort_matrix_invalid", t_flow, area_i=a_idx, areas=areas_total)
                _trm_trace("abort_matrix_invalid", t_flow, area_i=a_idx, areas=areas_total)
                return
            area_payloads.append({"y1": y1_i, "x1": x1_i, "yn": yn_i, "xn": xn_i, "arr": arr_i})

        _perf_trm("after_matrix_read", t_flow, areas=areas_total, yn=total_rows, xn=total_cols)
        _trm_trace("after_matrix_read", t_flow, areas=areas_total, yn=total_rows, xn=total_cols)
        # 文字列セルについて文頭・文末の件数と、着色用 1×1 矩形（シート座標）を集計
        n_leading = 0
        n_trailing = 0
        n_any_target = 0
        hl_leading: list[list[int]] = []
        hl_trailing: list[list[int]] = []
        for ap in area_payloads:
            y1_i = int(ap["y1"])
            x1_i = int(ap["x1"])
            arr_i = ap["arr"]
            for ri, row in enumerate(arr_i):
                for ci, cell in enumerate(row):
                    if not isinstance(cell, str):
                        continue
                    r_abs = y1_i + ri
                    c_abs = x1_i + ci
                    if cell != cell.lstrip():
                        n_leading += 1
                        hl_leading.append([r_abs, c_abs, r_abs, c_abs])
                    if cell != cell.rstrip():
                        n_trailing += 1
                        hl_trailing.append([r_abs, c_abs, r_abs, c_abs])
                    if (cell != cell.lstrip()) or (cell != cell.rstrip()):
                        n_any_target += 1

        logger.info(
            "[TRM_EX] 走査 areas=%s rows=%s cols=%s 文頭=%s 文末=%s",
            areas_total,
            total_rows,
            total_cols,
            n_leading,
            n_trailing,
        )
        _perf_trm(
            "after_scan_counts",
            t_flow,
            n_leading=n_leading,
            n_trailing=n_trailing,
            areas=areas_total,
            yn=total_rows,
            xn=total_cols,
        )
        _trm_trace(
            "after_scan_counts",
            t_flow,
            n_leading=n_leading,
            n_trailing=n_trailing,
            areas=areas_total,
            yn=total_rows,
            xn=total_cols,
        )

        if n_leading == 0 and n_trailing == 0:
            _status_bar_set(ptr_w, _msg(cfg, "NO_TARGET"))
            _submit_no_target_ui(ph, sheet_id, _msg(cfg, "NO_TARGET"), cfg)
            logger.info("[TRM_EX] 削除対象なし")
            _perf_trm("early_no_target", t_flow)
            _trm_trace("early_no_target", t_flow)
            return

        # 選択ダイアログを IPC で表示し、result_path に結果が書かれるまでポーリング
        res_dir = Path(get_ipc_root()) / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        ts_ms = int(time.time() * 1000)
        result_path = res_dir / f"res_trm_ex_choice_{ts_ms}_{os.getpid()}.pkl"
        bgr_l = _trm_highlight_bgr_from_cfg(cfg, "LEADING")
        bgr_t = _trm_highlight_bgr_from_cfg(cfg, "TRAILING")
        hl_path = _write_trm_ex_hl_sidecar(sid, hl_leading, hl_trailing, bgr_l, bgr_t)
        highlight_trm: dict[str, Any] = {
            "rects_path": str(hl_path),
            "book_name": _book_name_for_highlight(ptr_s),
            "sheet_name": _sheet_name_for_highlight(ptr_s),
            "viewport_follow": True,
        }
        _submit_choice_ui(ph, sid, result_path, n_leading, n_trailing, cfg, highlight_trm)
        _perf_trm("after_choice_ui_submit", t_flow)
        _trm_trace("after_choice_ui_submit", t_flow)

        res = _poll_result(result_path)
        choice = (res or {}).get("choice") if isinstance(res, dict) else None
        if choice == "cancel" or choice is None:
            _status_bar_restore(ptr_w, saved_status)
            logger.info("[TRM_EX] ユーザーキャンセル")
            _perf_trm("user_cancel", t_flow, choice=choice)
            _trm_trace("user_cancel", t_flow, choice=choice)
            return

        try:
            if progress_closed_path.exists():
                progress_closed_path.unlink(missing_ok=True)
        except Exception:
            pass
        _progress_write(
            progress_path,
            {
                "status": "RUN",
                "phase_i": 1,
                "phase_total": 2,
                "phase": _msg(cfg, "PHASE_APPLY"),
                "message": _msg(cfg, "PROGRESS_CUSTOM_APPLY"),
                "pct": 5,
                "done": 0,
                "total": 1,
                "seq": 0,
            },
        )
        _submit_progress_ui(
            ph,
            sid,
            progress_path,
            phase_total=2,
            progress_closed_path=progress_closed_path,
        )

        # ユーザー選択に応じて lstrip / rstrip / strip を適用し、完了通知用の件数をカウント
        n_leading_done = 0
        n_trailing_done = 0
        if choice == "leading":
            total_target_cells = max(1, int(n_leading))
        elif choice == "trailing":
            total_target_cells = max(1, int(n_trailing))
        else:
            total_target_cells = max(1, int(n_any_target))
        apply_done_cells = 0
        _progress_write(
            progress_path,
            {
                "status": "RUN",
                "phase_i": 1,
                "phase_total": 2,
                "phase": _msg(cfg, "PHASE_APPLY"),
                "message": _msg(cfg, "PROGRESS_CUSTOM_APPLY"),
                "pct": 5,
                "done": 0,
                "total": total_target_cells,
            },
        )
        apply_emit_step = max(1, total_target_cells // 200)
        apply_emit_interval_sec = 0.05
        apply_last_emit_done = -1
        apply_last_emit_at = time.perf_counter()
        for a_idx, ap in enumerate(area_payloads, start=1):
            arr_i = ap["arr"]
            for ri, row in enumerate(arr_i):
                for ci, cell in enumerate(row):
                    if not isinstance(cell, str):
                        continue
                    lead_hit = cell != cell.lstrip()
                    trail_hit = cell != cell.rstrip()
                    if choice == "leading":
                        new_val = cell.lstrip()
                        if new_val != cell:
                            n_leading_done += 1
                        arr_i[ri][ci] = new_val
                        target_hit = lead_hit
                    elif choice == "trailing":
                        target_hit = trail_hit
                        new_val = cell.rstrip()
                        if new_val != cell:
                            n_trailing_done += 1
                        arr_i[ri][ci] = new_val
                    else:  # all
                        target_hit = lead_hit or trail_hit
                        new_val = cell.strip()
                        if lead_hit:
                            n_leading_done += 1
                        if trail_hit:
                            n_trailing_done += 1
                        arr_i[ri][ci] = new_val
                    if not target_hit:
                        continue
                    apply_done_cells += 1
                    now = time.perf_counter()
                    need_emit = (
                        apply_done_cells >= total_target_cells
                        or (apply_done_cells - apply_last_emit_done) >= apply_emit_step
                        or (now - apply_last_emit_at) >= apply_emit_interval_sec
                    )
                    if need_emit:
                        _progress_write(
                            progress_path,
                            {
                                "status": "RUN",
                                "phase_i": 1,
                                "phase_total": 2,
                                "phase": _msg(cfg, "PHASE_APPLY"),
                                "message": _msg(cfg, "PROGRESS_CUSTOM_APPLY"),
                                "pct": min(70, max(5, int(5 + (65 * apply_done_cells / total_target_cells)))),
                                "done": apply_done_cells,
                                "total": total_target_cells,
                            },
                        )
                        apply_last_emit_done = apply_done_cells
                        apply_last_emit_at = now

        _progress_write(
            progress_path,
            {
                "status": "RUN",
                "phase_i": 2,
                "phase_total": 2,
                "phase": _msg(cfg, "PHASE_WRITE"),
                "message": _msg(cfg, "PROGRESS_CUSTOM_WRITE"),
                "pct": 75,
                "done": total_target_cells,
                "total": total_target_cells,
            },
        )

        _perf_trm("after_apply_choice", t_flow, choice=choice)
        _trm_trace("after_apply_choice", t_flow, choice=choice)

        # 書込前に Undo 用スナップショットを保存し、Interactive 停止のうえ write_chunk で反映
        try:
            from svc.svc_undo import save_undo_snapshot

            save_undo_snapshot(ptr_w, sheet_id=sheet_id, target_hwnd=ph, excel_hwnd=ph)
        except Exception as e:
            logger.warning("[TRM_EX] save_undo_snapshot failed (undo unavailable): %s", e)
        _perf_trm("after_undo_snapshot_attempt", t_flow)
        _trm_trace("after_undo_snapshot_attempt", t_flow)

        try:
            from core import core_xlc

            for ap in area_payloads:
                core_xlc.write_chunk(
                    ptr_s,
                    int(ap["y1"]),
                    int(ap["x1"]),
                    ap["arr"],
                )
                _progress_write(
                    progress_path,
                    {
                        "status": "RUN",
                        "phase_i": 2,
                        "phase_total": 2,
                        "phase": _msg(cfg, "PHASE_WRITE"),
                        "message": _msg(cfg, "PROGRESS_CUSTOM_WRITE"),
                        "pct": 99,
                        "done": total_target_cells,
                        "total": total_target_cells,
                    },
                )
        except Exception as e:
            logger.exception("[TRM_EX] 書込失敗: %s", e)
            _status_bar_set(ptr_w, f"{_msg(cfg, 'ERROR_PREFIX')}: {e}")
            _progress_write(
                progress_path,
                {
                    "status": "DONE",
                    "pct": 100,
                    "phase_i": 2,
                    "phase_total": 2,
                    "done": total_target_cells,
                    "total": total_target_cells,
                },
            )
            if not progress_close_waited:
                _wait_progress_closed_ack(progress_closed_path)
                progress_close_waited = True
            _perf_trm("abort_write_failed", t_flow)
            _trm_trace("abort_write_failed", t_flow)
            return

        _perf_trm("after_write_chunk", t_flow)
        _trm_trace("after_write_chunk", t_flow)

        logger.info(
            "[TRM_EX] 運用ログ トリム 種別=%s 削除文頭=%s 削除文末=%s",
            choice,
            n_leading_done,
            n_trailing_done,
        )
        done_msg = _msg(
            cfg,
            "STATUS_DONE",
            n_leading=n_leading_done,
            n_trailing=n_trailing_done,
        )
        _status_bar_set(ptr_w, done_msg.replace("\n", " "))
        _progress_write(
            progress_path,
            {
                "status": "DONE",
                "pct": 100,
                "phase_i": 2,
                "phase_total": 2,
                "done": total_target_cells,
                "total": total_target_cells,
            },
        )
        if not progress_close_waited:
            _wait_progress_closed_ack(progress_closed_path)
            progress_close_waited = True
        _submit_done_ui(ph, sid, done_msg, cfg.get("SCREENS", {}).get("DONE", {}).get("TITLE") or "文頭・文末トリム", cfg)
        _perf_trm("after_done_ui", t_flow)
        _trm_trace("after_done_ui", t_flow)

    except Exception as ex:
        logger.exception("[TRM_EX] %s", ex)
        try:
            _status_bar_set(ptr_w, f"{_msg(cfg, 'ERROR_PREFIX')}: {ex}")
        except Exception:
            pass
        _perf_trm("except_flow", t_flow)
        _trm_trace("except_flow", t_flow)
    finally:
        try:
            _status_bar_restore(ptr_w, saved_status)
        except Exception:
            pass
        try:
            from core import core_w32

            core_w32.bring_to_front(ph)
        except Exception:
            pass
        _perf_trm("flow_end", t_flow)
        _trm_trace("flow_end", t_flow)
