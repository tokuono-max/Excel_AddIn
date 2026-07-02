# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: svc/svc_trm_ex.py
Created: 2026-03-19
Version: 1.3.1
Purpose:
  文頭・文末トリム（選択範囲の空白削除）。選択範囲または有効データ領域を走査し、
  文頭/文末の件数表示後に「文頭削除」「文末削除」「全削除」のいずれかを実行。Undo 対応。
  画面は ui_qt.ui_trm_ex + config/ui_trm_ex.json。
History (latest 1):
  - 1.3.1 (2026-07-02) 走査進捗を選択 UI 直前で 100% DONE クローズ（背面表示の解消）。適用フェーズで進捗を再表示。
  - 1.3.0 (2026-07-02) 速度: UsedRange 交差・チャンク読込・対象セルのみ適用/差分書込。region スナップショット。走査段階から進捗表示。
  - 1.2.6 (2026-07-02) 進捗 pickle を verified 書込（DONE 検証、csv_ld 同型）。
  - 1.2.5 (2026-06-29) 進捗クローズ ACK を共通モジュールへ統一（15秒+nudge）。3秒タイムアウトの独自待機を廃止。
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
from core.progress_close_ack import (  # noqa: E402
    progress_closed_ack_path,
    reset_progress_closed_ack,
    wait_progress_closed_with_nudge,
)
from core.progress_pickle_write import dispatch_progress_write  # noqa: E402
from svc.dt_convert_helpers import (  # noqa: E402
    read_sheet_matrix,
    snapshot_region_from_areas,
    trim_areas_to_used_range,
    write_changed_slices,
)

_TRM_HL_SIDECAR_GLOB = "trm_ex_hl_rects_*.pkl"

logger = get_logger(__name__)
_trm_diag = get_diag_logger("hc_csv_tool.diag.trm_ex")
_perf = get_perf_logger("svc.svc_trm_ex.perf")
__version__ = "1.3.1"

_TRM_PHASE_TOTAL = 3


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


def _areas_to_tuples(areas: list[dict[str, Any]]) -> list[tuple[int, int, int, int]]:
    out: list[tuple[int, int, int, int]] = []
    for a in areas:
        yn_i = int(a.get("yn") or 0)
        xn_i = int(a.get("xn") or 0)
        if yn_i < 1 or xn_i < 1:
            continue
        out.append((int(a["y1"]), int(a["x1"]), yn_i, xn_i))
    return out


def _tuples_to_areas(tuples: list[tuple[int, int, int, int]]) -> list[dict[str, Any]]:
    return [{"y1": y1, "x1": x1, "yn": yn, "xn": xn} for y1, x1, yn, xn in tuples]


def _scan_trim_targets(
    arr: list[list[Any]],
    y1_i: int,
    x1_i: int,
) -> tuple[
    int,
    int,
    int,
    list[list[int]],
    list[list[int]],
    list[tuple[int, int, bool, bool]],
]:
    """文字列セルの文頭/文末空白を 1 パスで検出。適用対象 (ri, ci, lead, trail) も返す。"""
    n_leading = 0
    n_trailing = 0
    n_any_target = 0
    hl_leading: list[list[int]] = []
    hl_trailing: list[list[int]] = []
    targets: list[tuple[int, int, bool, bool]] = []
    for ri, row in enumerate(arr):
        for ci, cell in enumerate(row):
            if not isinstance(cell, str):
                continue
            lead_hit = cell != cell.lstrip()
            trail_hit = cell != cell.rstrip()
            if not lead_hit and not trail_hit:
                continue
            r_abs = y1_i + ri
            c_abs = x1_i + ci
            if lead_hit:
                n_leading += 1
                hl_leading.append([r_abs, c_abs, r_abs, c_abs])
            if trail_hit:
                n_trailing += 1
                hl_trailing.append([r_abs, c_abs, r_abs, c_abs])
            n_any_target += 1
            targets.append((ri, ci, lead_hit, trail_hit))
    return n_leading, n_trailing, n_any_target, hl_leading, hl_trailing, targets


def _apply_trim_targets(
    arr: list[list[Any]],
    targets: list[tuple[int, int, bool, bool]],
    choice: str,
    *,
    progress_cb: Any | None = None,
    total_target_cells: int = 1,
) -> tuple[int, int]:
    """検出済み対象セルのみ lstrip/rstrip/strip を適用。"""
    n_leading_done = 0
    n_trailing_done = 0
    done_cells = 0
    emit_step = max(1, int(total_target_cells) // 200)
    last_emit = -1
    last_emit_at = time.perf_counter()
    emit_interval = 0.05
    for ri, ci, lead_hit, trail_hit in targets:
        cell = arr[ri][ci]
        if not isinstance(cell, str):
            continue
        if choice == "leading":
            if not lead_hit:
                continue
            new_val = cell.lstrip()
            if new_val != cell:
                n_leading_done += 1
            arr[ri][ci] = new_val
        elif choice == "trailing":
            if not trail_hit:
                continue
            new_val = cell.rstrip()
            if new_val != cell:
                n_trailing_done += 1
            arr[ri][ci] = new_val
        else:
            new_val = cell.strip()
            if lead_hit:
                n_leading_done += 1
            if trail_hit:
                n_trailing_done += 1
            arr[ri][ci] = new_val
        done_cells += 1
        if progress_cb is None:
            continue
        now = time.perf_counter()
        if (
            done_cells >= total_target_cells
            or (done_cells - last_emit) >= emit_step
            or (now - last_emit_at) >= emit_interval
        ):
            try:
                progress_cb(done_cells)
            except Exception:
                pass
            last_emit = done_cells
            last_emit_at = now
    return n_leading_done, n_trailing_done


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


def _progress_write(path: Path, obj: dict[str, Any]) -> bool:
    return dispatch_progress_write(path, obj, log_tag="TRM_EX")


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
    progress_closed_path = progress_closed_ack_path("trm_ex", sid)
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

        trimmed = trim_areas_to_used_range(ptr_s, _areas_to_tuples(areas))
        if trimmed:
            areas = _tuples_to_areas(trimmed)
            areas_total = max(1, len(areas))
            total_rows = sum(int(a["yn"]) for a in areas)
            total_cols = max(int(a["xn"]) for a in areas) if areas else 0
            range_src = f"{range_src}_used_trim"
        _perf_trm(
            "after_used_range_trim",
            t_flow,
            source=range_src,
            areas=areas_total,
            yn=total_rows,
            xn=total_cols,
        )
        _trm_trace(
            "after_used_range_trim",
            t_flow,
            source=range_src,
            areas=areas_total,
            yn=total_rows,
            xn=total_cols,
        )

        if not areas:
            _status_bar_set(ptr_w, _msg(cfg, "NO_DATA"))
            _submit_no_target_ui(ph, sheet_id, _msg(cfg, "NO_DATA"), cfg)
            _perf_trm("early_no_data_after_trim", t_flow)
            _trm_trace("early_no_data_after_trim", t_flow)
            return

        scan_progress_started = False

        def _close_scan_progress() -> None:
            nonlocal progress_close_waited
            if not scan_progress_started or progress_close_waited:
                return
            _progress_write(
                progress_path,
                {
                    "status": "DONE",
                    "pct": 100,
                    "phase_i": 1,
                    "phase_total": _TRM_PHASE_TOTAL,
                    "phase": _msg(cfg, "PHASE_SCAN"),
                    "message": _msg(cfg, "PROGRESS_CUSTOM_SCAN"),
                    "done": 1,
                    "total": 1,
                },
            )
            wait_progress_closed_with_nudge(
                progress_closed_path,
                parent_hwnd=ph,
                sheet_id=sid,
                progress_path=progress_path,
                log_tag="TRM_EX",
            )
            progress_close_waited = True

        def _emit_scan_progress(pct: int, done: int, total: int, phase: str, custom: str) -> None:
            _progress_write(
                progress_path,
                {
                    "status": "RUN",
                    "phase_i": 1,
                    "phase_total": _TRM_PHASE_TOTAL,
                    "phase": phase,
                    "message": custom,
                    "pct": max(0, min(40, int(pct))),
                    "done": int(done),
                    "total": max(1, int(total)),
                },
            )

        try:
            reset_progress_closed_ack(progress_closed_path)
        except Exception:
            pass
        _emit_scan_progress(0, 0, 1, _msg(cfg, "PHASE_READ"), _msg(cfg, "PROGRESS_CUSTOM_READ"))
        _submit_progress_ui(
            ph,
            sid,
            progress_path,
            phase_total=_TRM_PHASE_TOTAL,
            progress_closed_path=progress_closed_path,
        )
        scan_progress_started = True
        time.sleep(0.25)

        msg_read = _msg(cfg, "PHASE_READ")
        msg_scan = _msg(cfg, "PHASE_SCAN")
        custom_read = _msg(cfg, "PROGRESS_CUSTOM_READ")
        custom_scan = _msg(cfg, "PROGRESS_CUSTOM_SCAN")

        # 各 area をチャンク読込し 2 次元化して保持
        area_payloads: list[dict[str, Any]] = []
        for a_idx, a in enumerate(areas, start=1):
            y1_i = int(a["y1"])
            x1_i = int(a["x1"])
            yn_i = int(a["yn"])
            xn_i = int(a["xn"])

            def _on_read_pct(pct: int, _phase: str, _custom: str, *, _a_idx: int = a_idx) -> None:
                _emit_scan_progress(
                    pct,
                    _a_idx - 1,
                    areas_total,
                    msg_read,
                    custom_read,
                )

            arr_i = read_sheet_matrix(
                ptr_s,
                y1_i,
                x1_i,
                yn_i,
                xn_i,
                _on_read_pct,
                msg_read=msg_read,
                custom_read=custom_read,
                normalize_2d=_normalize_2d,
            )
            if arr_i is None or len(arr_i) != yn_i:
                _close_scan_progress()
                _status_bar_set(ptr_w, _msg(cfg, "ERROR_PREFIX"))
                _perf_trm("abort_matrix_read_failed", t_flow, area_i=a_idx, areas=areas_total)
                _trm_trace("abort_matrix_read_failed", t_flow, area_i=a_idx, areas=areas_total)
                return
            area_payloads.append({"y1": y1_i, "x1": x1_i, "yn": yn_i, "xn": xn_i, "arr": arr_i})

        _perf_trm("after_matrix_read", t_flow, areas=areas_total, yn=total_rows, xn=total_cols)
        _trm_trace("after_matrix_read", t_flow, areas=areas_total, yn=total_rows, xn=total_cols)

        n_leading = 0
        n_trailing = 0
        n_any_target = 0
        hl_leading: list[list[int]] = []
        hl_trailing: list[list[int]] = []
        for ap in area_payloads:
            (
                nl,
                nt,
                na,
                hll,
                hlt,
                targets,
            ) = _scan_trim_targets(ap["arr"], int(ap["y1"]), int(ap["x1"]))
            n_leading += nl
            n_trailing += nt
            n_any_target += na
            hl_leading.extend(hll)
            hl_trailing.extend(hlt)
            ap["targets"] = targets

        _emit_scan_progress(
            40,
            areas_total,
            areas_total,
            msg_scan,
            custom_scan,
        )

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
            _close_scan_progress()
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
        _close_scan_progress()
        _submit_choice_ui(ph, sid, result_path, n_leading, n_trailing, cfg, highlight_trm)
        _perf_trm("after_choice_ui_submit", t_flow)
        _trm_trace("after_choice_ui_submit", t_flow)

        res = _poll_result(result_path)
        choice = (res or {}).get("choice") if isinstance(res, dict) else None
        if choice == "cancel" or choice is None:
            _close_scan_progress()
            _status_bar_restore(ptr_w, saved_status)
            logger.info("[TRM_EX] ユーザーキャンセル")
            _perf_trm("user_cancel", t_flow, choice=choice)
            _trm_trace("user_cancel", t_flow, choice=choice)
            return

        progress_close_waited = False
        try:
            reset_progress_closed_ack(progress_closed_path)
        except Exception:
            pass
        if choice == "leading":
            total_target_cells = max(1, int(n_leading))
        elif choice == "trailing":
            total_target_cells = max(1, int(n_trailing))
        else:
            total_target_cells = max(1, int(n_any_target))

        _progress_write(
            progress_path,
            {
                "status": "RUN",
                "phase_i": 2,
                "phase_total": _TRM_PHASE_TOTAL,
                "phase": _msg(cfg, "PHASE_APPLY"),
                "message": _msg(cfg, "PROGRESS_CUSTOM_APPLY"),
                "pct": 45,
                "done": 0,
                "total": total_target_cells,
                "seq": 0,
            },
        )
        _submit_progress_ui(
            ph,
            sid,
            progress_path,
            phase_total=_TRM_PHASE_TOTAL,
            progress_closed_path=progress_closed_path,
        )
        time.sleep(0.25)

        n_leading_done = 0
        n_trailing_done = 0

        def _on_apply_progress(done_cells: int) -> None:
            pct = min(74, max(45, int(45 + (29 * done_cells / total_target_cells))))
            _progress_write(
                progress_path,
                {
                    "status": "RUN",
                    "phase_i": 2,
                    "phase_total": _TRM_PHASE_TOTAL,
                    "phase": _msg(cfg, "PHASE_APPLY"),
                    "message": _msg(cfg, "PROGRESS_CUSTOM_APPLY"),
                    "pct": pct,
                    "done": done_cells,
                    "total": total_target_cells,
                },
            )

        for ap in area_payloads:
            ap["orig_arr"] = [list(row) for row in ap["arr"]]
            ld, td = _apply_trim_targets(
                ap["arr"],
                ap.get("targets") or [],
                str(choice),
                progress_cb=_on_apply_progress,
                total_target_cells=total_target_cells,
            )
            n_leading_done += ld
            n_trailing_done += td

        _progress_write(
            progress_path,
            {
                "status": "RUN",
                "phase_i": 3,
                "phase_total": _TRM_PHASE_TOTAL,
                "phase": _msg(cfg, "PHASE_WRITE"),
                "message": _msg(cfg, "PROGRESS_CUSTOM_WRITE"),
                "pct": 75,
                "done": 0,
                "total": total_target_cells,
            },
        )

        _perf_trm("after_apply_choice", t_flow, choice=choice)
        _trm_trace("after_apply_choice", t_flow, choice=choice)

        _snap_region = snapshot_region_from_areas(_areas_to_tuples(areas))
        try:
            from svc.svc_undo import save_undo_snapshot

            save_undo_snapshot(
                ptr_w,
                sheet_id=sheet_id,
                target_hwnd=ph,
                excel_hwnd=ph,
                snapshot_region=_snap_region,
            )
        except Exception as e:
            logger.warning("[TRM_EX] save_undo_snapshot failed (undo unavailable): %s", e)
        _perf_trm("after_undo_snapshot_attempt", t_flow)
        _trm_trace("after_undo_snapshot_attempt", t_flow)

        try:
            for ap in area_payloads:
                orig = ap.get("orig_arr") or ap["arr"]
                final = ap["arr"]

                def _on_write_progress(written: int) -> None:
                    pct = min(99, max(75, int(75 + (24 * written / max(1, total_target_cells)))))
                    _progress_write(
                        progress_path,
                        {
                            "status": "RUN",
                            "phase_i": 3,
                            "phase_total": _TRM_PHASE_TOTAL,
                            "phase": _msg(cfg, "PHASE_WRITE"),
                            "message": _msg(cfg, "PROGRESS_CUSTOM_WRITE"),
                            "pct": pct,
                            "done": min(written, total_target_cells),
                            "total": total_target_cells,
                        },
                    )

                write_changed_slices(
                    ptr_s,
                    int(ap["y1"]),
                    int(ap["x1"]),
                    orig,
                    final,
                    progress_cb=_on_write_progress,
                )
        except Exception as e:
            logger.exception("[TRM_EX] 書込失敗: %s", e)
            _status_bar_set(ptr_w, f"{_msg(cfg, 'ERROR_PREFIX')}: {e}")
            _progress_write(
                progress_path,
                {
                    "status": "DONE",
                    "pct": 100,
                    "phase_i": 3,
                    "phase_total": _TRM_PHASE_TOTAL,
                    "done": total_target_cells,
                    "total": total_target_cells,
                },
            )
            if not progress_close_waited:
                wait_progress_closed_with_nudge(
                    progress_closed_path,
                    parent_hwnd=ph,
                    sheet_id=sid,
                    progress_path=progress_path,
                    log_tag="TRM_EX",
                )
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
                "phase_i": 3,
                "phase_total": _TRM_PHASE_TOTAL,
                "done": total_target_cells,
                "total": total_target_cells,
            },
        )
        if not progress_close_waited:
            wait_progress_closed_with_nudge(
                progress_closed_path,
                parent_hwnd=ph,
                sheet_id=sid,
                progress_path=progress_path,
                log_tag="TRM_EX",
            )
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
