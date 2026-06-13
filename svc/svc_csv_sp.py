# -*- coding: utf-8 -*-
"""
Python: 3.10+
Module: svc/svc_csv_sp.py
Created: 2026-03-05
Updated: 2026-06-04
Version: 2.5.8
Purpose:
  CSVファイル分割（選択行による範囲分割）。アクティブシートの選択行を境界に分割し、
  各範囲をヘッダ付きで UTF-8(BOM) CSV として保存。データ不足時は Excel 中央でワーニング通知。

History (latest 3):
  - 2.5.8 (2026-06-13) 進捗表示: 保存中→分割保存中。同名確認「分割実施」後に保存されない不具合を修正。
  - 2.5.7 (2026-06-13) 進捗 UI 共通設定（poll/creep）と分割保存開始直前の砂時計 ON。
  - 2.5.6 (2026-06-06) 分割保存ループを suspend(restore_on_exit=False) 内で実行。DONE+wait も suspend 内。
  - 2.5.3 (2026-06-04) 保存: 通常 CSV 保存と同様、既定で表示文字列（Copy→クリップボード）でヘッダ・分割範囲を読込。HC_CSV_SV_USE_VALUE_READ=1 で .value 経路。
  - 2.5.2 (2026-06-04) 分割保存処理中の砂時計 ON（出力先確定後〜完了）。保存ループで tick 再武装。
  - 2.5.1 (2026-04-09) 重複キャンセル後の分割再表示で、再フォルダ選択まで split_csv が結果 pickle を監視し続ける（再オープン後の無処理を修正）。
  - 2.5.0 (2026-04-09) 分割→重複→進捗: 進捗 UI は重複解決後に IPC のみ。重複キャンセルで分割再表示 IPC。progress pickle 書込リトライ（WinError 5 等）。
  - 2.4.8 (2026-04-08) 完了 pickle に output_dir を付与。重複キャンセル時の診断ログ。
  - 2.4.7 (2026-04-08) rename_map 参照をキー正規化＋大文字小文字無視で重複名解決。手動リネームが確実に反映されるよう修正。
  - 2.4.6 (2026-04-08) 重複確認キャンセル時に進捗 pickle を CANCEL 化し進捗を閉じる（UI が分割に戻れるように）。
  - 2.4.5 (2026-04-08) 重複解決で連番リネームを廃止。指定名のまま保存し既存・同名は上書き。
  - 2.4.4 (2026-04-08) 未使用の _submit_progress_ui / _submit_done_ui を削除（進捗・完了は VBA 起動＋Pickle 経路に統一済みのため）。
  - 2.4.3 (2026-04-08) rename_map を重複一覧 index（sorted(dup_names)）で参照。done_items の name は実保存パス基準。
  - 2.4.2 (2026-04-08) 進捗 DONE pickle に show_done_dialog: True を付与（完了通知を ProgressDialog 経路に統一）。
  - 2.4.1 (2026-04-08) 完了ログの範囲文字列を実保存 plans 基準へ修正（削除後件数と不一致にならないよう調整）。
  - 2.4.0 (2026-04-08) 重複確認の返却を行インデックス管理へ変更（drop_rows / rename_map）。選択削除行は分割対象から除外し、1行も残らない場合は中止。
  - 2.3.1 (2026-04-08) 重複候補一覧を問い合わせメッセージへ付与（JSON: SHOW_DUPLICATE_LIST/MAX_LIST_ITEMS）。UI から rename_map を受け取り手動改名を優先。
  - 2.3.0 (2026-04-08) JSON 駆動の重複保存ポリシー（ask/overwrite/rename/cancel）を追加。完了通知 done_items に size_bytes を付与。重複判定・解決ログを追加。
  - 2.2.3 (2026-04-08) 分割結果 ranges の file_name を保存名に反映。進捗 pickle 書込失敗を warning ログ化（準備中で停止する診断を容易化）。
  - 2.2.2 (2026-04-06) _watch_ready に t_load0 渡し（csv_ld シグネチャ整合）。警告 UI に ready_path＋監視スレッド。運用・診断ログ、book=None で notify_wait_form_ready。
  - 2.2.1 (2026-04-06) READY_UI 監視スレッド追加（砂時計・VBA WaitForm と同タイミングで解除）。
  - 2.2.0 (2026-03-09) 進捗が 0/0 で止まらないよう、ループ前に初期 RUN（0/phase_total）を 1 回書き込み。
  - 2.1.0 (2026-03-06) データ不足・使用範囲取得失敗などで Excel 中央にワーニング通知。Excel 親子・前面・表示中無効・戻るとき有効。
  - 2.0.0 (2026-03-06) 仕様変更: キー列分割→選択行範囲分割。分割画面・進捗・終了通知フロー。core_log 使用。
"""
from __future__ import annotations

import csv
import os
import sys
import threading
import time
from contextlib import nullcontext
from pathlib import Path
from typing import Any, List, Optional, Tuple

_path_here = os.path.abspath(os.path.dirname(__file__))
_path_root = os.path.dirname(_path_here)
if _path_root not in sys.path:
    sys.path.insert(0, _path_root)

from core.core_log import get_diag_logger, get_logger
from core.core_progress_wait import wait_after_progress_done
from core.excel_display_read import read_range_display_text_matrix, use_display_text_for_csv_save
from core.core_cursor import notify_wait_form_ready, progress_dialog_wait_cursor_on
from core.csv_tool_progress_ui import enrich_progress_req_dict
from ui_qt.ipc_file import get_ipc_root, get_last_folder, get_request_dir, read_pickle, set_last_folder, write_pickle
from svc.svc_host import ensure_ui_server

logger = get_logger(__name__)
_sp_diag = get_diag_logger("hc_csv_tool.diag.csv_sp")
__version__ = "2.5.8"

_SP_CFG_CACHE: dict[str, Any] | None = None


def _elapsed_ms(since: float) -> int:
    return max(0, int((time.perf_counter() - since) * 1000))


def _sp_trace(fmt: str, *args: object) -> None:
    try:
        _sp_diag.info(fmt, *args)
    except Exception:
        pass


def _get_sp_cfg() -> dict[str, Any]:
    """config/ui_csv_sp.json を取得（失敗時は空辞書）。"""
    global _SP_CFG_CACHE
    if isinstance(_SP_CFG_CACHE, dict):
        return _SP_CFG_CACHE
    try:
        from core import core_cst as cst
        _SP_CFG_CACHE = cst.get_ui_config_from_file_required("csv_sp")
    except Exception:
        _SP_CFG_CACHE = {}
    return _SP_CFG_CACHE or {}

try:
    from core import core_xlc as xlc
    from core import core_stat
    from core import core_w32 as w32
except ImportError:
    xlc = None  # type: ignore[assignment]
    core_stat = None  # type: ignore[assignment]
    w32 = None  # type: ignore[assignment]

# 進捗画面表示待ち（秒）。UI が進捗窓を表示・ポーリング開始するまでの余裕。
_PROGRESS_STARTUP_WAIT_SEC: float = 2.0
# 進捗 DONE 後の進捗画面表示時間（秒）。仕様: 100% 後 1 秒で閉じる。
_DONE_DISPLAY_SEC: float = 1.0


def _watch_ready_ui(ready_path: str, sheet_id: str, t_enter0: float) -> None:
    """ui_server が ready_path に READY_UI を書いたら notify_ui_ready（WaitForm 含む）。"""
    from svc.svc_csv_ld import _watch_ready

    _watch_ready(ready_path, sheet_id, t_enter0)


def _get_sheet(book: Any, sheet_id: str) -> Any:
    """ブックと sheet_id から対象シートを取得。GUID 未設定時はアクティブシート。"""
    if sheet_id and xlc:
        sh = xlc.find_sheet_by_guid(book, sheet_id)
        if sh is not None:
            return sh
    try:
        return book.sheets.active
    except Exception:
        return None


def _get_selected_row_numbers(book: Any) -> List[int]:
    """
    【概要】
        Excel の現在の選択範囲から、選択されている行番号の一覧を取得する。
    【補足】
        複数領域（Areas）の場合はすべての行を集約し、ソート済みで返す。
        選択がない・取得失敗時は空リスト。
    """
    out: List[int] = []
    try:
        sel = book.app.selection
        if sel is None:
            return []
        api = getattr(sel, "api", None)
        if api is not None:
            n_areas = getattr(api.Areas, "Count", 0) or 0
            for i in range(1, int(n_areas) + 1):
                area = api.Areas(i)
                r1 = int(getattr(area, "Row", 0))
                r_count = int(getattr(area.Rows, "Count", 0))
                for r in range(r1, r1 + r_count):
                    out.append(r)
        if not out:
            r1 = int(getattr(sel, "row", 0))
            r_count = int(getattr(sel.rows, "count", 0))
            for r in range(r1, r1 + r_count):
                out.append(r)
    except Exception:
        pass
    return sorted(set(out))


def _build_initial_ranges(
    selected_rows: List[int],
    last_data_row: int,
) -> List[Tuple[int, int]]:
    """
    【概要】
        選択行とデータ最終行から、分割範囲（開始行, 終了行）のリストを組み立てる。
    【補足】
        1 行目はヘッダのためデータ開始は 2 行目。最初の分割の開始行は 2。
        次の選択行がない場合は last_data_row を終了行とする。
    """
    if last_data_row < 2:
        return []
    if not selected_rows:
        return [(2, last_data_row)]
    # 選択行は「次の分割の開始行」。最初の分割は 2 ～ (最初の選択行 - 1)
    ranges: List[Tuple[int, int]] = []
    start = 2
    for r in selected_rows:
        if r < 2:
            continue
        end = min(r - 1, last_data_row)
        if start <= end:
            ranges.append((start, end))
        start = r
    if start <= last_data_row:
        ranges.append((start, last_data_row))
    return ranges


def _submit_request_dict(req_dict: dict[str, Any]) -> Path:
    """UI サーバへリクエストを投入する。"""
    req_dir = get_request_dir()
    req_dir.mkdir(parents=True, exist_ok=True)
    ts_ms = int(time.time() * 1000)
    req_path = req_dir / f"req_{ts_ms}_{os.getpid()}_{threading.get_ident()}.pkl"
    write_pickle(req_path, req_dict)
    return req_path


def _progress_write_retry(path: Path, obj: dict[str, Any]) -> None:
    """進捗 pickle 書込（アクセス拒否等で短いリトライ）。"""
    last_exc: BaseException | None = None
    for attempt in range(5):
        try:
            write_pickle(path, obj)
            return
        except OSError as e:
            last_exc = e
            wn = getattr(e, "winerror", None)
            if wn is None:
                wn = getattr(e, "errno", None)
            if attempt < 4 and int(wn or 0) in (5, 13, 32, 11):
                time.sleep(0.05 * (attempt + 1))
                continue
            break
        except Exception as e:
            last_exc = e
            break
    logger.warning(
        "[CSV_SP] progress書込失敗 path=%s status=%s phase=%s err=%s",
        str(path),
        obj.get("status"),
        obj.get("phase"),
        last_exc,
    )


def _build_csv_sp_reopen_template(
    parent_hwnd: int,
    sheet_id: str,
    sheet_name: str,
    headers: List[Any],
    last_data_row: int,
    initial_ranges: List[dict[str, Any]],
    initial_dir: str,
) -> dict[str, Any]:
    return {
        "parent_hwnd": int(parent_hwnd or 0),
        "sheet_id": str(sheet_id or "_"),
        "sheet_name": str(sheet_name or "Sheet1"),
        "headers": list(headers or []),
        "last_data_row": int(last_data_row or 2),
        "initial_ranges": list(initial_ranges or []),
        "initial_dir": str(initial_dir or "").strip(),
    }


def _submit_csv_sp_split_reopen_ui(
    parent_hwnd: int,
    sheet_id: str,
    sheet_name: str,
    headers: List[Any],
    last_data_row: int,
    initial_ranges: List[dict[str, Any]],
    initial_dir: str,
    t_enter0: float,
) -> str:
    """重複確認キャンセル後など、分割画面を IPC で開き直す。戻り値は結果待ち用 result_path。"""
    ensure_ui_server()
    ipc_root = Path(get_ipc_root())
    res_dir = ipc_root / "results"
    ready_dir = ipc_root / "ready"
    res_dir.mkdir(parents=True, exist_ok=True)
    ready_dir.mkdir(parents=True, exist_ok=True)
    progress_path = _progress_path(sheet_id or "_")
    try:
        progress_path.unlink(missing_ok=True)
    except Exception:
        pass
    res_path = str(res_dir / f"res_sp_{sheet_id or '_'}_{int(time.time() * 1000)}.pkl")
    ready_path = str(ready_dir / f"ready_sp_{sheet_id or '_'}_{int(time.time() * 1000)}.pkl")
    try:
        Path(res_path).unlink(missing_ok=True)
        Path(ready_path).unlink(missing_ok=True)
    except Exception:
        pass
    req_dict: dict[str, Any] = {
        "parent_hwnd": int(parent_hwnd or 0),
        "result_path": res_path,
        "ready_path": ready_path,
        "progress_path": str(progress_path),
        "sheet_id": str(sheet_id or "_"),
        "action": "csv_sp",
        "module": "ui_qt.ui_csv_sp",
        "sheet_name": str(sheet_name or "Sheet1"),
        "headers": list(headers or []),
        "last_data_row": int(last_data_row or 2),
        "initial_ranges": list(initial_ranges or []),
        "initial_dir": str(initial_dir or "").strip(),
    }
    _submit_request_dict(req_dict)
    th_ready = threading.Thread(
        target=_watch_ready_ui,
        args=(ready_path, str(sheet_id or ""), t_enter0),
        name=f"ready_watch_sp_reopen_{sheet_id}",
        daemon=True,
    )
    th_ready.start()
    logger.info(
        "[CSV_SP] split_reopen_ui submitted sheet_id=%s hwnd=%s res=%s",
        sheet_id or "",
        int(parent_hwnd or 0),
        res_path,
    )
    return res_path


def _submit_csv_sp_progress_modeless_ui(
    parent_hwnd: int,
    sheet_id: str,
    progress_path: Path,
    phase_total: int,
    reopen_template: dict[str, Any],
) -> None:
    """重複解決後にモデルレス進捗を表示（分割画面は既に閉じている）。"""
    ensure_ui_server()
    ipc_root = Path(get_ipc_root())
    res_dir = ipc_root / "results"
    res_dir.mkdir(parents=True, exist_ok=True)
    ts_ms = int(time.time() * 1000)
    result_path = str(res_dir / f"res_sp_progress_{sheet_id or '_'}_{ts_ms}.pkl")
    try:
        Path(result_path).unlink(missing_ok=True)
    except Exception:
        pass
    ph = int(parent_hwnd or 0)
    req_inner: dict[str, Any] = enrich_progress_req_dict(
        {
            "action": "progress",
            "progress_path": str(progress_path),
            "phase_total": int(phase_total or 0),
            "no_native_window": True,
            "excel_lock": True,
            "partner_csv_sp_reopen_template": reopen_template,
        },
        done_delay_ms=400,
        no_native_window=True,
    )
    if w32 is not None and ph:
        try:
            er = w32.get_window_rect(ph)
            if er and len(er) >= 4:
                req_inner["excel_rect"] = [int(er[0]), int(er[1]), int(er[2]), int(er[3])]
        except Exception:
            pass
    payload: dict[str, Any] = {
        "parent_hwnd": ph,
        "result_path": result_path,
        "ready_path": "",
        "sheet_id": str(sheet_id or "_"),
        "action": "progress",
        "module": "ui_qt.ui_csv_sp",
        "req_dict": req_inner,
    }
    _submit_request_dict(payload)
    logger.info(
        "[CSV_SP] progress_modeless_ui submitted sheet_id=%s phase_total=%s",
        sheet_id or "",
        int(phase_total or 0),
    )


def _show_warning_dialog(
    book: Any,
    parent_hwnd: int,
    sheet_id: str,
    message: str,
    title: str = "ファイル分割",
) -> None:
    """
    【概要】
        ワーニングを Excel 中央に表示する。Excel 親子・前面・表示中は Excel 操作無効・OK で有効に戻す。
    """
    t_warn0 = time.perf_counter()
    ensure_ui_server()
    ph = int(parent_hwnd or 0)
    if not ph:
        try:
            ph = int(getattr(book.app, "hwnd", 0))
        except Exception:
            pass
    if not ph:
        logger.warning("[CSV_SP] phase=warn_skip_no_hwnd sheet_id=%s", sheet_id or "")
        try:
            notify_wait_form_ready(parent_hwnd=ph)
        except Exception:
            pass
        if core_stat:
            try:
                core_stat.set_status_info(book.sheets.active, message)
            except Exception:
                pass
        return
    ipc_root = Path(get_ipc_root())
    res_dir = ipc_root / "results"
    res_dir.mkdir(parents=True, exist_ok=True)
    ready_dir = ipc_root / "ready"
    ready_dir.mkdir(parents=True, exist_ok=True)
    res_path = str(res_dir / f"res_sp_warn_{sheet_id or '_'}_{int(time.time() * 1000)}.pkl")
    warn_ready_path = str(ready_dir / f"ready_sp_warn_{sheet_id or '_'}_{int(time.time() * 1000)}.pkl")
    try:
        Path(res_path).unlink(missing_ok=True)
        Path(warn_ready_path).unlink(missing_ok=True)
    except Exception:
        pass
    req_dict = {
        "parent_hwnd": ph,
        "result_path": res_path,
        "ready_path": warn_ready_path,
        "sheet_id": str(sheet_id or "_"),
        "action": "csv_sp_warning",
        "module": "ui_qt.ui_csv_sp",
        "message": message,
        "title": title,
    }
    req_path = _submit_request_dict(req_dict)
    logger.info(
        "[CSV_SP] ui_ipc ok kind=sp_warn req=%s sheet_id=%s hwnd=%s",
        req_path.name,
        sheet_id or "",
        ph,
    )
    _sp_trace(
        "[CSV_SP_TRACE] ui_ipc ok kind=sp_warn req=%s elapsed_ms=%s wall_perf_s=%.6f",
        req_path.name,
        _elapsed_ms(t_warn0),
        time.perf_counter(),
    )
    th_warn = threading.Thread(
        target=_watch_ready_ui,
        args=(warn_ready_path, str(sheet_id or ""), t_warn0),
        name=f"ready_watch_sp_warn_{sheet_id}",
        daemon=True,
    )
    th_warn.start()
    _watch_result_sp(res_path, timeout_sec=30.0)
    logger.info(
        "[CSV_SP] phase=warn_flow_done elapsed_ms=%s sheet_id=%s",
        _elapsed_ms(t_warn0),
        sheet_id or "",
    )
    if ph and w32:
        try:
            w32.bring_to_front(ph)
        except Exception:
            pass


def _watch_result_sp(result_path: str, timeout_sec: float = 120.0) -> Optional[dict[str, Any]]:
    """UI 結果を待ち、OK なら output_dir / base_filename / ranges を返す。"""
    p = Path(result_path)
    t0 = time.time()
    while (time.time() - t0) < timeout_sec:
        if p.exists() and p.stat().st_size > 0:
            try:
                d = read_pickle(p)
            except Exception:
                time.sleep(0.05)
                continue
            status = str(d.get("status", "")).strip().upper()
            if status == "OK":
                return d
            return None
        time.sleep(0.05)
    return None


def _ask_conflict_policy(
    parent_hwnd: int,
    sheet_id: str,
    message: str,
    conflict_cfg: dict[str, Any],
    dup_names: List[str],
) -> tuple[str, dict[str, str], List[int]]:
    """重複保存時の選択を UI に問い合わせる。戻り値: (choice, rename_map, drop_rows)。"""
    try:
        ensure_ui_server()
        ipc_root = Path(get_ipc_root())
        res_dir = ipc_root / "results"
        res_dir.mkdir(parents=True, exist_ok=True)
        ts_ms = int(time.time() * 1000)
        res_path = str(res_dir / f"res_sp_conflict_{sheet_id or '_'}_{ts_ms}.pkl")
        try:
            Path(res_path).unlink(missing_ok=True)
        except Exception:
            pass
        req_dict = {
            "parent_hwnd": int(parent_hwnd or 0),
            "result_path": res_path,
            "ready_path": "",
            "sheet_id": str(sheet_id or "_"),
            "action": "csv_sp_conflict",
            "module": "ui_qt.ui_csv_sp",
            "message": message,
            "conflict_cfg": conflict_cfg or {},
            "dup_names": list(dup_names or []),
        }
        req_path = _submit_request_dict(req_dict)
        logger.info("[CSV_SP] conflict問い合わせ req=%s", req_path.name)
        d = _watch_result_sp(res_path, timeout_sec=120.0)
        if isinstance(d, dict) and str(d.get("status", "")).strip().upper() == "OK":
            ch = str(d.get("choice", "")).strip().lower()
            if ch in ("overwrite", "rename", "apply", "cancel"):
                rm = d.get("rename_map") or {}
                if not isinstance(rm, dict):
                    rm = {}
                cleaned: dict[str, str] = {}
                for k, v in rm.items():
                    sk = str(k or "").strip()
                    sv = str(v or "").strip()
                    if sk and sv:
                        cleaned[sk] = sv
                dr = d.get("drop_rows") or []
                if not isinstance(dr, list):
                    dr = []
                drops: List[int] = []
                for x in dr:
                    try:
                        drops.append(int(x))
                    except Exception:
                        pass
                return ch, cleaned, drops
    except Exception as e:
        logger.warning("[CSV_SP] conflict問い合わせ失敗: %s", e)
    return "cancel", {}, []


def _canonical_dup_name(nm: str, dup_set: set[str]) -> str | None:
    """dup_set に含まれる表記を返す（Windows 等での大文字小文字差を吸収）。"""
    if nm in dup_set:
        return nm
    low = nm.lower()
    for x in dup_set:
        if x.lower() == low:
            return x
    return None


def _rename_map_get(rm: dict[str, Any], k_idx: str | None) -> str:
    """UI 由来の rename_map から値を取得（キーは str / int 両対応）。"""
    if not rm or k_idx is None:
        return ""
    s = str(k_idx).strip()
    if not s:
        return ""
    v = rm.get(s)
    if v is None:
        try:
            v = rm.get(str(int(s)))
        except Exception:
            v = None
    if v is None:
        return ""
    return str(v).strip()


def _plans_after_conflict_drop_rows(
    plans: List[dict[str, Any]],
    drop_rows: List[int],
    dup_names: set[str],
) -> List[dict[str, Any]]:
    """重複確認テーブルで削除した行（dup_names インデックス）に対応する plan を除外する。"""
    if not drop_rows or not dup_names:
        return plans
    dup_list = sorted(dup_names)
    drop_low: set[str] = set()
    for idx in drop_rows:
        try:
            i = int(idx)
            if 0 <= i < len(dup_list):
                drop_low.add(dup_list[i].lower())
        except (TypeError, ValueError):
            continue
    if not drop_low:
        return plans
    return [
        p
        for p in plans
        if str(p.get("file_name", "")).strip().lower() not in drop_low
    ]


def _resolve_duplicate_output(
    plans: List[dict[str, Any]],
    output_dir: str,
    conflict_cfg: dict[str, Any],
    parent_hwnd: int,
    sheet_id: str,
) -> tuple[str, List[dict[str, Any]]]:
    """
    予定保存ファイルの重複を解決する。
    returns: (choice, updated_plans)
    """
    when_exists = str(conflict_cfg.get("WHEN_EXISTS") or "ask").strip().lower()
    if when_exists not in ("ask", "overwrite", "rename", "cancel", "apply"):
        when_exists = "ask"
    enabled = bool(conflict_cfg.get("ENABLED", True))
    if not enabled:
        return "overwrite", plans

    # 既存ファイル重複 + 同一バッチ内重複を抽出
    seen_names: set[str] = set()
    dup_names: set[str] = set()
    for p in plans:
        nm = str(p.get("file_name", "")).strip()
        if not nm:
            continue
        lp = nm.lower()
        if lp in seen_names:
            dup_names.add(nm)
        seen_names.add(lp)
        if os.path.exists(os.path.join(output_dir, nm)):
            dup_names.add(nm)

    if not dup_names:
        return "overwrite", plans

    logger.info("[CSV_SP] 重複候補検出 count=%s names=%s", len(dup_names), ",".join(sorted(dup_names)[:5]))
    choice = when_exists
    rename_map: dict[str, str] = {}
    drop_rows: List[int] = []
    if when_exists == "ask":
        msg_default = (
            "保存先に同名ファイルが存在します。\n"
            "上書き / ファイル名変更 / キャンセル を選択してください。"
        )
        ask_cfg = (conflict_cfg.get("ASK_DIALOG") or {})
        msg = str(ask_cfg.get("MSG") or msg_default).strip()
        if bool(ask_cfg.get("SHOW_DUPLICATE_LIST", True)):
            try:
                max_items = int(ask_cfg.get("MAX_LIST_ITEMS") or 10)
            except Exception:
                max_items = 10
            max_items = max(1, min(50, max_items))
            listed = sorted(dup_names)[:max_items]
            suffix = ""
            if len(dup_names) > max_items:
                suffix = f"\n... 他 {len(dup_names) - max_items} 件"
            msg = f"{msg}\n\n重複ファイル:\n- " + "\n- ".join(listed) + suffix
        choice, rename_map, drop_rows = _ask_conflict_policy(parent_hwnd, sheet_id, msg, conflict_cfg, sorted(dup_names))
    if choice == "cancel":
        return "cancel", plans
    if choice == "overwrite":
        return "overwrite", plans

    # apply/rename: drop_rows は重複テーブル行インデックス（plan インデックスではない）
    if drop_rows:
        plans = _plans_after_conflict_drop_rows(plans, drop_rows, dup_names)
    if not plans:
        return "cancel", plans

    # rename / apply: 連番による救済リネームは行わず、指定ファイル名のまま保存（既存・同一バッチ内の同名は上書き）
    dup_sorted = sorted(dup_names)
    rm_norm: dict[str, str] = {}
    for k, v in (rename_map or {}).items():
        sk = str(k).strip()
        sv = str(v).strip() if v is not None else ""
        if sk and sv:
            rm_norm[sk] = sv
    rename_map = rm_norm

    for p in plans:
        nm = str(p.get("file_name", "")).strip()
        if not nm:
            continue
        cand = nm
        canon = _canonical_dup_name(nm, dup_names)
        k_idx: str | None = None
        if canon is not None:
            try:
                k_idx = str(dup_sorted.index(canon))
            except ValueError:
                k_idx = None
        manual = _rename_map_get(rename_map, k_idx)
        if manual:
            if not manual.lower().endswith(".csv"):
                manual = f"{manual}.csv"
            cand = manual
        if os.path.exists(os.path.join(output_dir, cand)):
            logger.info("[CSV_SP] 重複解決 上書き予定 path=%s (old_plan=%s)", cand, nm)
        p["file_name"] = cand
    return "apply", plans


def _sp_read_header_row(ptr_s: Any, ncols: int, *, use_display_text: bool) -> list[str]:
    if use_display_text:
        rows = read_range_display_text_matrix(
            ptr_s, row_start=1, col_start=1, n_rows=1, n_cols=ncols
        )
        if rows:
            return rows[0]
        return [""] * ncols
    row1 = ptr_s.range((1, 1), (1, ncols)).value
    if isinstance(row1, list):
        return [str(x) if x is not None else "" for x in row1]
    return [str(row1)] if row1 is not None else [""]


def _sp_read_body_rows(
    ptr_s: Any,
    start_row: int,
    end_row: int,
    ncols: int,
    *,
    use_display_text: bool,
) -> list[list[str]]:
    n_rows = end_row - start_row + 1
    if use_display_text:
        return read_range_display_text_matrix(
            ptr_s,
            row_start=start_row,
            col_start=1,
            n_rows=n_rows,
            n_cols=ncols,
        )
    rows_data = ptr_s.range((start_row, 1), (end_row, ncols)).value
    if isinstance(rows_data, (list, tuple)):
        return [
            [str(c) if c is not None else "" for c in (list(r) if isinstance(r, (list, tuple)) else [r])]
            for r in rows_data
        ]
    return [[str(rows_data) if rows_data is not None else ""]]


def _progress_path(sheet_id: str) -> Path:
    """進捗用 pickle のパスを返す。"""
    root = Path(get_ipc_root())
    d = root / "progress"
    d.mkdir(parents=True, exist_ok=True)
    return d / f"progress_sp_{sheet_id or '_'}.pkl"


def split_csv(
    book: Any,
    sheet_id: str = "",
    target_hwnd: Optional[int] = None,
    excel_hwnd: Optional[int] = None,
    **kwargs: Any,
) -> None:
    """
    【概要】
        ファイル分割のエントリ。アクティブシートの選択行から初期分割範囲を組み立て、
        分割画面で編集後にフォルダ選択 → 進捗表示で保存 → 終了通知。
    """
    t_enter0 = time.perf_counter()
    logger.info(
        "[CSV_SP] 開始 sheet_id=%s svc_pid=%s",
        sheet_id or "",
        os.getpid(),
    )
    _sp_trace(
        "[CSV_SP_TRACE] phase=enter sheet_id=%s svc_pid=%s wall_perf_s=%.6f",
        sheet_id or "",
        os.getpid(),
        time.perf_counter(),
    )
    hwnd = int(target_hwnd or excel_hwnd or 0)
    if book is None:
        logger.error("[CSV_SP] book=None のため中断（WaitForm 解除を試行）")
        try:
            notify_wait_form_ready(parent_hwnd=hwnd)
        except Exception:
            pass
        return

    parent_hwnd = hwnd
    try:
        if not parent_hwnd:
            parent_hwnd = int(getattr(book.app, "hwnd", 0))
    except Exception:
        pass

    ptr_s = _get_sheet(book, sheet_id)
    if ptr_s is None:
        logger.warning("[CSV_SP] 対象シートなし sheet_id=%s", sheet_id)
        _show_warning_dialog(book, parent_hwnd, sheet_id, "シートを特定できませんでした。", "ファイル分割")
        return

    try:
        sheet_name = str(getattr(ptr_s, "name", "") or "").strip() or "Sheet1"
    except Exception:
        sheet_name = "Sheet1"

    ur = getattr(ptr_s, "used_range", None)
    if ur is None:
        _show_warning_dialog(book, parent_hwnd, sheet_id, "使用範囲を取得できませんでした。", "ファイル分割")
        return
    nr = getattr(ur, "rows", None)
    nc = getattr(ur, "columns", None)
    if nr is None or nc is None:
        _show_warning_dialog(book, parent_hwnd, sheet_id, "行/列を取得できませんでした。", "ファイル分割")
        return
    nrows = int(nr.count)
    ncols = int(nc.count)
    if nrows < 2 or ncols < 1:
        # メッセージは空にして config/ui_csv_sp.json の SCREENS.WARNING.MSG を表示（改行・句点は JSON で管理）
        _show_warning_dialog(book, parent_hwnd, sheet_id, "", "ファイル分割")
        return

    use_display_text = use_display_text_for_csv_save()
    headers = _sp_read_header_row(ptr_s, ncols, use_display_text=use_display_text)

    last_data_row = nrows
    selected_rows = _get_selected_row_numbers(book)
    initial_ranges = _build_initial_ranges(selected_rows, last_data_row)
    if not initial_ranges:
        _show_warning_dialog(book, parent_hwnd, sheet_id, "分割範囲を組み立てられませんでした。", "ファイル分割")
        return

    ensure_ui_server()
    logger.info(
        "[CSV_SP] phase=after_ensure_ui_server elapsed_ms=%s",
        _elapsed_ms(t_enter0),
    )
    _sp_trace(
        "[CSV_SP_TRACE] phase=after_ensure_ui_server elapsed_ms=%s wall_perf_s=%.6f",
        _elapsed_ms(t_enter0),
        time.perf_counter(),
    )

    ipc_root = Path(get_ipc_root())
    res_dir = ipc_root / "results"
    res_dir.mkdir(parents=True, exist_ok=True)
    ready_dir = ipc_root / "ready"
    ready_dir.mkdir(parents=True, exist_ok=True)
    res_path = str(res_dir / f"res_sp_{sheet_id or '_'}_{int(time.time() * 1000)}.pkl")
    ready_path = str(ready_dir / f"ready_sp_{sheet_id or '_'}_{int(time.time() * 1000)}.pkl")
    try:
        Path(res_path).unlink(missing_ok=True)
        Path(ready_path).unlink(missing_ok=True)
    except Exception:
        pass

    initial_dir = get_last_folder() or ""
    if not initial_dir and getattr(book, "fullname", None):
        initial_dir = os.path.dirname(str(book.fullname))

    progress_path = _progress_path(sheet_id or "_")
    try:
        progress_path.unlink(missing_ok=True)
    except Exception:
        pass

    req_dict = {
        "parent_hwnd": parent_hwnd,
        "result_path": res_path,
        "ready_path": ready_path,
        "progress_path": str(progress_path),
        "sheet_id": str(sheet_id or "_"),
        "action": "csv_sp",
        "module": "ui_qt.ui_csv_sp",
        "sheet_name": sheet_name,
        "headers": headers,
        "initial_ranges": initial_ranges,
        "last_data_row": last_data_row,
        "initial_dir": initial_dir,
    }

    req_path = _submit_request_dict(req_dict)
    logger.info(
        "[CSV_SP] ui_ipc ok req=%s sheet_id=%s hwnd=%s",
        req_path.name,
        sheet_id or "",
        parent_hwnd,
    )
    _sp_trace(
        "[CSV_SP_TRACE] ui_ipc ok req=%s sheet_id=%s hwnd=%s elapsed_ms=%s wall_perf_s=%.6f",
        req_path.name,
        sheet_id or "",
        parent_hwnd,
        _elapsed_ms(t_enter0),
        time.perf_counter(),
    )

    th_ready = threading.Thread(
        target=_watch_ready_ui,
        args=(ready_path, str(sheet_id or ""), t_enter0),
        name=f"ready_watch_sp_{sheet_id}",
        daemon=True,
    )
    th_ready.start()

    res_path_active = res_path
    while True:
        result = _watch_result_sp(res_path_active)
        if not result or str(result.get("status", "")).strip().upper() != "OK":
            logger.info(
                "[CSV_SP] phase=result_cancel elapsed_ms=%s sheet_id=%s",
                _elapsed_ms(t_enter0),
                sheet_id or "",
            )
            _sp_trace(
                "[CSV_SP_TRACE] phase=result_cancel elapsed_ms=%s wall_perf_s=%.6f",
                _elapsed_ms(t_enter0),
                time.perf_counter(),
            )
            if hwnd and w32:
                try:
                    w32.bring_to_front(hwnd)
                except Exception:
                    pass
            return

        logger.info(
            "[CSV_SP] phase=result_ok elapsed_ms=%s sheet_id=%s",
            _elapsed_ms(t_enter0),
            sheet_id or "",
        )
        _sp_trace(
            "[CSV_SP_TRACE] phase=result_ok elapsed_ms=%s wall_perf_s=%.6f",
            _elapsed_ms(t_enter0),
            time.perf_counter(),
        )

        output_dir = str(result.get("output_dir", "")).strip()
        base_filename = str(result.get("base_filename", "")).strip()
        ranges_raw = result.get("ranges") or []
        if not output_dir or not base_filename or not isinstance(ranges_raw, list) or not ranges_raw:
            logger.warning("[CSV_SP] 無効な結果 output_dir/base_filename/ranges")
            if hwnd and w32:
                try:
                    w32.bring_to_front(hwnd)
                except Exception:
                    pass
            return

        try:
            set_last_folder(output_dir)
        except Exception:
            pass

        ranges_list: List[dict[str, Any]] = []
        for r in ranges_raw:
            if isinstance(r, dict):
                s = int(r.get("start_row", 0))
                e = int(r.get("end_row", 0))
                fn = str(r.get("file_name", "")).strip()
                if s >= 2 and e >= s:
                    ranges_list.append({"start_row": s, "end_row": e, "file_name": fn})
            elif isinstance(r, (list, tuple)) and len(r) >= 2:
                s, e = int(r[0]), int(r[1])
                if s >= 2 and e >= s:
                    ranges_list.append({"start_row": s, "end_row": e, "file_name": ""})

        if not ranges_list:
            logger.warning("[CSV_SP] 有効な範囲なし")
            if hwnd and w32:
                try:
                    w32.bring_to_front(hwnd)
                except Exception:
                    pass
            return

        reopen_ranges: List[dict[str, Any]] = [
            {
                "start_row": int(r["start_row"]),
                "end_row": int(r["end_row"]),
                "file_name": str(r.get("file_name") or "").strip(),
            }
            for r in ranges_list
        ]
        reopen_template = _build_csv_sp_reopen_template(
            parent_hwnd,
            str(sheet_id or "_"),
            sheet_name,
            headers,
            last_data_row,
            reopen_ranges,
            output_dir,
        )

        try:
            progress_path.unlink(missing_ok=True)
        except Exception:
            pass

        # 重複保存ポリシー（JSON）: ask / overwrite / rename / cancel
        cfg_root = _get_sp_cfg()
        main_cfg = (cfg_root or {}).get("MAIN") or {}
        conflict_cfg = (main_cfg.get("CONFLICT") or (cfg_root or {}).get("CONFLICT") or {})

        plans: List[dict[str, Any]] = []
        for idx, r in enumerate(ranges_list):
            phase_i = idx + 1
            start_row = int(r["start_row"])
            end_row = int(r["end_row"])
            file_name_base = str(r.get("file_name") or "").strip()
            if not file_name_base:
                file_name_base = f"{base_filename}_{phase_i}"
            if file_name_base.lower().endswith(".csv"):
                file_name_base = file_name_base[:-4].rstrip() or f"{base_filename}_{phase_i}"
            file_name = f"{file_name_base}.csv"
            plans.append(
                {
                    "phase_i": phase_i,
                    "start_row": start_row,
                    "end_row": end_row,
                    "row_count": end_row - start_row + 1,
                    "file_name": file_name,
                }
            )

        choice, plans = _resolve_duplicate_output(plans, output_dir, conflict_cfg, parent_hwnd, sheet_id)
        logger.info("[CSV_SP] 重複ポリシー choice=%s planned_files=%s", choice, len(plans))
        if choice == "cancel":
            logger.info("[CSV_SP] 重複確認でキャンセル → 分割再表示を待機")
            _progress_write_retry(progress_path, {"status": "CANCEL"})
            logger.info(
                "[CSV_SP_DIAG] progress_cancel_written path=%s sheet_id=%s abs=%s",
                str(progress_path),
                sheet_id or "",
                str(progress_path.resolve()),
            )
            try:
                res_path_active = _submit_csv_sp_split_reopen_ui(
                    parent_hwnd,
                    str(sheet_id or "_"),
                    sheet_name,
                    headers,
                    last_data_row,
                    reopen_ranges,
                    output_dir,
                    t_enter0,
                )
            except Exception as _re_exc:
                logger.warning("[CSV_SP] split_reopen_ui 失敗: %s", _re_exc)
                if hwnd and w32:
                    try:
                        w32.bring_to_front(hwnd)
                    except Exception:
                        pass
                return
            if hwnd and w32:
                try:
                    w32.bring_to_front(hwnd)
                except Exception:
                    pass
            continue

        break

    phase_total = len(plans)
    total_rows = sum(int(p["row_count"]) for p in plans)

    try:
        progress_dialog_wait_cursor_on(str(sheet_id or "progress"))
    except Exception:
        pass
    try:
        _submit_csv_sp_progress_modeless_ui(
            parent_hwnd,
            str(sheet_id or "_"),
            progress_path,
            phase_total,
            reopen_template,
        )
        time.sleep(0.25)
    except Exception as _pui_exc:
        logger.warning("[CSV_SP] progress_modeless_ui 失敗: %s", _pui_exc)

    def _progress_write(obj: dict[str, Any]) -> None:
        _progress_write_retry(progress_path, obj)

    # 進捗が 0/0 で止まらないよう、ループ前に初期状態を 1 回書く（UI がすぐ読めるように）
    _progress_write({
        "status": "RUN",
        "phase_i": 0,
        "phase_total": phase_total,
        "phase": "準備中",
        "current_file": "",
        "done": 0,
        "total": phase_total,
        "pct": 0,
    })

    logger.info(
        "[CSV_SP] phase=before_save_sleep sleep_sec=0.3 elapsed_ms=%s",
        _elapsed_ms(t_enter0),
    )
    time.sleep(0.3)

    done_accum = 0
    success_count = 0
    done_items: List[dict[str, Any]] = []

    use_display_text = use_display_text_for_csv_save()
    t_save_loop0 = time.perf_counter()
    total_size_bytes = 0
    logger.info("[CSV_SP] phase=split_save_loop_enter display_text=%s files=%s", use_display_text, len(plans))
    try:
        with (xlc.suspend_sheet_updates(ptr_s, restore_on_exit=False) if xlc else nullcontext()):
            for p in plans:
                start_row = int(p["start_row"])
                end_row = int(p["end_row"])
                phase_i = int(p["phase_i"])
                row_count = int(p["row_count"])
                file_name = str(p["file_name"])
                out_path = os.path.join(output_dir, file_name)

                _progress_write({
                    "status": "RUN",
                    "phase_i": phase_i,
                    "phase_total": phase_total,
                    "phase": "分割保存中",
                    "current_file": file_name,
                    "done": phase_i,
                    "total": phase_total,
                    "pct": int(100 * (phase_i - 1) / phase_total) if phase_total else 0,
                })

                try:
                    rows = _sp_read_body_rows(
                        ptr_s,
                        start_row,
                        end_row,
                        ncols,
                        use_display_text=use_display_text,
                    )
                    with open(out_path, "w", encoding="utf-8-sig", newline="", errors="replace") as f:
                        writer = csv.writer(f, quoting=csv.QUOTE_MINIMAL)
                        writer.writerow(headers)
                        for row in rows:
                            writer.writerow(row)
                    success_count += 1
                    try:
                        size_bytes = int(os.path.getsize(out_path))
                    except Exception:
                        size_bytes = 0
                    total_size_bytes += max(0, size_bytes)
                    done_items.append(
                        {
                            "no": phase_i,
                            "name": os.path.basename(out_path),
                            "rows": row_count,
                            "size_bytes": size_bytes,
                        }
                    )
                    logger.info("[CSV_SP] 保存完了 file=%s rows=%s size_bytes=%s", file_name, row_count, size_bytes)
                except Exception as e:
                    logger.warning("[CSV_SP] 書込失敗 ファイル=%s: %s", os.path.basename(out_path), e)

                done_accum += row_count
                pct = min(99, int(100 * done_accum / total_rows)) if total_rows else 99
                _progress_write({
                    "status": "RUN",
                    "phase_i": phase_i,
                    "phase_total": phase_total,
                    "phase": "分割保存中",
                    "current_file": file_name,
                    "done": done_accum,
                    "total": total_rows,
                    "pct": pct,
                })

            _progress_write({
                "status": "DONE",
                "phase_i": phase_total,
                "phase_total": phase_total,
                "done": total_rows,
                "total": total_rows,
                "pct": 100,
                "show_done_dialog": True,
                "done_items": done_items,
                "total_size_bytes": total_size_bytes,
                "output_dir": output_dir,
            })
            wait_after_progress_done(min_sec=_DONE_DISPLAY_SEC + 0.5)
    finally:
        if xlc is not None:
            try:
                xlc.restore_screen_updating(ptr_s)
            except Exception:
                pass

    save_loop_ms = _elapsed_ms(t_save_loop0)
    logger.info(
        "[CSV_SP] phase=split_save_loop_done save_loop_ms=%s files=%s elapsed_ms=%s",
        save_loop_ms,
        len(ranges_list),
        _elapsed_ms(t_enter0),
    )
    _sp_trace(
        "[CSV_SP_TRACE] phase=split_save_loop_done save_loop_ms=%s wall_perf_s=%.6f",
        save_loop_ms,
        time.perf_counter(),
    )

    # 完了通知は ProgressDialog の DONE pickle（show_done_dialog）に統一。分割画面の exec 解放は partner_widget_after_done

    ranges_str = ",".join(f"{int(p['start_row'])}-{int(p['end_row'])}" for p in plans)
    msg = (
        f"ファイル分割が正常に完了しました。 | "
        f"シート: {sheet_name} | {success_count} ファイル | 出力先: {output_dir}"
    )
    if core_stat:
        try:
            core_stat.set_status_info(ptr_s, msg)
        except Exception:
            pass
    logger.info(
        "[CSV_SP] 完了 シート=%s 分割数=%s 出力先=%s 範囲=%s 総行数=%s 総容量B=%s",
        sheet_name, success_count, output_dir, ranges_str, total_rows, total_size_bytes,
    )
    logger.info(
        "[CSV_SP] phase=split_csv_flow_done elapsed_ms=%s sheet_id=%s",
        _elapsed_ms(t_enter0),
        sheet_id or "",
    )
    _sp_trace(
        "[CSV_SP_TRACE] phase=split_csv_flow_done elapsed_ms=%s wall_perf_s=%.6f",
        _elapsed_ms(t_enter0),
        time.perf_counter(),
    )

    if hwnd and w32:
        try:
            w32.bring_to_front(hwnd)
        except Exception:
            pass
