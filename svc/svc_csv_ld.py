# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: svc/svc_csv_ld.py
Created: 2026-03-05
Updated: 2026-04-10
Version: 1.3.13
Purpose:
  CSV読込（Qt UIサーバ方式 / 2プロセス分離）。
  期待フロー: CSVファイル選択 → 進捗画面表示（準備中→ファイル解析中…）→ CSV読込 → Excelシート出力 → セルオートフィット → 進捗画面閉じる → 完了通知。
  進捗は 0=準備中 1=ファイル解析 2=Excel書き込み 3=列幅調整 4=完了 の5段階で表示。無表示1秒未満のため ui_server がファイル選択OK直後に進捗を即表示する経路では progress_ui_already_shown で二重依頼を避ける。

History (latest 3):
  - 1.3.13 (2026-04-10) 計測: `pick_to_done_ms`（ファイル確定〜`load_csv_flow_done`）。`phase=pick_confirmed` で区間 B 始点。docs/csv_ld_perf_measurement.md 参照。
  - 1.3.12 (2026-04-07) ファイル選択後・処理完了後に core_w32.bring_to_front で Excel 前面復帰（背面に回る事象の緩和）。ui_csv_ld のダイアログ終了時も同対応。
  - 1.3.11 (2026-04-06) 環境変数: HC_PROGRESS_WINDOW_STARTUP_WAIT_SEC / HC_EXCEL_HWND を core.core_env 経由に統一。
  - 1.3.10 (2026-04-06) 運用ログ・診断ログ: load_csv / READY_UI / 結果待ち / do_load の phase と経過 ms、req 相関。book=None 時は notify_wait_form_ready。
  - 1.3.9 (2026-03-13) 無表示1秒未満化: _watch_result からは progress_ui_already_shown=True で呼び出し、準備中ブロックをスキップ。ui_server が即表示する前提。
  - 1.3.8 (2026-03-13) 進捗をファイル選択直後に表示（準備中）。find_sheet_by_guid/行数取得の前に進捗UI依頼し体感遅延を短縮。早期return時はDONEで閉じる。
  - 1.3.7 (2026-03-11) 進捗まわり: 表示待ちを HC_PROGRESS_WINDOW_STARTUP_WAIT_SEC で環境変数化（既定1.0秒）。進捗に seq を付与して順序保証。
  - 1.3.6 (2026-03-09) オートフィットをセル数→行数基準に変更。定数は core_cst.AUTOFIT_MAX_ROWS に統一。
  - 1.3.5 (2026-03-09) 一括書込みを表示停止で高速化。書込み後は有効領域外を clear_used_range_overflow でクリア。
"""
from __future__ import annotations

import ctypes
import math
import os
import threading
import time
from pathlib import Path
from typing import Any

from core import core_env
from core.core_log import get_diag_logger, get_logger
from core.core_cursor import notify_ui_ready, notify_wait_form_ready
from ui_qt.ipc_file import get_ipc_root, get_last_folder, get_request_dir, read_pickle, set_last_folder, write_pickle
from svc.svc_host import ensure_ui_server

__version__ = "1.3.13"

try:
    from core import core_cst as cst
except Exception:
    cst = None  # type: ignore[assignment]

try:
    from core import core_xlc as xlc
except Exception:
    xlc = None  # type: ignore[assignment]

try:
    from core import core_stat
except Exception:
    core_stat = None  # type: ignore[assignment]

try:
    from core import core_w32 as _w32
except Exception:
    _w32 = None  # type: ignore[assignment]

logger = get_logger(__name__)
_ld_diag = get_diag_logger("hc_csv_tool.diag.csv_ld")


def _elapsed_ms(since: float) -> int:
    return max(0, int((time.perf_counter() - since) * 1000))


def _bring_excel_to_front(hwnd: int) -> None:
    """ネイティブ UI や別プロセス表示のあと Excel が背面に回る事象の緩和（ベストエフォート）。"""
    if not int(hwnd or 0) or _w32 is None or os.name != "nt":
        return
    try:
        _w32.bring_to_front(int(hwnd))
    except Exception:
        pass


def _ld_trace(fmt: str, *args: object) -> None:
    """診断ログ有効時のみ hc_csv_diag に [CSV_LD_TRACE] を出す。"""
    try:
        _ld_diag.info(fmt, *args)
    except Exception:
        pass


# 物理制約定数（hc_csv_ld と同様）
CHUNK_PCT_BASE: float = 0.1
MAX_CHUNK_LIMIT: int = 50000
MIN_CHUNK_LIMIT: int = 100
MAX_ROWS_PER_SHEET: int = 1000000

def _submit_request_dict(req_dict: dict[str, Any]) -> Path:
    """ui_server への要求を request_dir に pickle で投げる。"""
    req_dir = get_request_dir()
    req_dir.mkdir(parents=True, exist_ok=True)
    ts_ms = int(time.time() * 1000)
    req_path = req_dir / f"req_{ts_ms}_{os.getpid()}_{threading.get_ident()}.pkl"
    write_pickle(req_path, req_dict)
    return req_path


def _progress_path(sheet_id: str) -> Path:
    root = Path(get_ipc_root())
    d = root / "progress"
    d.mkdir(parents=True, exist_ok=True)
    return d / f"progress_ld_{sheet_id}.pkl"


def _progress_write(path: Path, obj: dict[str, Any]) -> None:
    try:
        write_pickle(path, obj)
    except Exception:
        pass


def _get_window_rect(hwnd: int) -> tuple[int, int, int, int] | None:
    """指定 HWND の現在の画面矩形 (left, top, right, bottom) を取得。送信時点の Excel 位置を UI に渡すため。"""
    if not int(hwnd or 0) or os.name != "nt":
        return None
    try:
        from ctypes import wintypes
        r = wintypes.RECT()
        if ctypes.windll.user32.GetWindowRect(int(hwnd), ctypes.byref(r)):
            return (int(r.left), int(r.top), int(r.right), int(r.bottom))
    except Exception:
        pass
    return None


def _normalize_cell_to_str(c: Any) -> str:
    """1セルを文字列に正規化する。None および float の nan は '' にし、それ以外は str(c)。"""
    if c is None:
        return ""
    if isinstance(c, float) and c != c:  # nan
        return ""
    return str(c)


def _normalize_row_to_strings(row: list[Any]) -> list[str]:
    """1行分のリストの各要素を文字列に正規化し、Excel へ文字列で書き込むためのリストを返す。"""
    return [_normalize_cell_to_str(c) for c in row]


def _submit_progress_ui(
    parent_hwnd: int, sheet_id: str, progress_path: Path, phase_total: int
) -> None:
    """ui_server へ進捗ウィンドウ表示を要求する（モデルレス）。送信時点の Excel 矩形を渡し、UI がその位置に合わせて中央配置する。"""
    try:
        excel_rect = _get_window_rect(int(parent_hwnd or 0))
        req_dir = get_request_dir()
        res_dir = Path(get_ipc_root()) / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        ts_ms = int(time.time() * 1000)
        result_path = str(res_dir / f"res_progress_ld_{ts_ms}_{os.getpid()}.pkl")
        req_dict = {
            "action": "progress",
            "progress_path": str(progress_path),
            "phase_total": int(phase_total),
            "excel_lock": False,
            # 枠だけ表示を避けるため Qt 描画を使用（WA_NativeWindow を付けない）
            "no_native_window": True,
        }
        if excel_rect is not None:
            req_dict["excel_rect"] = list(excel_rect)
        payload = {
            "parent_hwnd": int(parent_hwnd),
            "result_path": result_path,
            "ready_path": "",
            "sheet_id": str(sheet_id),
            "log_path": "",
            "action": "progress",
            "module": "ui_qt.ui_csv_ld",
            "req_dict": req_dict,
        }
        req_path = req_dir / f"req_{ts_ms}_{os.getpid()}.pkl"
        write_pickle(req_path, payload)
    except Exception:
        pass


def _set_performance_mode(app: Any, on: bool) -> None:
    """Excel の画面更新・計算を抑止/復帰させる（高速化）。on=True のとき画面更新を抑止。"""
    try:
        api = getattr(app, "api", None) or app
        # 抑止時は ScreenUpdating=False（on=True → False）。復帰時は True。
        api.ScreenUpdating = not on
        api.Calculation = -4135 if on else -4105  # xlCalculationManual / xlCalculationAutomatic
    except Exception:
        pass


def _safe_rename_sheet(sh: Any, name: str) -> None:
    """シート名を安全に変更する（衝突時は枝番付与）。"""
    try:
        sh.name = name
        return
    except Exception:
        pass
    for seq in range(1, 1000):
        try:
            sh.name = f"{name}_{seq}"
            return
        except Exception:
            continue


def _watch_ready(ready_path: str, sheet_id: str, t_load0: float) -> None:
    """READY_UI 監視（非同期）。受信したら notify_ui_ready で砂時計・WaitForm 解除。"""
    p = Path(ready_path)
    while True:
        if p.exists() and p.stat().st_size > 0:
            try:
                d = read_pickle(p)
            except Exception:
                time.sleep(0.05)
                continue
            st = str(d.get("status", "")).strip().upper()
            if st == "READY_UI":
                notify_ui_ready(cancel_reason="READY_UI")
                logger.info(
                    "[CSV_LD] phase=ready_ui elapsed_ms=%s sheet_id=%s",
                    _elapsed_ms(t_load0),
                    sheet_id or "",
                )
                _ld_trace(
                    "[CSV_LD_TRACE] phase=ready_ui sheet_id=%s elapsed_since_load_ms=%s wall_perf_s=%.6f",
                    sheet_id or "",
                    _elapsed_ms(t_load0),
                    time.perf_counter(),
                )
            return
        time.sleep(0.05)


def _message_yesno(hwnd: int, text: str, caption: str = "確認") -> bool:
    """Win32 MessageBox で Yes/No を表示し、Yes なら True。"""
    try:
        MB_YESNO = 0x04
        MB_ICONQUESTION = 0x20
        IDYES = 6
        u = ctypes.windll.user32  # type: ignore[attr-defined]
        r = u.MessageBoxW(int(hwnd), text, caption, MB_YESNO | MB_ICONQUESTION)
        return r == IDYES
    except Exception:
        return False


def _get_row_count_binary(str_path: str) -> int:
    cnt = 0
    with open(str_path, "rb") as f:
        for _ in f:
            cnt += 1
    return cnt


def _get_formatted_size(str_path: str) -> str:
    sz = os.path.getsize(str_path)
    if sz >= 1048576:
        return f"{sz / 1048576:.2f} MB"
    return f"{sz / 1024:.1f} KB"


def _detect_encoding(str_path: str) -> tuple[None, str, str]:
    try:
        with open(str_path, "rb") as f_h:
            if f_h.read(3) == b"\xef\xbb\xbf":
                return None, "UTF-8 (BOM)", "utf-8-sig"
        try:
            with open(str_path, "r", encoding="utf-8") as f_t:
                f_t.read(2048)
            return None, "UTF-8", "utf-8"
        except Exception:
            return None, "Shift-JIS", "cp932"
    except Exception:
        return None, "不明", "utf-8"


def _get_unique_base_name(book_p: Any, str_base: str) -> str:
    list_names = [s.name for s in book_p.sheets]
    if str_base not in list_names:
        return str_base
    seq = 1
    while True:
        target = f"{str_base}_{seq}"
        if target not in list_names:
            return target
        seq += 1


def _add_new_sheet_direct(book_p: Any, target_name: str) -> Any:
    try:
        return book_p.sheets.add(name=target_name)
    except Exception:
        for seq in range(1, 1000):
            try:
                return book_p.sheets.add(name=f"{target_name}_{seq}")
            except Exception:
                continue
    return None


def _is_sheet_empty(sh_ref: Any) -> bool:
    try:
        return (
            sh_ref.used_range.address == "$A$1"
            and sh_ref.range("A1").value is None
        )
    except Exception:
        return True


def _finalize_sheet_context(
    book: Any,
    sh: Any,
    fname: str,
    size_lbl: str,
    enc_lbl: str,
    accum_total: int,
    total_file: int,
    p_idx: int,
    p_total: int,
    is_split: bool,
    accum_in_sheet: int,
    base_name: str,
    sh_origin: Any,
) -> None:
    """共通サービス (core_stat) で HC_STATUS_INFO を保存。HC_NOTIFY_RETV は特別な時のみ設定する設計のため通常完了時は設定しない。"""
    if core_stat is None:
        return
    str_sh_name_v = f"シート名：{sh.name}"
    str_rows_v = f"行数：{accum_in_sheet:,} 行"
    str_fname_v = f"ファイル名：{fname}"
    str_size_v = f"容量：{size_lbl}"
    str_enc_v = f"文字コード：{enc_lbl}"
    str_data_v = f"データ：{max(0, total_file - 1):,} 行"
    str_total_st_v = f"総数(ヘッダ含む)：{total_file:,} 行"
    str_total_not_v = f"総数(ヘッダ含)：{total_file:,} 行"

    if is_split:
        info_status_body = (
            f"{str_sh_name_v} ｜ 分割：{p_idx}/{p_total} ｜ {str_rows_v} ┃ "
            f"{str_fname_v} ｜ {str_size_v} ｜ {str_enc_v} ｜ {str_data_v} ｜ {str_total_st_v}"
        )
    else:
        info_status_body = (
            f"{str_sh_name_v} ｜ {str_rows_v} ┃ "
            f"{str_fname_v} ｜ {str_size_v} ｜ {str_enc_v} ｜ {str_data_v} ｜ {str_total_st_v}"
        )

    core_stat.set_status_info(sh, info_status_body)

    try:
        book.app.api.StatusBar = f"CSV読込終了｜{info_status_body}"
    except Exception:
        pass


def _autofit_used_range(sht: Any, last_row: int, max_col: int, sheet_name: str) -> None:
    """列幅オートフィット。行数が core_cst.AUTOFIT_MAX_ROWS を超えるときはスキップ。"""
    if last_row <= 0 or max_col <= 0:
        return
    max_rows = int(getattr(cst, "AUTOFIT_MAX_ROWS", 100000) or 100000)

    def _do_autofit_range(rng: Any, label: str) -> bool:
        """rng に対して xlwings の columns.autofit または COM の Columns.AutoFit を実行。成功時 True。"""
        try:
            cols = getattr(rng, "columns", None)
            if cols is not None:
                autofit_fn = getattr(cols, "autofit", None) or getattr(cols, "AutoFit", None)
                if callable(autofit_fn):
                    autofit_fn()
                    return True
            api = getattr(rng, "api", None)
            if api is not None:
                api_cols = getattr(api, "Columns", None)
                if api_cols is not None:
                    api_autofit = getattr(api_cols, "AutoFit", None)
                    if callable(api_autofit):
                        api_autofit()
                        return True
        except Exception:
            pass
        return False

    if last_row <= max_rows:
        try:
            rng = sht.range((1, 1), (last_row, max_col))
            if _do_autofit_range(rng, "full_range"):
                return
        except Exception:
            pass
        try:
            ur = getattr(sht, "used_range", None)
            if ur is not None and _do_autofit_range(ur, "used_range"):
                return
        except Exception:
            pass
        return

    try:
        sht.activate()
    except Exception:
        pass
    try:
        book = getattr(sht, "book", None)
        app = getattr(book, "app", None) if book else None
        api = getattr(app, "api", None) if app else None
        aw = getattr(api, "ActiveWindow", None) if api else None
        if aw is not None:
            vis = getattr(aw, "VisibleRange", None)
            if vis is not None:
                parent = getattr(vis, "Parent", None)
                pname = getattr(parent, "Name", "") if parent else ""
                if pname == sheet_name:
                    cols = getattr(vis, "Columns", None)
                    if cols is not None:
                        autofit_fn = getattr(cols, "AutoFit", None) or getattr(cols, "Autofit", None)
                        if callable(autofit_fn):
                            autofit_fn()
                            return
    except Exception:
        pass
    try:
        rng = sht.range((1, 1), (1, max_col))
        _do_autofit_range(rng, "header_row_fallback")
    except Exception:
        pass


def _execute_jit_import(
    book: Any,
    sh_origin: Any,
    str_path: str,
    str_enc: str,
    val_total: int,
    val_size_lbl: str,
    str_enc_lbl: str,
    progress_path: Path | None,
    parent_hwnd: int,
    sheet_id: str,
    progress_ui_already_shown: bool = False,
) -> None:
    """境界行制御でチャンク読込し、各シートへ書き込み・統計保存。"""
    import pandas as pd

    str_fname = os.path.basename(str_path)
    str_base_resolved = _get_unique_base_name(book, os.path.splitext(str_fname)[0])

    val_accum_total = 0
    val_accum_in_sheet = 0
    val_sheets_total = math.ceil(val_total / MAX_ROWS_PER_SHEET)
    is_split_mode = val_total > MAX_ROWS_PER_SHEET
    curr_part_idx = 0
    sh_target = None
    max_col = 0  # 書き込み列数（オートフィット用）

    calc_chunk_v = int(val_total * CHUNK_PCT_BASE)
    chunk_size_v = min(MAX_CHUNK_LIMIT, max(MIN_CHUNK_LIMIT, calc_chunk_v))
    # 読込値は全て文字列として扱う（空セルは NaN にせず '' のまま）
    reader = pd.read_csv(
        str_path,
        encoding=str_enc,
        dtype=str,
        keep_default_na=False,
        na_filter=False,
        chunksize=chunk_size_v,
    )

    # 工程: 1=ファイル解析 2=Excel書き込み 3=列幅調整 4=完了
    total_steps = 4
    # 準備中を _do_load_csv で既に seq=0 で出している場合は phase1 を seq=1 から開始
    progress_seq = 1 if progress_ui_already_shown else 0
    if progress_path is not None:
        _progress_write(
            progress_path,
            {
                "status": "RUN",
                "phase_i": 1,
                "phase": "ファイル解析中",
                "done": 0,
                "total": val_total,
                "pct": 0,
                "current_file": str_fname,
                "seq": progress_seq,
            },
        )
        progress_seq += 1
        if progress_ui_already_shown:
            time.sleep(0.2)
        else:
            if parent_hwnd:
                _submit_progress_ui(parent_hwnd, sheet_id, progress_path, total_steps)
                time.sleep(core_env.progress_window_startup_wait_sec())

    # 一括書込みは表示停止で高速化（シート更新禁止→処理→更新再開）
    with xlc.suspend_sheet_updates(book):
        for df_chunk in reader:
            val_rows_in_chunk = len(df_chunk)
            # ヘッダのみ→1 データ行にした場合、新シート先頭で「列名行＋データ行」とすると
            # 値が列名と同一のため 2 行重複になるため、列名行は付けない
            header_only_as_one_body_row = False
            if val_rows_in_chunk == 0 and len(df_chunk.columns) > 0:
                df_chunk = pd.DataFrame([df_chunk.columns.tolist()], columns=df_chunk.columns)
                val_rows_in_chunk = 1
                header_only_as_one_body_row = True
            processed_chunk_idx = 0

            while processed_chunk_idx < val_rows_in_chunk:
                if sh_target is None or val_accum_in_sheet >= MAX_ROWS_PER_SHEET:
                    if sh_target is not None:
                        if xlc and max_col > 0:
                            xlc.clear_used_range_overflow(sh_target, val_accum_in_sheet, max_col)
                        _finalize_sheet_context(
                            book, sh_target, str_fname, val_size_lbl, str_enc_lbl,
                            val_accum_total, val_total, curr_part_idx, val_sheets_total,
                            is_split_mode, val_accum_in_sheet, str_base_resolved, sh_origin,
                        )
                    curr_part_idx += 1
                    str_target_name = f"{str_base_resolved}-{curr_part_idx}" if is_split_mode else str_base_resolved
                    if curr_part_idx == 1:
                        if _is_sheet_empty(sh_origin):
                            sh_target = sh_origin
                            _safe_rename_sheet(sh_target, str_target_name)
                        else:
                            sh_target = _add_new_sheet_direct(book, str_target_name)
                    else:
                        sh_target = _add_new_sheet_direct(book, str_target_name)
                        if xlc:
                            xlc.yield_to_excel()
                    val_accum_in_sheet = 0
                    is_new_sheet_boundary = True
                else:
                    is_new_sheet_boundary = False

                remaining_cap = MAX_ROWS_PER_SHEET - val_accum_in_sheet
                rows_to_write = min(val_rows_in_chunk - processed_chunk_idx, remaining_cap)
                df_slice = df_chunk.iloc[processed_chunk_idx : processed_chunk_idx + rows_to_write]

                if core_stat and xlc:
                    if core_stat.get_guid(sh_target) == "":
                        core_stat.set_guid(sh_target, xlc.create_guid_b64())
                        core_stat.set_prop(sh_target, core_stat.KEY_BOOK_NAME, book.name)

                if is_new_sheet_boundary:
                    if header_only_as_one_body_row:
                        matrix = df_slice.values.tolist()
                    else:
                        matrix = [df_slice.columns.tolist()] + df_slice.values.tolist()
                else:
                    matrix = df_slice.values.tolist()
                if matrix and len(matrix[0]) > max_col:
                    max_col = len(matrix[0])

                # 全セルを文字列に正規化して Excel へ文字列で書き込む（None/nan は '' に）
                matrix = [_normalize_row_to_strings(row) for row in matrix]

                start_row = val_accum_in_sheet + 1
                nrows, ncols = len(matrix), len(matrix[0]) if matrix else 0
                try:
                    if nrows and ncols:
                        sh_target.range((start_row, 1), (start_row + nrows - 1, ncols)).number_format = "@"
                except Exception:
                    pass
                if xlc:
                    xlc.write_chunk(sh_target, start_row, 1, matrix, None)

                val_accum_total += rows_to_write
                val_accum_in_sheet += len(matrix)
                processed_chunk_idx += rows_to_write

                if progress_path is not None:
                    pct = int((val_accum_total / val_total) * 100)
                    _progress_write(
                        progress_path,
                        {
                            "status": "RUN",
                            "phase_i": 2,
                            "phase": "Excelへ書き込み中",
                            "done": val_accum_total,
                            "total": val_total,
                            "pct": pct,
                            "current_file": str_fname,
                            "seq": progress_seq,
                        },
                    )
                    progress_seq += 1
                try:
                    if is_split_mode:
                        book.app.api.StatusBar = f"Excel書込中... {sh_target.name} [分割：{curr_part_idx}/{val_sheets_total}]"
                    else:
                        book.app.api.StatusBar = "Excel書込中..."
                except Exception:
                    pass
                if xlc:
                    xlc.yield_to_excel()

        # 更新停止のまま: 最終シートのクリア・確定 → オートフィット → その後 with を抜けて更新開始
        if sh_target is not None:
            if xlc and max_col > 0:
                xlc.clear_used_range_overflow(sh_target, val_accum_in_sheet, max_col)
            _finalize_sheet_context(
                book, sh_target, str_fname, val_size_lbl, str_enc_lbl,
                val_accum_total, val_total, curr_part_idx, val_sheets_total,
                is_split_mode, val_accum_in_sheet, str_base_resolved, sh_origin,
            )

        # 工程3: 列幅調整（オートフィット）。更新停止中に実行し、with 抜けで更新開始
        if progress_path is not None:
            _progress_write(
                progress_path,
                {
                    "status": "RUN",
                    "phase_i": 3,
                    "phase": "列幅調整中",
                    "done": val_total,
                    "total": val_total,
                    "pct": 99,
                    "current_file": str_fname,
                    "seq": progress_seq,
                },
            )
            progress_seq += 1
        if max_col > 0:
            total_parts = curr_part_idx
            for part_i in range(1, curr_part_idx + 1):
                if progress_path is not None and total_parts > 1:
                    _progress_write(
                        progress_path,
                        {
                            "status": "RUN",
                            "phase_i": 3,
                            "phase": f"列幅調整中 ({part_i}/{total_parts} シート)",
                            "done": val_total,
                            "total": val_total,
                            "pct": 99,
                            "current_file": str_fname,
                            "seq": progress_seq,
                        },
                    )
                    progress_seq += 1
                try:
                    sheet_name = str_base_resolved if part_i == 1 else f"{str_base_resolved}-{part_i}"
                    last_row_i = MAX_ROWS_PER_SHEET if part_i < curr_part_idx else val_accum_in_sheet
                    for s in book.sheets:
                        if getattr(s, "name", "") != sheet_name:
                            continue
                        _autofit_used_range(s, last_row_i, max_col, sheet_name)
                        break
                except Exception:
                    pass
                if xlc:
                    xlc.yield_to_excel()


    if progress_path is not None:
        logger.info("[CSV_LD] 完了 行数=%s シート数=%s", val_total, curr_part_idx)
        data_rows = max(0, val_total - 1)
        done_detail_text = (
            f"シート名：{str_base_resolved}\n"
            f"ファイル名：{str_fname or 'loaded.csv'}\n"
            f"容量：{val_size_lbl}\n"
            f"文字コード：{str_enc_lbl}\n"
            f"データ：{data_rows} 行\n"
            f"総数(ヘッダ含)：{val_total} 行"
        )
        _progress_write(
            progress_path,
            {
                "status": "DONE",
                "phase_i": total_steps,
                "phase": "完了",
                "done": val_total,
                "total": val_total,
                "pct": 100,
                "current_file": str_fname,
                "show_done_dialog": True,
                "done_items": [
                    {"no": 1, "name": str_fname or "loaded.csv", "rows": val_total},
                ],
                "done_detail_text": done_detail_text,
                "seq": progress_seq,
            },
        )


def _do_load_csv(
    book: Any,
    sheet_id: str,
    str_csv_path: str,
    parent_hwnd: int,
    progress_ui_already_shown: bool = False,
    t_load0: float = 0.0,
) -> None:
    """ファイルパス確定後の読込実行（分割確認・進捗・Excel書込）。"""
    t_do0 = time.perf_counter()
    logger.info(
        "[CSV_LD] phase=do_load_enter file=%s sheet_id=%s elapsed_since_load_ms=%s",
        os.path.basename(str_csv_path),
        sheet_id or "",
        _elapsed_ms(t_load0) if t_load0 else 0,
    )
    _ld_trace(
        "[CSV_LD_TRACE] phase=do_load_enter file=%s sheet_id=%s wall_perf_s=%.6f",
        os.path.basename(str_csv_path),
        sheet_id or "",
        time.perf_counter(),
    )
    progress_path = _progress_path(sheet_id)
    # 進捗を早期表示（ui_server が既に表示済みでない場合のみ。csv_ld ファイル選択経路では ui_server が即表示するためスキップ）
    if not progress_ui_already_shown:
        _progress_write(
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
        if parent_hwnd:
            _submit_progress_ui(parent_hwnd, sheet_id, progress_path, 4)

    if xlc is None:
        logger.error("[CSV_LD] core_xlc not available")
        _progress_write(progress_path, {"status": "DONE", "seq": 1})
        _bring_excel_to_front(parent_hwnd)
        return
    sh_origin = xlc.find_sheet_by_guid(book, sheet_id)
    if sh_origin is None:
        logger.error("[CSV_LD] 対象シートなし GUID=%s", sheet_id)
        _progress_write(progress_path, {"status": "DONE", "seq": 1})
        _bring_excel_to_front(parent_hwnd)
        return

    t_rows0 = time.perf_counter()
    val_total = _get_row_count_binary(str_csv_path)
    str_readable_size = _get_formatted_size(str_csv_path)
    row_count_ms = _elapsed_ms(t_rows0)
    logger.info(
        "[CSV_LD] 読込 行数=%s 容量=%s row_count_ms=%s",
        f"{val_total:,}",
        str_readable_size,
        row_count_ms,
    )
    _ld_trace(
        "[CSV_LD_TRACE] phase=row_count_done rows=%s row_count_ms=%s wall_perf_s=%.6f",
        val_total,
        row_count_ms,
        time.perf_counter(),
    )

    if val_total > MAX_ROWS_PER_SHEET:
        val_sheets_total = math.ceil(val_total / MAX_ROWS_PER_SHEET)
        msg = (
            f"総行数 {val_total:,} 行を検知しました。\n"
            f"全 {val_sheets_total} 枚のシートに分割して読み込みますか？"
        )
        if not _message_yesno(parent_hwnd, msg, "CSV読込"):
            logger.info("[CSV_LD] 分割読込ユーザー拒否")
            _progress_write(progress_path, {"status": "DONE", "seq": 1})
            _bring_excel_to_front(parent_hwnd)
            return

    _set_performance_mode(book.app, True)
    t_jit0 = time.perf_counter()
    try:
        _, str_enc_label, str_encoding_name = _detect_encoding(str_csv_path)
        _execute_jit_import(
            book=book,
            sh_origin=sh_origin,
            str_path=str_csv_path,
            str_enc=str_encoding_name,
            val_total=val_total,
            val_size_lbl=str_readable_size,
            str_enc_lbl=str_enc_label,
            progress_path=progress_path,
            parent_hwnd=parent_hwnd,
            sheet_id=sheet_id,
            progress_ui_already_shown=progress_ui_already_shown,
        )
        logger.info(
            "[CSV_LD] phase=jit_import_done jit_ms=%s do_load_ms=%s since_load_ms=%s",
            _elapsed_ms(t_jit0),
            _elapsed_ms(t_do0),
            _elapsed_ms(t_load0) if t_load0 else 0,
        )
        _ld_trace(
            "[CSV_LD_TRACE] phase=jit_import_done jit_ms=%s do_load_ms=%s wall_perf_s=%.6f",
            _elapsed_ms(t_jit0),
            _elapsed_ms(t_do0),
            time.perf_counter(),
        )
    except Exception as ex_fatal:
        logger.error("[CSV_LD] 致命的エラー: %s", ex_fatal, exc_info=True)
        try:
            _progress_write(progress_path, {"status": "DONE", "seq": 1})
        except Exception:
            pass
        try:
            book.app.api.StatusBar = f"ERROR: CSV読込不全 | {ex_fatal}"
        except Exception:
            pass
    finally:
        _set_performance_mode(book.app, False)
        if xlc:
            xlc.yield_to_excel()


def _watch_result(
    result_path: str,
    book: Any,
    sheet_id: str,
    parent_hwnd: int,
    t_load0: float,
    pick_anchor: list[float | None] | None = None,
) -> None:
    """UI結果を待ち、OK かつ path があれば _do_load_csv を実行。

    pick_anchor: 長さ 1 のリストを渡すと、ファイル確定直前に perf_counter を格納し
    load_csv 終端で pick_to_done_ms を算出する（区間 B 計測用）。
    """
    p = Path(result_path)
    while True:
        if p.exists() and p.stat().st_size > 0:
            try:
                d = read_pickle(p)
            except Exception:
                time.sleep(0.05)
                continue
            status = str(d.get("status", "")).strip().upper()
            path = str(d.get("path", "")).strip()
            if status == "OK" and path:
                logger.info(
                    "[CSV_LD] phase=result_ok file=%s elapsed_ms=%s sheet_id=%s",
                    os.path.basename(path),
                    _elapsed_ms(t_load0),
                    sheet_id or "",
                )
                _ld_trace(
                    "[CSV_LD_TRACE] phase=result_ok path=%s elapsed_since_load_ms=%s wall_perf_s=%.6f",
                    path,
                    _elapsed_ms(t_load0),
                    time.perf_counter(),
                )
                if pick_anchor is not None:
                    # 呼び出し側は [None] を渡すこと（空リスト不可）
                    pick_anchor[0] = time.perf_counter()
                logger.info(
                    "[CSV_LD] phase=pick_confirmed file=%s elapsed_since_load_ms=%s sheet_id=%s",
                    os.path.basename(path),
                    _elapsed_ms(t_load0),
                    sheet_id or "",
                )
                try:
                    set_last_folder(os.path.dirname(path))
                except Exception:
                    pass
                _bring_excel_to_front(parent_hwnd)
                _do_load_csv(
                    book,
                    sheet_id,
                    path,
                    parent_hwnd,
                    progress_ui_already_shown=True,
                    t_load0=t_load0,
                )
            else:
                logger.info(
                    "[CSV_LD] phase=result_cancel elapsed_ms=%s sheet_id=%s",
                    _elapsed_ms(t_load0),
                    sheet_id or "",
                )
                _ld_trace(
                    "[CSV_LD_TRACE] phase=result_cancel status=%s elapsed_since_load_ms=%s wall_perf_s=%.6f",
                    status,
                    _elapsed_ms(t_load0),
                    time.perf_counter(),
                )
            _bring_excel_to_front(parent_hwnd)
            return
        time.sleep(0.05)


def load_csv(book: Any, sheet_id: str = "") -> None:
    """
    CSV読込のエントリ。Qt UIサーバでファイル選択 → 結果取得後に読込実行。
    """
    t_load0 = time.perf_counter()
    logger.info(
        "[CSV_LD] 開始 sheet_id=%s svc_pid=%s",
        sheet_id or "",
        os.getpid(),
    )
    _ld_trace(
        "[CSV_LD_TRACE] phase=enter sheet_id=%s svc_pid=%s wall_perf_s=%.6f",
        sheet_id or "",
        os.getpid(),
        time.perf_counter(),
    )

    if book is None:
        logger.error("[CSV_LD] book=None のため中断（WaitForm 解除を試行）")
        try:
            notify_wait_form_ready()
        except Exception:
            pass
        return

    ensure_ui_server()
    logger.info(
        "[CSV_LD] phase=after_ensure_ui_server elapsed_ms=%s",
        _elapsed_ms(t_load0),
    )
    _ld_trace(
        "[CSV_LD_TRACE] phase=after_ensure_ui_server elapsed_ms=%s wall_perf_s=%.6f",
        _elapsed_ms(t_load0),
        time.perf_counter(),
    )

    parent_hwnd = 0
    try:
        parent_hwnd = int(getattr(book.app, "hwnd", 0))
        core_env.set_excel_hwnd_for_spawn(parent_hwnd)
    except Exception as e:
        logger.warning("[CSV_LD] Excel HWND 取得失敗: %s", e)
    if not parent_hwnd:
        logger.warning("[CSV_LD] parent_hwnd=0")

    ipc_root = Path(get_ipc_root())
    res_path = str(ipc_root / "results" / f"res_ld_{sheet_id}_{int(time.time()*1000)}.pkl")
    ready_path = str(ipc_root / "ready" / f"ready_ld_{sheet_id}_{int(time.time()*1000)}.pkl")
    Path(res_path).parent.mkdir(parents=True, exist_ok=True)
    Path(ready_path).parent.mkdir(parents=True, exist_ok=True)
    try:
        Path(res_path).unlink(missing_ok=True)
        Path(ready_path).unlink(missing_ok=True)
    except Exception:
        pass

    req_dict = {
        "parent_hwnd": parent_hwnd,
        "result_path": res_path,
        "ready_path": ready_path,
        "sheet_id": str(sheet_id) if sheet_id else "_",
        "log_path": "",
        "action": "csv_ld",
        "module": "ui_qt.ui_csv_ld",
        "initial_dir": get_last_folder(),
    }

    req_path = _submit_request_dict(req_dict)
    logger.info(
        "[CSV_LD] ui_ipc ok req=%s sheet_id=%s hwnd=%s",
        req_path.name,
        sheet_id or "",
        parent_hwnd,
    )
    _ld_trace(
        "[CSV_LD_TRACE] ui_ipc ok req=%s sheet_id=%s hwnd=%s elapsed_ms=%s wall_perf_s=%.6f",
        req_path.name,
        sheet_id or "",
        parent_hwnd,
        _elapsed_ms(t_load0),
        time.perf_counter(),
    )

    th_ready = threading.Thread(
        target=_watch_ready,
        args=(ready_path, str(sheet_id or ""), t_load0),
        name=f"ready_watch_ld_{sheet_id}",
        daemon=True,
    )
    th_ready.start()

    pick_anchor: list[float | None] = [None]
    _watch_result(
        res_path, book, str(sheet_id or ""), parent_hwnd, t_load0, pick_anchor
    )
    t_end = time.perf_counter()
    pick_t0 = pick_anchor[0]
    pick_to_done_ms = (
        max(0, int((t_end - pick_t0) * 1000)) if pick_t0 is not None else -1
    )
    logger.info(
        "[CSV_LD] phase=load_csv_flow_done elapsed_since_load_ms=%s pick_to_done_ms=%s sheet_id=%s",
        _elapsed_ms(t_load0),
        pick_to_done_ms,
        sheet_id or "",
    )
    _ld_trace(
        "[CSV_LD_TRACE] phase=load_csv_flow_done elapsed_since_load_ms=%s pick_to_done_ms=%s wall_perf_s=%.6f",
        _elapsed_ms(t_load0),
        pick_to_done_ms,
        t_end,
    )
