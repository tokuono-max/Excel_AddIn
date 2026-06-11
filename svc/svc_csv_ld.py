# -*- coding: utf-8 -*-
"""
Python: 3.12+
Module: svc/svc_csv_ld.py
Created: 2026-03-05
Updated: 2026-06-06
Version: 1.3.26
Purpose:
  CSV読込（Qt UIサーバ方式 / 2プロセス分離）。
  期待フロー: CSVファイル選択 → 進捗画面表示（準備中→ファイル解析中…）→ CSV読込 → Excelシート出力 → セルオートフィット → 進捗画面閉じる → 完了通知。
  進捗は 0=準備中 1=ファイル解析 2=Excel書き込み 3=列幅調整 4=完了 の5段階で表示。無表示1秒未満のため ui_server がファイル選択OK直後に進捗を即表示する経路では progress_ui_already_shown で二重依頼を避ける。

History (latest 3):
  - 1.3.26 (2026-06-06) ハング緩和: DONE を ScreenUpdating 復帰前に書込・1秒待機。suspend restore_on_exit=False。
  - 1.3.25 (2026-06-06) 進捗表示中は excel_lock=True（読込中 Excel 操作無効）。完了時 teardown で解除。
  - 1.3.24 (2026-06-04) 読込終了時に EnableEvents=True を保証（core.excel_perf_mode / restore_excel_host_after_operation）。シート切替でステータスバー復帰。
  - 1.3.23 (2026-06-04) シート名: ファイル名を Excel 最大 31 文字・禁止文字除去に整形。分割時は -N 分を確保して切り詰め。
  - 1.3.22 (2026-06-04) 進捗更新を細かく: 書込み notify 5k 行・stride/間隔/poll/creep 既定を強化。HC_CSV_LD_PROGRESS_WRITE_NOTIFY_ROWS。
  - 1.3.21 (2026-06-04) 進捗比率: バー用 total をデータ行数（ファイル行−ヘッダ）に統一。pct は done/total と一致。
  - 1.3.20 (2026-06-04) ファイル選択後の COM 切れ対策: book 再取得・HWND からシート解決。失敗時は進捗 ERROR（即 DONE 閉じを避ける）。
  - 1.3.19 (2026-06-03) 砂時計: ファイル確定後〜処理完了まで ON＋再武装。進捗: poll 60ms / creep 5 / 書込み中は時間ベース IPC。
  - 1.3.18 (2026-06-03) 大容量: read chunk / write_step 既定強化。文字列書込は範囲 @ + 素の値（' 全セル変換省略）。csv_read_wait_ms 計測修正。読込中砂時計。
  - 1.3.17 (2026-06-03) jit_breakdown 計測ログ（pandas_read / matrix_tolist / excel_write / sheet_boundary / finalize / autofit ms）。
  - 1.3.16 (2026-05-29) 進捗滑らか化: Excel 書込み 50k 刻み + 時間ベース IPC・UI バー補間・poll 100ms。
  - 1.3.15 (2026-05-29) 進捗 IPC 既定間隔を 50k 行に戻し UI 更新を高頻度化（Excel 書込み最適化は維持）。
  - 1.3.14 (2026-05-29) Excel 書込み高速化: チャンク 100k 既定・書式 @ をシート単位一括・進捗 IPC 間引き・EnableEvents 抑止・ループ内 yield 削減。
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
from core.excel_host_restore import restore_excel_host_after_operation
from core.core_progress_wait import wait_after_progress_done
from core.excel_perf_mode import set_excel_performance_mode
from core.core_cursor import notify_ui_ready, notify_wait_form_ready
from ui_qt.ipc_file import get_ipc_root, get_last_folder, get_request_dir, read_pickle, set_last_folder, write_pickle
from svc.svc_host import ensure_ui_server

__version__ = "1.3.26"


def _log_jit_breakdown(
    *,
    csv_read_wait_ms: int,
    matrix_tolist_ms: int,
    excel_write_ms: int,
    sheet_boundary_ms: int,
    finalize_ms: int,
    autofit_ms: int,
    jit_total_ms: int,
    fast_text_write: bool,
) -> None:
    """jit 区間の内訳を運用・診断ログへ出力（15秒目標のボトルネック切り分け用）。"""
    logger.info(
        "[CSV_LD] phase=jit_breakdown csv_read_wait_ms=%s matrix_tolist_ms=%s "
        "excel_write_ms=%s sheet_boundary_ms=%s finalize_ms=%s autofit_ms=%s "
        "jit_total_ms=%s fast_text_write=%s",
        csv_read_wait_ms,
        matrix_tolist_ms,
        excel_write_ms,
        sheet_boundary_ms,
        finalize_ms,
        autofit_ms,
        jit_total_ms,
        fast_text_write,
    )
    _ld_trace(
        "[CSV_LD_TRACE] phase=jit_breakdown csv_read_wait_ms=%s matrix_tolist_ms=%s "
        "excel_write_ms=%s sheet_boundary_ms=%s finalize_ms=%s autofit_ms=%s "
        "jit_total_ms=%s fast_text_write=%s",
        csv_read_wait_ms,
        matrix_tolist_ms,
        excel_write_ms,
        sheet_boundary_ms,
        finalize_ms,
        autofit_ms,
        jit_total_ms,
        fast_text_write,
    )

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
MAX_CHUNK_LIMIT: int = 100000
MIN_CHUNK_LIMIT: int = 100
MAX_ROWS_PER_SHEET: int = 1000000
DEFAULT_PROGRESS_STRIDE_ROWS: int = 5000
DEFAULT_PROGRESS_POLL_MS: int = 40
DEFAULT_PROGRESS_BAR_CREEP_PCT: int = 2
DEFAULT_EXCEL_WRITE_STEP_ROWS: int = 50000
DEFAULT_PROGRESS_MIN_INTERVAL_SEC: float = 0.12
DEFAULT_PROGRESS_WRITE_NOTIFY_ROWS: int = 5000
LARGE_FILE_PROGRESS_MIN_INTERVAL_SEC: float = 0.08
LARGE_FILE_PROGRESS_STRIDE_ROWS: int = 2000
LARGE_FILE_PROGRESS_WRITE_NOTIFY_ROWS: int = 5000
LARGE_FILE_ROWS_THRESHOLD: int = 500_000
LARGE_FILE_READ_CHUNK_ROWS: int = 100_000
LARGE_FILE_WRITE_STEP_ROWS: int = 200_000

def _csv_ld_legacy_text_write() -> bool:
    """HC_CSV_LD_LEGACY_TEXT_WRITE=1 のとき従来の text_mode（' 付与 + 塊ごと @）。"""
    return core_env.truthy(core_env.get("HC_CSV_LD_LEGACY_TEXT_WRITE"))


def _max_chunk_rows() -> int:
    """HC_CSV_LD_MAX_CHUNK_ROWS で上書き可能（既定 MAX_CHUNK_LIMIT）。"""
    raw = core_env.get("HC_CSV_LD_MAX_CHUNK_ROWS")
    if raw is not None:
        try:
            return max(MIN_CHUNK_LIMIT, min(500000, int(str(raw).strip())))
        except ValueError:
            pass
    return MAX_CHUNK_LIMIT


def resolve_read_chunk_size(val_total: int) -> int:
    """pandas read_csv の chunksize。行数の 10% を cap 内に収める。大容量は下限を引き上げる。"""
    vt = max(0, int(val_total))
    calc_chunk_v = int(vt * CHUNK_PCT_BASE)
    cap = _max_chunk_rows()
    chunk_target = calc_chunk_v
    if vt >= LARGE_FILE_ROWS_THRESHOLD:
        # 大容量は cap いっぱいまでチャンクを大きく（HC_CSV_LD_MAX_CHUNK_ROWS が効く）
        chunk_target = max(chunk_target, cap)
    return min(cap, max(MIN_CHUNK_LIMIT, chunk_target))


def resolve_excel_write_step_rows(val_total: int) -> int:
    """write_chunk の COM 分割行数（pandas 読込 chunk より小さくし進捗を細かく更新）。"""
    raw = core_env.get("HC_CSV_LD_EXCEL_WRITE_ROWS")
    if raw is not None:
        try:
            return max(5000, min(200000, int(str(raw).strip())))
        except ValueError:
            pass
    vt = max(1, int(val_total))
    if vt >= LARGE_FILE_ROWS_THRESHOLD:
        return max(5000, min(200000, LARGE_FILE_WRITE_STEP_ROWS))
    return max(5000, min(DEFAULT_EXCEL_WRITE_STEP_ROWS, vt))


def resolve_progress_min_interval_sec(val_total: int = 0) -> float:
    """進捗 IPC の最短更新間隔（秒）。HC_CSV_LD_PROGRESS_MIN_INTERVAL_SEC。"""
    raw = core_env.get("HC_CSV_LD_PROGRESS_MIN_INTERVAL_SEC")
    if raw is not None:
        try:
            return max(0.05, min(2.0, float(str(raw).strip())))
        except ValueError:
            pass
    if int(val_total) >= LARGE_FILE_ROWS_THRESHOLD:
        return LARGE_FILE_PROGRESS_MIN_INTERVAL_SEC
    return DEFAULT_PROGRESS_MIN_INTERVAL_SEC


def resolve_progress_stride_rows(val_total: int) -> int:
    """進捗 pickle 更新の行間隔。HC_CSV_LD_PROGRESS_STRIDE_ROWS で上書き可能。"""
    raw = core_env.get("HC_CSV_LD_PROGRESS_STRIDE_ROWS")
    if raw is not None:
        try:
            return max(1000, int(str(raw).strip()))
        except ValueError:
            pass
    vt = max(1, int(val_total))
    cap_stride = LARGE_FILE_PROGRESS_STRIDE_ROWS if vt >= LARGE_FILE_ROWS_THRESHOLD else DEFAULT_PROGRESS_STRIDE_ROWS
    return max(1000, min(cap_stride, vt))


def resolve_progress_write_notify_rows(val_total: int) -> int:
    """write_chunk 内の COM 分割（progress_cb 頻度）。HC_CSV_LD_EXCEL_WRITE_ROWS より小さくできる。"""
    raw = core_env.get("HC_CSV_LD_PROGRESS_WRITE_NOTIFY_ROWS")
    if raw is not None:
        try:
            return max(500, min(50000, int(str(raw).strip())))
        except ValueError:
            pass
    vt = max(1, int(val_total))
    if vt >= LARGE_FILE_ROWS_THRESHOLD:
        return max(500, min(LARGE_FILE_PROGRESS_WRITE_NOTIFY_ROWS, vt))
    return max(500, min(DEFAULT_PROGRESS_WRITE_NOTIFY_ROWS, vt))


def resolve_progress_poll_ms() -> int:
    """進捗 UI のポーリング間隔（ms）。HC_CSV_LD_PROGRESS_POLL_MS。"""
    raw = core_env.get("HC_CSV_LD_PROGRESS_POLL_MS")
    if raw is not None:
        try:
            return max(50, min(500, int(str(raw).strip())))
        except ValueError:
            pass
    return DEFAULT_PROGRESS_POLL_MS


def resolve_progress_bar_creep_pct() -> int:
    """進捗バーの 1 ティックあたりの繰り上げ幅（%）。HC_CSV_LD_PROGRESS_BAR_CREEP_PCT。"""
    raw = core_env.get("HC_CSV_LD_PROGRESS_BAR_CREEP_PCT")
    if raw is not None:
        try:
            return max(0, min(10, int(str(raw).strip())))
        except ValueError:
            pass
    return DEFAULT_PROGRESS_BAR_CREEP_PCT


def resolve_progress_row_total(val_total: int) -> int:
    """進捗バー・done/total の分母。ファイル行数の先頭 1 行は CSV ヘッダ想定。"""
    vt = max(0, int(val_total))
    if vt <= 1:
        return max(1, vt)
    return vt - 1


def calc_progress_pct(phase_i: int, done: int, row_total: int) -> int:
    """工程と done/row_total から進捗バー用 pct（0〜100）を算出。"""
    pi = int(phase_i)
    rt = max(1, int(row_total))
    dn = max(0, int(done))
    if pi <= 1:
        return 0
    if pi == 2:
        return min(99, int(dn * 100 / rt))
    if pi == 3:
        return 99
    return min(100, int(dn * 100 / rt))


def should_emit_progress_update(
    val_accum_total: int,
    val_total: int,
    last_progress_at: int,
    *,
    stride: int,
    last_progress_mono: float = 0.0,
    min_interval_sec: float = DEFAULT_PROGRESS_MIN_INTERVAL_SEC,
) -> bool:
    """進捗 IPC を書くべきタイミング（行数 stride または最短時間経過）。"""
    now = time.monotonic()
    if last_progress_mono > 0 and (now - last_progress_mono) >= max(0.1, float(min_interval_sec)):
        return True
    if val_accum_total <= 0:
        return True
    if val_accum_total >= val_total:
        return True
    if last_progress_at <= 0:
        return True
    return (val_accum_total - last_progress_at) >= max(1, int(stride))


def _apply_sheet_text_number_format(sh: Any, last_row: int, max_col: int) -> None:
    """データ範囲全体を文字列書式 (@) に一括設定（chunk 毎 COM 呼び出しを避ける）。"""
    if last_row <= 0 or max_col <= 0:
        return
    try:
        sh.range((1, 1), (last_row, max_col)).number_format = "@"
    except Exception:
        pass


def _apply_range_text_number_format(
    sh: Any, start_row: int, nrows: int, max_col: int
) -> None:
    """書込み直前に当該矩形だけ @ にする（text_mode の塊ごと @ を避ける）。"""
    if nrows <= 0 or max_col <= 0 or start_row < 1:
        return
    try:
        sh.range((start_row, 1), (start_row + nrows - 1, max_col)).number_format = "@"
    except Exception:
        pass

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


def _capture_book_attach_keys(book: Any) -> tuple[int, str, str]:
    """svc_server 再 attach 用に HWND / fullname / name を保存する。"""
    excel_hwnd = 0
    book_fullname = ""
    book_name = ""
    if book is None:
        return excel_hwnd, book_fullname, book_name
    try:
        excel_hwnd = int(getattr(book.app, "hwnd", 0) or 0)
    except Exception:
        excel_hwnd = 0
    try:
        book_fullname = str(getattr(book, "fullname", "") or "")
    except Exception:
        book_fullname = ""
    try:
        book_name = str(getattr(book, "name", "") or "")
    except Exception:
        book_name = ""
    return excel_hwnd, book_fullname, book_name


def _resolve_book_and_sheet(
    book: Any,
    sheet_id: str,
    parent_hwnd: int,
    attach_keys: tuple[int, str, str] | None = None,
) -> tuple[Any | None, Any | None]:
    """対象シートを解決する。ファイルダイアログ待ち後の COM 切れ時は book を再取得する。"""
    if xlc is None or not sheet_id:
        return book, None

    sh = xlc.find_sheet_by_guid(book, sheet_id)
    if sh is not None:
        return book, sh

    excel_hwnd, book_fullname, book_name = attach_keys or (0, "", "")
    if not excel_hwnd and not book_fullname and not book_name:
        excel_hwnd, book_fullname, book_name = _capture_book_attach_keys(book)
    hwnd_for_attach = int(excel_hwnd or parent_hwnd or 0)

    try:
        from svc.svc_server import _attach_book

        book2 = _attach_book(
            excel_hwnd=hwnd_for_attach,
            book_fullname=book_fullname,
            book_name=book_name,
        )
        sh2 = xlc.find_sheet_by_guid(book2, sheet_id)
        if sh2 is not None:
            logger.info(
                "[CSV_LD] phase=book_reattached sheet_id=%s hwnd=%s",
                sheet_id,
                hwnd_for_attach,
            )
            return book2, sh2
    except Exception as ex:
        logger.warning(
            "[CSV_LD] book reattach failed sheet_id=%s hwnd=%s ex=%r",
            sheet_id,
            hwnd_for_attach,
            ex,
        )

    if parent_hwnd:
        ctx = xlc.get_excel_context_from_hwnd(int(parent_hwnd), sheet_id)
        if ctx is not None:
            _app, book3, sh3, _hwnd = ctx
            if sh3 is not None:
                logger.info(
                    "[CSV_LD] phase=book_resolved_via_hwnd sheet_id=%s hwnd=%s",
                    sheet_id,
                    parent_hwnd,
                )
                return book3, sh3

    return book, None


def _progress_write_sheet_error(progress_path: Path, sheet_id: str, seq: int = 1) -> None:
    """COM 切れ等でシート解決できないとき。即 DONE ではなく ERROR で理由を表示する。"""
    _progress_write(
        progress_path,
        {
            "status": "ERROR",
            "phase": "CSV読込を開始できません",
            "detail": (
                "対象シートに接続できませんでした。\n"
                f"（GUID={sheet_id}）\n"
                "Excel を前面にしたうえで、もう一度 CSV 読込を実行してください。"
            ),
            "seq": int(seq),
        },
    )


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
            "excel_lock": True,
            "progress_poll_ms": resolve_progress_poll_ms(),
            "progress_bar_creep_pct": resolve_progress_bar_creep_pct(),
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


def _safe_rename_sheet(sh: Any, name: str) -> None:
    """シート名を安全に変更する（31 文字・禁止文字整形、衝突時は枝番付与）。"""
    if xlc is None:
        return
    safe = xlc.sanitize_excel_sheet_name(name, fallback="CSV")
    try:
        sh.name = safe
        return
    except Exception:
        pass
    for seq in range(1, 1000):
        suf = f"_{seq}"
        cand = (safe[: max(1, xlc.EXCEL_SHEET_NAME_MAX_LEN - len(suf))] + suf)[
            : xlc.EXCEL_SHEET_NAME_MAX_LEN
        ]
        try:
            sh.name = cand
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
    """ファイル名（拡張子除く）から、ブック内で使えるユニークなシート名ベースを返す。"""
    if xlc is None:
        return (str_base or "CSV")[:31]
    names = {str(s.name) for s in book_p.sheets}
    return xlc.unique_excel_sheet_name_in_names(names, str_base, fallback="CSV")


def _csv_ld_target_sheet_name(str_base_resolved: str, part_idx: int, is_split_mode: bool) -> str:
    """分割時は base-N（N 分を確保して 31 文字以内）、単一シート時は base。"""
    if xlc is None:
        raw = f"{str_base_resolved}-{part_idx}" if is_split_mode else str_base_resolved
        return raw[:31]
    if is_split_mode:
        return xlc.excel_sheet_name_for_split_part(str_base_resolved, part_idx)
    return str_base_resolved


def _add_new_sheet_direct(book_p: Any, target_name: str) -> Any:
    if xlc is None:
        safe = (target_name or "CSV")[:31]
    else:
        safe = xlc.sanitize_excel_sheet_name(target_name, fallback="CSV")
    try:
        return book_p.sheets.add(name=safe)
    except Exception:
        for seq in range(1, 1000):
            suf = f"_{seq}"
            cand = (safe[: max(1, xlc.EXCEL_SHEET_NAME_MAX_LEN - len(suf))] + suf)[
                : xlc.EXCEL_SHEET_NAME_MAX_LEN
            ] if xlc else f"{safe[:28]}_{seq}"[:31]
            try:
                return book_p.sheets.add(name=cand)
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

    t_jit_body0 = time.perf_counter()
    ms_csv_read_wait = 0
    ms_matrix_tolist = 0
    ms_excel_write = 0
    use_fast_text_write = not _csv_ld_legacy_text_write()
    ms_sheet_boundary = 0
    ms_finalize = 0
    ms_autofit = 0

    str_fname = os.path.basename(str_path)
    str_base_resolved = _get_unique_base_name(book, os.path.splitext(str_fname)[0])

    val_accum_total = 0
    val_accum_in_sheet = 0
    val_sheets_total = math.ceil(val_total / MAX_ROWS_PER_SHEET)
    is_split_mode = val_total > MAX_ROWS_PER_SHEET
    curr_part_idx = 0
    sh_target = None
    max_col = 0  # 書き込み列数（オートフィット用）

    chunk_size_v = resolve_read_chunk_size(val_total)
    write_step_v = resolve_excel_write_step_rows(val_total)
    progress_stride = resolve_progress_stride_rows(val_total)
    progress_min_sec = resolve_progress_min_interval_sec(val_total)
    progress_poll_ms = resolve_progress_poll_ms()
    progress_creep_pct = resolve_progress_bar_creep_pct()
    progress_notify_rows = resolve_progress_write_notify_rows(val_total)
    prog_row_total = resolve_progress_row_total(val_total)
    last_progress_rows = 0
    last_progress_mono = 0.0
    logger.info(
        "[CSV_LD] phase=jit_import_config rows=%s prog_row_total=%s chunk_size=%s write_step=%s "
        "progress_stride=%s progress_min_sec=%s progress_poll_ms=%s progress_creep_pct=%s "
        "progress_notify_rows=%s",
        val_total,
        prog_row_total,
        chunk_size_v,
        write_step_v,
        progress_stride,
        progress_min_sec,
        progress_poll_ms,
        progress_creep_pct,
        progress_notify_rows,
    )
    _ld_trace(
        "[CSV_LD_TRACE] jit_import_config rows=%s prog_row_total=%s chunk_size=%s write_step=%s "
        "progress_stride=%s progress_min_sec=%s progress_poll_ms=%s progress_creep_pct=%s "
        "progress_notify_rows=%s",
        val_total,
        prog_row_total,
        chunk_size_v,
        write_step_v,
        progress_stride,
        progress_min_sec,
        progress_poll_ms,
        progress_creep_pct,
        progress_notify_rows,
    )

    def _emit_write_progress(done_rows: int, *, time_only: bool = False) -> None:
        nonlocal progress_seq, last_progress_rows, last_progress_mono
        if progress_path is None:
            return
        emit_ok = False
        if time_only:
            now = time.monotonic()
            if (
                last_progress_mono <= 0
                or (now - last_progress_mono) >= max(0.05, float(progress_min_sec))
                or done_rows >= prog_row_total
            ):
                emit_ok = True
        elif should_emit_progress_update(
            done_rows,
            prog_row_total,
            last_progress_rows,
            stride=progress_stride,
            last_progress_mono=last_progress_mono,
            min_interval_sec=progress_min_sec,
        ):
            emit_ok = True
        if not emit_ok:
            return
        pct = calc_progress_pct(2, done_rows, prog_row_total)
        _progress_write(
            progress_path,
            {
                "status": "RUN",
                "phase_i": 2,
                "phase": "Excelへ書き込み中",
                "done": done_rows,
                "total": prog_row_total,
                "pct": pct,
                "current_file": str_fname,
                "seq": progress_seq,
            },
        )
        progress_seq += 1
        last_progress_rows = done_rows
        last_progress_mono = time.monotonic()
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
                "total": prog_row_total,
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

    # 一括書込みは表示停止で高速化（ScreenUpdating 復帰は DONE 後に遅延）
    t_csv_read_wait0 = time.perf_counter()
    with xlc.suspend_sheet_updates(book, restore_on_exit=False):
        for df_chunk in reader:
            ms_csv_read_wait += _elapsed_ms(t_csv_read_wait0)
            t_csv_read_wait0 = time.perf_counter()
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
                        t_bound0 = time.perf_counter()
                        if max_col > 0:
                            _apply_sheet_text_number_format(sh_target, val_accum_in_sheet, max_col)
                        if xlc and max_col > 0:
                            xlc.clear_used_range_overflow(sh_target, val_accum_in_sheet, max_col)
                        _finalize_sheet_context(
                            book, sh_target, str_fname, val_size_lbl, str_enc_lbl,
                            val_accum_total, val_total, curr_part_idx, val_sheets_total,
                            is_split_mode, val_accum_in_sheet, str_base_resolved, sh_origin,
                        )
                        ms_sheet_boundary += _elapsed_ms(t_bound0)
                    curr_part_idx += 1
                    str_target_name = _csv_ld_target_sheet_name(
                        str_base_resolved, curr_part_idx, is_split_mode
                    )
                    if curr_part_idx == 1:
                        if _is_sheet_empty(sh_origin):
                            sh_target = sh_origin
                            _safe_rename_sheet(sh_target, str_target_name)
                        else:
                            sh_target = _add_new_sheet_direct(book, str_target_name)
                    else:
                        sh_target = _add_new_sheet_direct(book, str_target_name)
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

                t_list0 = time.perf_counter()
                if is_new_sheet_boundary:
                    if header_only_as_one_body_row:
                        matrix = df_slice.values.tolist()
                    else:
                        matrix = [df_slice.columns.tolist()] + df_slice.values.tolist()
                else:
                    matrix = df_slice.values.tolist()
                ms_matrix_tolist += _elapsed_ms(t_list0)
                if matrix and len(matrix[0]) > max_col:
                    max_col = len(matrix[0])

                start_row = val_accum_in_sheet + 1
                nrows = len(matrix)
                ncol = len(matrix[0]) if matrix else 0
                pending_base = val_accum_total
                if xlc and nrows:

                    def _write_progress_cb(sub_rows: int, _base: int = pending_base) -> None:
                        if not matrix:
                            return
                        ratio = min(1.0, float(sub_rows) / float(len(matrix)))
                        est_done = _base + int(rows_to_write * ratio)
                        _emit_write_progress(est_done, time_only=True)

                    t_write0 = time.perf_counter()
                    if use_fast_text_write:
                        if ncol > 0:
                            _apply_range_text_number_format(
                                sh_target, start_row, nrows, ncol
                            )
                        xlc.write_chunk(
                            sh_target,
                            start_row,
                            1,
                            matrix,
                            None,
                            chunk_rows=write_step_v,
                            progress_notify_rows=progress_notify_rows,
                            progress_cb=_write_progress_cb,
                            text_mode=False,
                        )
                    else:
                        xlc.write_chunk(
                            sh_target,
                            start_row,
                            1,
                            matrix,
                            None,
                            chunk_rows=write_step_v,
                            progress_notify_rows=progress_notify_rows,
                            progress_cb=_write_progress_cb,
                            text_mode=True,
                        )
                    ms_excel_write += _elapsed_ms(t_write0)

                val_accum_total += rows_to_write
                val_accum_in_sheet += len(matrix)
                processed_chunk_idx += rows_to_write
                _emit_write_progress(val_accum_total)

        # 更新停止のまま: 最終シートのクリア・確定 → オートフィット → その後 with を抜けて更新開始
        if sh_target is not None:
            t_fin0 = time.perf_counter()
            if max_col > 0:
                _apply_sheet_text_number_format(sh_target, val_accum_in_sheet, max_col)
            if xlc and max_col > 0:
                xlc.clear_used_range_overflow(sh_target, val_accum_in_sheet, max_col)
            _finalize_sheet_context(
                book, sh_target, str_fname, val_size_lbl, str_enc_lbl,
                val_accum_total, val_total, curr_part_idx, val_sheets_total,
                is_split_mode, val_accum_in_sheet, str_base_resolved, sh_origin,
            )
            ms_finalize += _elapsed_ms(t_fin0)

        # 工程3: 列幅調整（オートフィット）。更新停止中に実行し、with 抜けで更新開始
        if progress_path is not None:
            _progress_write(
                progress_path,
                {
                    "status": "RUN",
                    "phase_i": 3,
                    "phase": "列幅調整中",
                    "done": prog_row_total,
                    "total": prog_row_total,
                    "pct": calc_progress_pct(3, prog_row_total, prog_row_total),
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
                            "done": prog_row_total,
                            "total": prog_row_total,
                            "pct": calc_progress_pct(3, prog_row_total, prog_row_total),
                            "current_file": str_fname,
                            "seq": progress_seq,
                        },
                    )
                    progress_seq += 1
                try:
                    t_af0 = time.perf_counter()
                    sheet_name = _csv_ld_target_sheet_name(
                        str_base_resolved, part_i, is_split_mode
                    )
                    last_row_i = MAX_ROWS_PER_SHEET if part_i < curr_part_idx else val_accum_in_sheet
                    for s in book.sheets:
                        if getattr(s, "name", "") != sheet_name:
                            continue
                        _autofit_used_range(s, last_row_i, max_col, sheet_name)
                        break
                    ms_autofit += _elapsed_ms(t_af0)
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
                    "done": prog_row_total,
                    "total": prog_row_total,
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
            wait_after_progress_done(min_sec=1.0)

    jit_total_ms = _elapsed_ms(t_jit_body0)
    _log_jit_breakdown(
        csv_read_wait_ms=ms_csv_read_wait,
        matrix_tolist_ms=ms_matrix_tolist,
        excel_write_ms=ms_excel_write,
        sheet_boundary_ms=ms_sheet_boundary,
        finalize_ms=ms_finalize,
        autofit_ms=ms_autofit,
        jit_total_ms=jit_total_ms,
        fast_text_write=use_fast_text_write,
    )


def _do_load_csv(
    book: Any,
    sheet_id: str,
    str_csv_path: str,
    parent_hwnd: int,
    progress_ui_already_shown: bool = False,
    t_load0: float = 0.0,
    attach_keys: tuple[int, str, str] | None = None,
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
        _progress_write_sheet_error(progress_path, sheet_id)
        _bring_excel_to_front(parent_hwnd)
        restore_excel_host_after_operation(parent_hwnd, sheet_id)
        return
    book, sh_origin = _resolve_book_and_sheet(
        book, sheet_id, parent_hwnd, attach_keys=attach_keys
    )
    if sh_origin is None:
        logger.error("[CSV_LD] 対象シートなし GUID=%s", sheet_id)
        _progress_write_sheet_error(progress_path, sheet_id)
        _bring_excel_to_front(parent_hwnd)
        restore_excel_host_after_operation(parent_hwnd, sheet_id)
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
            restore_excel_host_after_operation(parent_hwnd, sheet_id)
            return

    try:
        set_excel_performance_mode(book.app, True)
        t_jit0 = time.perf_counter()
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
        set_excel_performance_mode(book.app, False)
        restore_excel_host_after_operation(parent_hwnd, sheet_id, book.app)
        if xlc:
            xlc.yield_to_excel()


def _watch_result(
    result_path: str,
    book: Any,
    sheet_id: str,
    parent_hwnd: int,
    t_load0: float,
    pick_anchor: list[float | None] | None = None,
    attach_keys: tuple[int, str, str] | None = None,
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
                    attach_keys=attach_keys,
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
            notify_wait_form_ready(book=book)
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

    attach_keys = _capture_book_attach_keys(book)
    pick_anchor: list[float | None] = [None]
    _watch_result(
        res_path,
        book,
        str(sheet_id or ""),
        parent_hwnd,
        t_load0,
        pick_anchor,
        attach_keys=attach_keys,
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
