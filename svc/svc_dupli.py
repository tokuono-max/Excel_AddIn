# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: svc/svc_dupli.py
Created: 2026-03-06
Updated: 2026-05-05
Version: 1.5.2
Purpose:
  重複チェック（Excel 選択範囲）。処理は本モジュール、画面は ui_qt.ui_dupli + config/ui_dupli.json。
History (latest 3):
  - 1.5.2 (2026-05-05) 進捗→結果の順序保証: progress_closed_ack を追加。進捗画面クローズACKを待ってから done/report を投入し、重なり・ちらつきを抑制。
  - 1.5.1 (2026-04-10) PHASE_ANALYZE 中の進捗: モード A/B の長いループで間引き _upd（45〜92%）。pickle 負荷は時間＋ストライドで抑制。
  - 1.5.0 (2026-04-12) 重複着色: svc 一括 Interior は廃止。hl_rects を sidecar pickle に書き、レポート表示中は UI が VisibleRange±余白で追従着色。重複開始時に古い dupli_hl_rects_*.pkl を掃除。
  - 1.4.4 (2026-04-10) キャンセル: 時間ゲート疑似割り込み（約60ms毎に pickle 確認・force で即確認）。未着色キャンセルは DONE に短い done_delay_ms。groupby/COM/ハイライト矩形化のループにもチェック。
  - 1.4.3 (2026-04-10) 進捗キャンセル: cancel_request_path 経由の中止、着色済み矩形の部分クリアと PHASE_CLEAR 進捗。読込チャンク・解析・着色ループで協力的中止。
  - 1.4.2 (2026-04-11) 診断: レポート投入時に highlight_clear の rects 件数・キー・book/sheet 有無を hc_csv_tool.diag.dupli へ。
  - 1.4.1 (2026-04-11) 重複ハイライト COM ループ中: ScreenUpdating=False・Calculation=xlCalculationManual、finally で計算→画面→Interactive の順に復元。
  - 1.4.0 (2026-04-11) bridge: VBA の selection_count_large / sheet_cells_count_large で全シート判定を優先し、取れるときは COM Selection を読まない。
  - 1.3.9 (2026-04-11) 仕様書 §3.0: モード B は左上コーナー（全シート）選択のみ。単一セル→B・Used 外交差の単一セル→B を廃止。bridge 時も COM の全シート判定を OR。
  - 1.3.8 (2026-04-11) bridge: 列全体/行全体等は Application.Range / シートローカルで解決。_dupli_sel_log の full_sheet キー重複を解消。
  - 1.3.7 (2026-04-11) bridge JSON の selection_areas（VBA Areas×External 付き）で交差矩形を構築し、Python 側の Application.Selection 読取を回避。失敗時は COM にフォールバック。
  - 1.3.6 (2026-04-11) DISP_E_EXCEPTION を hresult の U32（0x80020009）で判定し、details タプル内のいずれかがビジー HRESULT ならリトライ対象に。
  - 1.3.5 (2026-04-11) ビジー HRESULT を符号なしで照合（0x80027EFA 等・DISP_E 内側）。selection_used_rects 失敗ログを is_busy 明示。
  - 1.3.4 (2026-04-11) Selection/UsedRange 読取で RPC_E_SERVERCALL_RETRYLATER 等のとき短いバックオフでリトライ（レポート直後の対象なし緩和）。
  - 1.3.3 (2026-04-11) 全シート判定: CountLarge 優先。交差矩形なしでも全シート相当ならモードB。[DUPLI_SEL] 診断ログ。
  - 1.3.2 (2026-04-10) 全シート選択（左上コーナー）を検知しモードBと同じセル単位走査。highlight_clear に book_name。
  - 1.3.1 (2026-04-11) 単一セルが UsedRange と交差しない場合もモードB。highlight_clear に sheet_name。UI 側はシート固定＋Interior 強化クリア・Goto スクロール。
  - 1.3.0 (2026-04-11) モード A: Areas×UsedRange・選択列のみ着色・空欄正規化。モード B: 単一セルで有効データ領域をセル値重複。highlight_clear.rects。仕様書 §3 準拠。
  - 1.2.1 (2026-04-11) UI 順序: 完了通知 pickle をレポートより先に投入（mtime 順）。着色クリアはレポート閉鎖時のみ（完了 OK では消さない）。
  - 1.2.0 (2026-04-11) レポート: 同一値グループを先頭行の若い順で一覧。タイトルと総数表示を分離。UI 閉鎖で着色クリア用 runs を渡す。
  - 1.1.1 (2026-04-11) 重複着色: 連続行を 1 レンジにマージして COM 削減。進捗 _upd を行数・時間の複合間隔に変更。
  - 1.1.0 (2026-04-06) HC_LOG_PERF: [DUPLI_PERF]。診断: [DUPLI_TRACE]。
  - 1.0.0 (2026-03-11) core_xlc.get_excel_context_from_hwnd 利用、完了通知・進捗・レポート表示、ログ出力（範囲・重複有無・重複数）。
  - 初出 (2026-03-06) hc_dupli から分離。svc_dupli + ui_dupli + config/ui_dupli.json で完結。
"""
from __future__ import annotations

import os
import sys
import threading
import time
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Callable, Optional

import pandas as pd

# プロジェクトルートをパスに追加（Excel アドインから呼ばれるため）
_path_svc = os.path.dirname(os.path.abspath(__file__))
_root = os.path.dirname(_path_svc)
if _root not in sys.path:
    sys.path.insert(0, _root)

from core.core_log import get_diag_logger, get_logger, get_perf_logger  # noqa: E402
from ui_qt.ipc_file import get_ipc_root, get_request_dir, read_pickle, write_pickle  # noqa: E402

logger = get_logger(__name__)
_dupli_diag = get_diag_logger("hc_csv_tool.diag.dupli")
_perf = get_perf_logger("svc.svc_dupli.perf")
__version__ = "1.5.2"


class _DupliCancelled(Exception):
    """ユーザーが進捗のキャンセルを要求したときに協力的に打ち切る。"""

# Excel COM（win32 / xlwings.api と同値）
_XL_CALC_MANUAL = -4135  # xlCalculationManual
_XL_CALC_AUTOMATIC = -4105  # xlCalculationAutomatic

# キャンセル要求の疑似割り込み: この間隔（秒）を空けて cancel pickle を読む（毎ループは monotonic のみ）
_CANCEL_POLL_INTERVAL_SEC = 0.06
# 着色前にキャンセルしたとき、進捗を閉じるまでの短い待ち（ms）
_CANCEL_DONE_DELAY_MS_NO_HIGHLIGHT = 180

# 解析フェーズ（45%〜）の進捗 pickle 更新: 時間＋件数ストライドで間引き
_DUPLI_ANALYZE_PROGRESS_MIN_INTERVAL_SEC = 0.22
_DUPLI_ANALYZE_STRIDE_CELLS = 2048
_DUPLI_ANALYZE_STRIDE_ROWS = 32
_DUPLI_ANALYZE_STRIDE_GROUPS = 12
_DUPLI_ANALYZE_STRIDE_HL = 512
_PROGRESS_CLOSE_ACK_TIMEOUT_SEC = 3.0
_PROGRESS_CLOSE_ACK_POLL_SEC = 0.03


class _DupliAnalyzeProgressEmitter:
    """PHASE_ANALYZE 中の pct を lo..hi に割り当て、単調に間引き更新する。"""

    def __init__(
        self,
        upd: Callable[[int, str, str], None],
        phase: str,
        cur: str,
        *,
        start_pct: int = 45,
        min_interval_sec: float = _DUPLI_ANALYZE_PROGRESS_MIN_INTERVAL_SEC,
        stride: int = _DUPLI_ANALYZE_STRIDE_CELLS,
    ) -> None:
        self._upd = upd
        self._phase = phase
        self._cur = cur
        self._min_interval = max(0.05, float(min_interval_sec))
        self._stride = max(1, int(stride))
        self._last_emit_t = 0.0
        self._last_pct = max(-1, int(start_pct))

    def set_stride(self, stride: int) -> None:
        self._stride = max(1, int(stride))

    def tick(self, processed: int, total: int, lo: int, hi: int) -> None:
        if total <= 0:
            return
        tot = max(1, int(total))
        proc = max(0, min(int(processed), tot))
        now = time.perf_counter()
        at_end = proc >= tot

        pct_raw = lo + int((hi - lo) * proc / tot + 1e-9)
        pct_raw = max(lo, min(hi, pct_raw))
        pct = max(pct_raw, self._last_pct)

        if at_end:
            pct = max(self._last_pct, hi)
            self._upd(min(100, pct), self._phase, self._cur)
            self._last_emit_t = now
            self._last_pct = min(100, pct)
            return

        if pct <= self._last_pct:
            return
        stride_hit = proc > 0 and proc % self._stride == 0
        time_hit = (now - self._last_emit_t) >= self._min_interval
        if not stride_hit and not time_hit:
            return

        self._upd(min(100, pct), self._phase, self._cur)
        self._last_emit_t = now
        self._last_pct = min(100, pct)


def _elapsed_ms(since: float) -> int:
    return max(0, int((time.perf_counter() - since) * 1000))


def _dupli_trace(phase: str, t0: float, **kv: object) -> None:
    try:
        if kv:
            _dupli_diag.info(
                "[DUPLI_TRACE] phase=%s cumulative_ms=%d %s",
                phase,
                _elapsed_ms(t0),
                " ".join("%s=%s" % (k, v) for k, v in kv.items()),
            )
        else:
            _dupli_diag.info("[DUPLI_TRACE] phase=%s cumulative_ms=%d", phase, _elapsed_ms(t0))
    except Exception:
        pass


def _perf_dupli(phase: str, t0: float, **kv: object) -> None:
    try:
        if kv:
            _perf.info(
                "[DUPLI_PERF] phase=%s cumulative_ms=%d %s",
                phase,
                _elapsed_ms(t0),
                " ".join("%s=%s" % (k, v) for k, v in kv.items()),
            )
        else:
            _perf.info("[DUPLI_PERF] phase=%s cumulative_ms=%d", phase, _elapsed_ms(t0))
    except Exception:
        pass


def _cleanup_stale_dupli_hl_rect_sidecars(*, max_age_sec: float = 86400.0) -> None:
    """古い hl_rects sidecar（クラッシュ残り）を削除する。"""
    d = Path(get_ipc_root()) / "progress"
    if not d.is_dir():
        return
    now = time.time()
    for p in d.glob("dupli_hl_rects_*.pkl"):
        try:
            if now - p.stat().st_mtime > max_age_sec:
                p.unlink(missing_ok=True)
        except OSError:
            pass


def _write_dupli_hl_rects_sidecar(sheet_id: str, hl_rects: list[list[int]]) -> Path:
    d = Path(get_ipc_root()) / "progress"
    d.mkdir(parents=True, exist_ok=True)
    sid = str(sheet_id or "").strip() or "_"
    ts_ms = int(time.time() * 1000)
    path = d / f"dupli_hl_rects_{sid}_{ts_ms}_{os.getpid()}.pkl"
    write_pickle(path, {"v": 1, "rects": hl_rects})
    return path


try:
    from core import core_cst as cst
except Exception:
    cst = None  # type: ignore

try:
    from core import core_xlc as core_xlc_mod
except Exception:
    core_xlc_mod = None  # type: ignore


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
    重複チェック用の画面・メッセージ設定を config/ui_dupli.json から読み込む。
    読込失敗時は UiConfigLoadError が発生する（救済なし）。
    """
    if cst is None:
        return {}
    return cst.get_ui_config_from_file_required("dupli")


def _msg(cfg: dict[str, Any], key: str, **fmt: Any) -> str:
    """
    設定の MESSAGES からキーに対応する文言を取得し、任意でフォーマットする。
    例: _msg(cfg, "STATUS_FINAL", count=5) → "重複: 5 件を…"
    """
    m = (cfg.get("MESSAGES") or {}).get(key) or key
    try:
        return str(m).format(**fmt)
    except Exception:
        return str(m)


def _progress_path(sheet_id: str) -> Path:
    """
    進捗状態を書き出す Pickle ファイルのパスを返す。
    ui_server 側の進捗ダイアログがこのファイルをポーリングして表示を更新する。
    """
    d = Path(get_ipc_root()) / "progress"
    d.mkdir(parents=True, exist_ok=True)
    return d / f"progress_dupli_{sheet_id}.pkl"


def _cancel_request_path(sheet_id: str) -> Path:
    """進捗キャンセル要求用（メイン進捗 pickle の status=CANCEL と衝突させない）。"""
    d = Path(get_ipc_root()) / "progress"
    d.mkdir(parents=True, exist_ok=True)
    return d / f"cancel_req_dupli_{sheet_id}.pkl"


def _progress_closed_ack_path(sheet_id: str) -> Path:
    """進捗ダイアログが閉じたことを示す ACK ファイル。"""
    d = Path(get_ipc_root()) / "progress"
    d.mkdir(parents=True, exist_ok=True)
    return d / f"progress_dupli_closed_{sheet_id}.pkl"


def _wait_progress_closed_ack(path: Optional[Path], timeout_sec: float = _PROGRESS_CLOSE_ACK_TIMEOUT_SEC) -> None:
    """
    進捗画面のクローズACKを短時間待つ。
    タイムアウト時はログだけ残して先へ進む（処理全体を止めない）。
    """
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
            logger.info("[DUPLI] progress close ack timeout: %s", str(p))
            return
        time.sleep(_PROGRESS_CLOSE_ACK_POLL_SEC)


def _reset_cancel_path(path: Path) -> None:
    try:
        path.unlink(missing_ok=True)  # type: ignore[call-arg]
    except TypeError:
        try:
            if path.exists():
                path.unlink()
        except Exception:
            pass
    except Exception:
        pass


def _cancel_requested(path: Path) -> bool:
    try:
        d = read_pickle(path)
        return isinstance(d, dict) and bool(d.get("cancel"))
    except Exception:
        return False


def _make_time_gated_cancel_check(cancel_path: Path) -> Callable[..., None]:
    """
    時間ベースの疑似割り込み用チェック。
    force=True のときは間隔に関係なく cancel pickle を読む（フェーズ境界・読込チャンク先頭など）。
    """
    last_poll: list[float] = [0.0]

    def _chk(*, force: bool = False) -> None:
        now = time.monotonic()
        if not force and (now - last_poll[0]) < _CANCEL_POLL_INTERVAL_SEC:
            return
        last_poll[0] = now
        if _cancel_requested(cancel_path):
            raise _DupliCancelled()

    return _chk


def _progress_write(path: Path, obj: dict[str, Any]) -> None:
    """
    進捗情報を Pickle で path に書き出す。
    seq が未指定の場合は既存ファイルの seq をインクリメントして順序を保証する。
    """
    try:
        from ui_qt.ipc_file import read_pickle

        obj = dict(obj)
        if "seq" not in obj:
            try:
                prev = read_pickle(path)
                seq = int(prev.get("seq", -1)) + 1 if isinstance(prev, dict) else 0
            except Exception:
                seq = 0
            obj["seq"] = seq
        write_pickle(path, obj)
    except Exception:
        pass


def _get_window_rect(hwnd: int) -> tuple[int, int, int, int] | None:
    """
    Win32 API でウィンドウのクライアント外枠（left, top, right, bottom）を取得する。
    進捗ダイアログを Excel ウィンドウ付近に表示する際の基準に使う。
    """
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


def _submit_progress_ui(
    parent_hwnd: int,
    sheet_id: str,
    progress_path: Path,
    phase_total: int,
    *,
    cancel_request_path: Optional[Path] = None,
    progress_closed_path: Optional[Path] = None,
) -> None:
    """
    UI サーバに進捗画面表示を依頼する。req_*.pkl に payload を書き、ui_server が ui_dupli.create_dialog を呼ぶ。
    進捗の実データは progress_path の Pickle を ui 側がポーリングして表示する。
    cancel_request_path を渡すと進捗ダイアログにキャンセルボタンが出る（重複チェック専用）。
    """
    try:
        from svc.svc_host import ensure_ui_server

        ensure_ui_server()
        excel_rect = _get_window_rect(int(parent_hwnd or 0))
        res_dir = Path(get_ipc_root()) / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        ts_ms = int(time.time() * 1000)
        result_path = str(res_dir / f"res_progress_dupli_{ts_ms}_{os.getpid()}.pkl")
        req_dict: dict[str, Any] = {
            "action": "progress",
            "progress_path": str(progress_path),
            "phase_total": int(phase_total),
            "excel_lock": False,
            "no_native_window": True,
        }
        if cancel_request_path is not None:
            cr = str(cancel_request_path).strip()
            if cr:
                req_dict["cancel_request_path"] = cr
        if progress_closed_path is not None:
            cp = str(progress_closed_path).strip()
            if cp:
                req_dict["progress_closed_path"] = cp
        if excel_rect is not None:
            req_dict["excel_rect"] = list(excel_rect)
        payload = {
            "parent_hwnd": int(parent_hwnd),
            "result_path": result_path,
            "ready_path": "",
            "sheet_id": str(sheet_id),
            "log_path": "",
            "action": "progress",
            "module": "ui_qt.ui_dupli",
            "req_dict": req_dict,
        }
        req_path = get_request_dir() / f"req_{ts_ms}_{os.getpid()}_{threading.get_ident()}.pkl"
        write_pickle(req_path, payload)
        logger.info("[DUPLI] progress UI request: %s", req_path)
    except Exception as exc:
        logger.warning("[DUPLI] progress UI request failed: %s", exc)


def _submit_report_ui(
    parent_hwnd: int,
    sheet_id: str,
    title: str,
    headers: list[dict[str, Any]],
    rows: list[list[str]],
    addresses: list[str],
    *,
    report_intro: str = "",
    dup_count: int = 0,
    count_caption_template: str = "",
    highlight_clear: Optional[dict[str, Any]] = None,
    link_col: int = 1,
) -> None:
    """
    重複検出レポートをモードレスで表示するよう UI サーバに依頼する。
    行・座標・内容の一覧と「セルへ移動」ボタン付きの DupliReportDialog が表示される。
    link_col はジャンプに使う列の 0 始まりインデックス（モード B は代表座標列＝既定 2）。
    """
    try:
        hl0 = highlight_clear if isinstance(highlight_clear, dict) else {}
        _rl = hl0.get("rects")
        _rects_n = len(_rl) if isinstance(_rl, list) else -1
        _runs_l = hl0.get("runs")
        _runs_n = len(_runs_l) if isinstance(_runs_l, list) else -1
        _rp = str(hl0.get("rects_path") or "").strip()
        _rc = hl0.get("rects_count")
        _dupli_diag.info(
            "[DUPLI_REP] submit_report parent_hwnd=%s sheet_id=%s highlight_attached=%s keys=%s rects_n=%s runs_n=%s rects_path_set=%s rects_count=%s viewport=%s book_name_set=%s sheet_name_set=%s",
            int(parent_hwnd or 0),
            str(sheet_id or ""),
            bool(highlight_clear),
            sorted(str(k) for k in hl0.keys()),
            _rects_n,
            _runs_n,
            bool(_rp),
            _rc,
            bool(hl0.get("viewport_follow")),
            bool(str(hl0.get("book_name") or "").strip()),
            bool(str(hl0.get("sheet_name") or "").strip()),
        )
    except Exception:
        pass
    try:
        from svc.svc_host import ensure_ui_server

        ensure_ui_server()
        ts_ms = int(time.time() * 1000)
        res_dir = Path(get_ipc_root()) / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        result_path = str(res_dir / f"res_dupli_report_{ts_ms}_{os.getpid()}.pkl")
        req_dict: dict[str, Any] = {
            "action": "dupli_report",
            "modeless": True,
            "title": title,
            "headers": headers,
            "rows": rows,
            "addresses": addresses,
            "link_col": int(link_col),
            "report_intro": str(report_intro or ""),
            "dup_count": int(dup_count),
            "count_caption_template": str(count_caption_template or ""),
        }
        if highlight_clear:
            req_dict["highlight_clear"] = highlight_clear
        er_r = _get_window_rect(int(parent_hwnd or 0))
        if er_r is not None:
            req_dict["excel_rect"] = list(er_r)
        payload = {
            "parent_hwnd": int(parent_hwnd),
            "result_path": result_path,
            "ready_path": "",
            "sheet_id": str(sheet_id),
            "log_path": "",
            "action": "dupli",
            "module": "ui_qt.ui_dupli",
            "req_dict": req_dict,
        }
        req_path = get_request_dir() / f"req_{ts_ms}_{os.getpid()}_rpt.pkl"
        write_pickle(req_path, payload)
        logger.info("[DUPLI] report UI request: %s", req_path)
    except Exception as exc:
        logger.warning("[DUPLI] report UI request failed: %s", exc)


def _submit_done_ui(
    parent_hwnd: int,
    sheet_id: str,
    message: str,
    title: str = "重複チェック",
) -> None:
    """
    完了通知（重複あり/なしの結果メッセージ）をモーダルで表示するため ui_server に依頼する。
    SCREENS.DONE の設定に従い、アイコン・中央表示・OK ボタンで閉じる。
    """
    try:
        from svc.svc_host import ensure_ui_server

        ensure_ui_server()
        ts_ms = int(time.time() * 1000)
        res_dir = Path(get_ipc_root()) / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        result_path = str(res_dir / f"res_dupli_done_{ts_ms}_{os.getpid()}.pkl")
        req_dict: dict[str, Any] = {
            "action": "dupli_done",
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
            "action": "dupli",
            "module": "ui_qt.ui_dupli",
            "req_dict": req_dict,
        }
        req_path = get_request_dir() / f"req_{ts_ms}_{os.getpid()}_done.pkl"
        write_pickle(req_path, payload)
        logger.info("[DUPLI] done UI request: %s", req_path)
    except Exception as exc:
        logger.warning("[DUPLI] done UI request failed: %s", exc)


def _norm_dup_cell(v: Any) -> str:
    """重複比較用にセル値を正規化（空同士一致・表示に近い文字列化）。"""
    if v is None:
        return ""
    try:
        if isinstance(v, float) and pd.isna(v):
            return ""
    except Exception:
        pass
    if isinstance(v, str):
        return v.strip()
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (int, float)):
        if isinstance(v, float):
            try:
                if abs(v - round(v)) < 1e-9:
                    return str(int(round(v)))
            except Exception:
                pass
        return str(v)
    try:
        from datetime import date, datetime

        if isinstance(v, (datetime, date)):
            return str(v)
    except Exception:
        pass
    return str(v).strip()


def _is_trim_empty(v: Any) -> bool:
    if v is None:
        return True
    try:
        if isinstance(v, float) and pd.isna(v):
            return True
    except Exception:
        pass
    return isinstance(v, str) and not v.strip()


def _trim_used_matrix(
    arr: list[list[Any]], uy1: int, ux1: int
) -> tuple[int, int, int, int, list[list[Any]]]:
    """UsedRange 行列の周囲の空行・空列を除いた有効データ領域 (vy1, vx1, vyn, vxn, matrix)。"""
    if not arr:
        return uy1, ux1, 0, 0, []
    a = [list(row) for row in arr]
    y, x = uy1, ux1
    while a and all(_is_trim_empty(c) for c in a[0]):
        a.pop(0)
        y += 1
    while a and all(_is_trim_empty(c) for c in a[-1]):
        a.pop()
    if not a:
        return y, x, 0, 0, []
    while a[0] and all(_is_trim_empty(row[0]) for row in a if row):
        for row in a:
            if row:
                row.pop(0)
        x += 1
    if not a or not a[0]:
        return y, x, 0, 0, []
    while a[0] and all(_is_trim_empty(row[-1]) for row in a if row):
        for row in a:
            if row:
                row.pop()
    if not a:
        return y, x, 0, 0, []
    vyn, vxn = len(a), len(a[0])
    return y, x, vyn, vxn, a


# Excel COM: UI 切替直後に Selection が一時的に読めないときの HRESULT（符号付きでブレるため U32 で照合）
_DISP_E_EXCEPTION_U32 = 0x80020009  # DISP_E_EXCEPTION（外側 hresult が正負どちらでも U32 で一致）
# 0x8001010A RPC_E_SERVERCALL_RETRYLATER, 0x80010101 RPC_E_SERVERCALL_WAIT,
# 0x80027EFA は DISP_E 内側で観測（日本語 Excel / pywin32 で Selection 取得失敗時）
_BUSY_HRESULT_U32: frozenset[int] = frozenset(
    {
        0x8001010A,
        0x80010101,
        0x80027EFA,
    }
)


def _hresult_u32(code: int) -> int:
    return int(code) & 0xFFFFFFFF


def _excel_com_is_busy(exc: BaseException) -> bool:
    """Application.Selection 等が「後で再試行」と返す COM エラーか。"""
    try:
        from pywintypes import com_error
    except ImportError:
        return False
    if not isinstance(exc, com_error):
        return False
    hr = int(exc.hresult)
    if _hresult_u32(hr) in _BUSY_HRESULT_U32:
        return True
    # pywin32 により hresult が符号付き・符号なしでブレるため DISP_E は U32 で判定
    if _hresult_u32(hr) == _DISP_E_EXCEPTION_U32 and len(exc.args) > 2:
        det = exc.args[2]
        if isinstance(det, tuple):
            for el in det:
                try:
                    if _hresult_u32(int(el)) in _BUSY_HRESULT_U32:
                        return True
                except (TypeError, ValueError):
                    continue
    return False


def _excel_busy_retry(op: Callable[[], Any], *, attempts: int = 10, base_delay: float = 0.06) -> Any:
    """ビジー系 COM エラーの間だけスリープを挟んで再試行。それ以外は直ちに再送出。"""
    last: BaseException | None = None
    for attempt in range(attempts):
        try:
            return op()
        except BaseException as e:
            last = e
            if _excel_com_is_busy(e) and attempt + 1 < attempts:
                time.sleep(base_delay * (attempt + 1))
                continue
            raise
    assert last is not None
    raise last


def _selection_used_rects_once(p_app_v: Any, p_sh_v: Any) -> Optional[list[tuple[int, int, int, int]]]:
    """交差矩形の算出（単発）。例外は呼び出し側でリトライ。"""
    ptr_sel = p_app_v.selection
    ptr_usd = p_sh_v.used_range
    if ptr_sel is None or ptr_usd is None:
        return None
    api_app = p_app_v.api
    api_sel = ptr_sel.api
    api_usd = ptr_usd.api
    rects: list[tuple[int, int, int, int]] = []
    n_areas = int(getattr(api_sel.Areas, "Count", 0) or 0)
    if n_areas < 1:
        inter = api_app.Intersect(api_sel, api_usd)
        if inter is None:
            return None
        rects.append(
            (int(inter.Row), int(inter.Column), int(inter.Rows.Count), int(inter.Columns.Count))
        )
        return rects
    for i in range(1, n_areas + 1):
        area = api_sel.Areas(i)
        inter = api_app.Intersect(area, api_usd)
        if inter is None:
            continue
        rects.append(
            (int(inter.Row), int(inter.Column), int(inter.Rows.Count), int(inter.Columns.Count))
        )
    return rects if rects else None


def _selection_used_rects(p_app_v: Any, p_sh_v: Any) -> Optional[list[tuple[int, int, int, int]]]:
    """
    Selection の各 Area と UsedRange の交差矩形の一覧を (Row, Column, RowCount, ColCount) で返す。
    交差が無い場合は None または空リスト。COM ビジー時は短いバックオフで再試行する。
    """
    try:
        return _excel_busy_retry(lambda: _selection_used_rects_once(p_app_v, p_sh_v))
    except Exception as exc:
        try:
            _dupli_diag.info(
                "[DUPLI_SEL] selection_used_rects failed is_busy_hresult=%s exc=%r",
                _excel_com_is_busy(exc),
                exc,
            )
        except Exception:
            pass
        return None


def _bridge_sheet_and_local_from_external(addr: str) -> tuple[Optional[str], str]:
    """
    External 付きアドレスから（シート名, ローカル A1 部分）を推定。
    例: [Book.xls]SMRT_01!$C:$C → (SMRT_01, $C:$C)
    """
    a = addr.strip().lstrip("=")
    if "!" not in a:
        return None, a
    left, right = a.rsplit("!", 1)
    local = right.strip()
    left = left.strip().strip("'\"")
    sheet_name: Optional[str] = None
    if "]" in left:
        sheet_name = left.split("]", 1)[-1].strip() or None
    else:
        sheet_name = left.strip() or None
    return sheet_name, local


def _bridge_same_sheet_name(a: str, b: str) -> bool:
    return (a or "").strip().casefold() == (b or "").strip().casefold()


def _bridge_resolve_range_api(
    addr: str,
    p_wb_v: Any,
    ptr_s: Any,
    p_app_v: Any,
) -> Optional[Any]:
    """
    bridge 用に COM Range を返す。book.range が列全体等で失敗しやすいため多段フォールバック。
    """
    a = addr.strip()
    if not a:
        return None
    api_app = p_app_v.api
    target_sheet = str(ptr_s.name or "")

    def _ok_parent(api_rng: Any) -> bool:
        try:
            par = api_rng.Parent
            sn = str(getattr(par, "Name", "") or "")
            return _bridge_same_sheet_name(sn, target_sheet)
        except Exception:
            return False

    for getter in (
        lambda: p_wb_v.range(a).api,
        lambda: p_app_v.range(a).api,
        lambda: api_app.Range(a),
    ):
        try:
            api_r = getter()
            if api_r is not None and _ok_parent(api_r):
                return api_r
        except Exception:
            continue

    sheet_guess, local = _bridge_sheet_and_local_from_external(a)
    if (
        local
        and sheet_guess is not None
        and _bridge_same_sheet_name(sheet_guess, target_sheet)
    ):
        try:
            api_r = ptr_s.range(local).api
            if api_r is not None and _ok_parent(api_r):
                return api_r
        except Exception:
            pass

    try:
        _dupli_diag.info("[DUPLI_SEL] bridge range resolve failed addr=%r", a[:240])
    except Exception:
        pass
    return None


def _selection_used_rects_from_bridge_addresses_once(
    p_wb_v: Any,
    ptr_s: Any,
    p_app_v: Any,
    addresses: list[str],
) -> Optional[list[tuple[int, int, int, int]]]:
    """各 External 付き Area をブックから解決し、ptr_s の UsedRange との交差矩形を列挙（単発）。"""
    ptr_usd = ptr_s.used_range
    if ptr_usd is None:
        return None
    api_app = p_app_v.api
    api_usd = ptr_usd.api
    rects: list[tuple[int, int, int, int]] = []
    for addr in addresses:
        a = addr.strip()
        if not a:
            continue
        api_rng = _bridge_resolve_range_api(a, p_wb_v, ptr_s, p_app_v)
        if api_rng is None:
            continue
        try:
            inter = api_app.Intersect(api_rng, api_usd)
        except Exception:
            continue
        if inter is None:
            continue
        rects.append(
            (int(inter.Row), int(inter.Column), int(inter.Rows.Count), int(inter.Columns.Count))
        )
    return rects if rects else None


def _selection_used_rects_from_bridge_addresses(
    p_wb_v: Any,
    ptr_s: Any,
    p_app_v: Any,
    addresses: list[str],
) -> Optional[list[tuple[int, int, int, int]]]:
    try:
        return _excel_busy_retry(
            lambda: _selection_used_rects_from_bridge_addresses_once(
                p_wb_v, ptr_s, p_app_v, addresses
            )
        )
    except Exception as exc:
        try:
            _dupli_diag.info(
                "[DUPLI_SEL] bridge selection_areas failed is_busy=%s exc=%r",
                _excel_com_is_busy(exc),
                exc,
            )
        except Exception:
            pass
        return None


def _selection_full_sheet_flags_once(p_app_v: Any, p_sh_v: Any) -> tuple[bool, str]:
    """全シート相当判定（単発）。"""
    ptr_sel = p_app_v.selection
    if ptr_sel is None:
        return False, "no_selection"
    api_sh = p_sh_v.api
    api_sel = ptr_sel.api
    try:
        cells_cnt = int(api_sh.Cells.CountLarge)
        sel_cnt = int(api_sel.CountLarge)
        if cells_cnt > 0 and sel_cnt >= cells_cnt:
            return True, f"count_large_ok sel={sel_cnt} sheet_cells={cells_cnt}"
    except Exception:
        pass
    max_r = int(api_sh.Rows.Count)
    max_c = int(api_sh.Columns.Count)
    sel_r = int(api_sel.Rows.Count)
    sel_c = int(api_sel.Columns.Count)
    if sel_r >= max_r and sel_c >= max_c:
        return True, f"max_rows_cols_ok sel_r={sel_r} sel_c={sel_c} max_r={max_r} max_c={max_c}"
    return (
        False,
        f"not_full_sheet sel_r={sel_r} sel_c={sel_c} max_r={max_r} max_c={max_c} count_large_ng",
    )


def _selection_full_sheet_flags(p_app_v: Any, p_sh_v: Any) -> tuple[bool, str]:
    """
    シート全体選択に相当するか。Excel 版差を吸収するため CountLarge を優先し、次に行×列最大一致。
    戻り値: (判定, ログ用の理由短文)
    """
    try:
        return _excel_busy_retry(lambda: _selection_full_sheet_flags_once(p_app_v, p_sh_v))
    except Exception as exc:
        return False, f"exc={exc!r}"


def _full_sheet_from_bridge_count_large(
    selection_count_large: Optional[int],
    sheet_cells_count_large: Optional[int],
) -> tuple[bool, str]:
    """VBA bridge が渡す CountLarge 組で全シート選択相当か（Python 側の判定式）。"""
    if selection_count_large is None or sheet_cells_count_large is None:
        return False, "bridge_count_large_absent"
    if sheet_cells_count_large <= 0:
        return False, "bridge_count_large_sheet_non_positive"
    if selection_count_large >= sheet_cells_count_large:
        return True, f"bridge_count_large_ok sel={selection_count_large} sheet={sheet_cells_count_large}"
    return False, f"bridge_count_large_not_full sel={selection_count_large} sheet={sheet_cells_count_large}"


def _corner_full_sheet_flags(
    ptr_a: Any,
    ptr_s: Any,
    selection_count_large: Optional[int],
    sheet_cells_count_large: Optional[int],
) -> tuple[bool, str]:
    """bridge 数値が十分なら COM の Selection を読まず全シート判定。否则 COM 補助。"""
    b_ok, b_why = _full_sheet_from_bridge_count_large(
        selection_count_large, sheet_cells_count_large
    )
    if b_ok:
        return True, b_why
    return _selection_full_sheet_flags(ptr_a, ptr_s)


def _dupli_selection_diag_snapshot_once(p_app_v: Any, p_sh_v: Any) -> dict[str, Any]:
    """診断スナップショット（単発・例外はリトライ側へ）。"""
    out: dict[str, Any] = {"tag": "selection_diag"}
    ptr_sel = p_app_v.selection
    if ptr_sel is None:
        out["selection"] = None
    else:
        api_sel = ptr_sel.api
        out["sel_rows"] = int(ptr_sel.rows.count)
        out["sel_cols"] = int(ptr_sel.columns.count)
        try:
            out["sel_count_large"] = int(api_sel.CountLarge)
        except Exception:
            out["sel_count_large"] = None
        try:
            out["sheet_cells_count_large"] = int(p_sh_v.api.Cells.CountLarge)
        except Exception:
            out["sheet_cells_count_large"] = None
        try:
            addr = str(api_sel.Address(False, False))
        except Exception:
            addr = ""
        out["sel_address"] = (addr[:800] + "…") if len(addr) > 800 else addr
        try:
            out["areas_count"] = int(api_sel.Areas.Count)
        except Exception:
            out["areas_count"] = None
    ptr_usd = p_sh_v.used_range
    out["used_rows"] = int(ptr_usd.rows.count)
    out["used_cols"] = int(ptr_usd.columns.count)
    try:
        uaddr = str(ptr_usd.api.Address(False, False))
    except Exception:
        uaddr = ""
    out["used_address"] = (uaddr[:400] + "…") if len(uaddr) > 400 else uaddr
    fs, why = _selection_full_sheet_flags(p_app_v, p_sh_v)
    out["full_sheet"] = fs
    out["full_sheet_why"] = why
    return out


def _dupli_selection_diag_snapshot(p_app_v: Any, p_sh_v: Any) -> dict[str, Any]:
    """[DUPLI_SEL] 用。左上コーナー等の COM 実値の切り分け。ビジー時はリトライ。"""
    try:
        return _excel_busy_retry(lambda: _dupli_selection_diag_snapshot_once(p_app_v, p_sh_v))
    except Exception as exc:
        out: dict[str, Any] = {"tag": "selection_diag", "selection_error": repr(exc)}
        try:
            ptr_usd = p_sh_v.used_range
            out["used_rows"] = int(ptr_usd.rows.count)
            out["used_cols"] = int(ptr_usd.columns.count)
            try:
                uaddr = str(ptr_usd.api.Address(False, False))
            except Exception:
                uaddr = ""
            out["used_address"] = (uaddr[:400] + "…") if len(uaddr) > 400 else uaddr
        except Exception as e2:
            out["used_range_error"] = repr(e2)
        fs, why = _selection_full_sheet_flags(p_app_v, p_sh_v)
        out["full_sheet"] = fs
        out["full_sheet_why"] = why
        return out


def _dupli_sel_log(event: str, t0: float, **extra: object) -> None:
    try:
        snap = {k: v for k, v in extra.items() if v is not None}
        _dupli_diag.info(
            "[DUPLI_SEL] event=%s cumulative_ms=%d %s",
            event,
            _elapsed_ms(t0),
            " ".join("%s=%r" % (k, v) for k, v in sorted(snap.items())),
        )
    except Exception:
        pass


def _dupli_intersection_log_extra(
    ptr_a: Any,
    ptr_s: Any,
    rects_source: str,
    rects: Optional[list[tuple[int, int, int, int]]],
) -> dict[str, Any]:
    """bridge 成功時は Selection を読まずログ用キーのみ付与。"""
    out: dict[str, Any] = {"rects_source": rects_source}
    if rects is not None:
        out["selection_rects"] = rects
    if rects_source != "bridge":
        try:
            out.update(_dupli_selection_diag_snapshot(ptr_a, ptr_s))
        except Exception:
            out["selection_diag_error"] = True
    return out


def _sheet_name_for_highlight(ptr_s: Any) -> str:
    """レポート閉鎖時に着色を解除する対象シート名（アクティブが変わっても特定できるように）。"""
    try:
        return str(ptr_s.name or "").strip()
    except Exception:
        return ""


def _book_name_for_highlight(ptr_s: Any) -> str:
    """レポート閉鎖時に着色解除する対象ブック名（アクティブブックが変わっても特定しやすくする）。"""
    try:
        return str(ptr_s.book.name or "").strip()
    except Exception:
        return ""


def _sel_cols_per_row_from_rects(rects: list[tuple[int, int, int, int]]) -> dict[int, list[int]]:
    m: dict[int, set[int]] = defaultdict(set)
    for r1, c1, rn, cn in rects:
        for r in range(r1, r1 + rn):
            for c in range(c1, c1 + cn):
                m[r].add(c)
    return {r: sorted(m[r]) for r in sorted(m.keys())}


def _bbox_union_rects(rects: list[tuple[int, int, int, int]]) -> tuple[int, int, int, int]:
    """包含矩形を (y1, x1, yn, xn) で返す。"""
    r1m = min(t[0] for t in rects)
    c1m = min(t[1] for t in rects)
    rmax = max(t[0] + t[2] - 1 for t in rects)
    cmax = max(t[1] + t[3] - 1 for t in rects)
    return r1m, c1m, rmax - r1m + 1, cmax - c1m + 1


def _full_sheet_hint_from_bridge_rects(
    ptr_s: Any,
    rects: list[tuple[int, int, int, int]],
) -> tuple[bool, str]:
    """
    bridge 由来の交差矩形から「シート全体グリッド選択相当」を推定（COM の Selection は読まない）。
    """
    if not rects:
        return False, "bridge_no_rects"
    try:
        api_sh = ptr_s.api
        max_r = int(api_sh.Rows.Count)
        max_c = int(api_sh.Columns.Count)
        rb, cb, ryn, cxn = _bbox_union_rects(rects)
        if rb <= 1 and cb <= 1 and ryn >= max_r and cxn >= max_c:
            return True, "bridge_union_covers_sheet_grid"
    except Exception as exc:
        return False, f"bridge_full_sheet_exc={exc!r}"
    return False, "bridge_not_full_sheet_grid"


def _submatrix_from_used(
    arr_used: list[list[Any]], uy1: int, ux1: int, y1: int, x1: int, yn: int, xn: int
) -> list[list[Any]]:
    r0, c0 = y1 - uy1, x1 - ux1
    out: list[list[Any]] = []
    for i in range(yn):
        src = arr_used[r0 + i] if 0 <= r0 + i < len(arr_used) else []
        seg: list[Any] = []
        for j in range(xn):
            ci = c0 + j
            seg.append(src[ci] if 0 <= ci < len(src) else None)
        out.append(seg)
    return out


def _merge_cols_to_row_rects(sheet_row: int, cols: list[int]) -> list[list[int]]:
    """同一行の列番号リストを [r,c1,r,c2] 連続区間にまとめる（シート座標・両端含む）。"""
    if not cols:
        return []
    rects: list[list[int]] = []
    s = e = cols[0]
    for c in cols[1:]:
        if c == e + 1:
            e = c
        else:
            rects.append([sheet_row, s, sheet_row, e])
            s = e = c
    rects.append([sheet_row, s, sheet_row, e])
    return rects


def _highlight_rects_mode_b(
    positions: list[tuple[int, int]],
    *,
    cancel_check: Optional[Callable[..., None]] = None,
    analyze_emit: Optional[_DupliAnalyzeProgressEmitter] = None,
    emit_lo: int = 86,
    emit_hi: int = 92,
) -> list[list[int]]:
    npos = len(positions)
    if analyze_emit is not None and npos == 0:
        analyze_emit.tick(1, 1, emit_lo, emit_hi)
        return []
    by_row: dict[int, list[int]] = defaultdict(list)
    for i, (r, c) in enumerate(positions, start=1):
        if cancel_check is not None:
            cancel_check()
        by_row[r].append(c)
        if analyze_emit is not None and npos > 0:
            analyze_emit.tick(i, npos, emit_lo, emit_hi)
    out: list[list[int]] = []
    for r in sorted(by_row):
        if cancel_check is not None:
            cancel_check()
        out.extend(_merge_cols_to_row_rects(r, sorted(set(by_row[r]))))
    return out


def _dup_sheet_rows_ordered_mode_a(
    df: pd.DataFrame,
    ser_dup: pd.Series,
    *,
    cancel_check: Optional[Callable[..., None]] = None,
    analyze_emit: Optional[_DupliAnalyzeProgressEmitter] = None,
    emit_lo: int = 62,
    emit_hi: int = 74,
) -> list[int]:
    """モード A: 重複行のシート行番号を、同一レコードグループは先頭行の若い順で並べる。"""
    dup_df = df.loc[ser_dup]
    if dup_df.empty:
        return []
    parts: list[tuple[int, list[int]]] = []
    try:
        gb = dup_df.groupby("rec", dropna=False)
        ngrp = int(gb.ngroups)
        for gi, (_, sub) in enumerate(gb, start=1):
            if cancel_check is not None:
                cancel_check()
            if analyze_emit is not None and ngrp > 0:
                analyze_emit.tick(gi, ngrp, emit_lo, emit_hi)
            if len(sub) < 2:
                continue
            rows = sorted(int(x) for x in sub["sheet_row"].tolist())
            parts.append((rows[0], rows))
    except (TypeError, ValueError):
        return sorted(int(x) for x in dup_df["sheet_row"].tolist())
    parts.sort(key=lambda t: t[0])
    return [r for _, lst in parts for r in lst]


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


def _read_sheet_matrix(
    ptr_s: Any,
    y1: int,
    x1: int,
    yn: int,
    xn: int,
    on_pct: Callable[[int, str, str], None],
    cfg: dict[str, Any],
    cancel_check: Optional[Callable[..., None]] = None,
) -> Optional[list[list[Any]]]:
    """
    シートの指定範囲 (y1,x1) から yn×xn をチャンク単位で読み、2 次元リストで返す。
    読込中は on_pct で進捗コールバックを呼ぶ。失敗時は None。
    cancel_check が指定され、呼び出しで例外が出た場合はそのまま伝播（協力的中止用）。
    チャンク先頭では force=True で呼ぶ（長い Range 読込に入る前に必ずキャンセル確認）。
    """
    msg_read = _msg(cfg, "PHASE_READ")
    custom = (cfg.get("MESSAGES") or {}).get("PROGRESS_CUSTOM_READ") or "読込中"
    chunk_rows = max(200, min(5000, yn))
    acc: list[list[Any]] = []
    try:
        for r0 in range(0, yn, chunk_rows):
            if cancel_check is not None:
                cancel_check(force=True)
            r1 = min(r0 + chunk_rows, yn)
            pct = int(5 + (r1 / max(yn, 1)) * 35)
            on_pct(pct, msg_read, custom)
            rng = ptr_s.range((y1 + r0, x1), (y1 + r1 - 1, x1 + xn - 1))
            part = rng.value
            sub = _normalize_2d(part, r1 - r0, xn)
            acc.extend(sub)
        return acc if len(acc) == yn else None
    except _DupliCancelled:
        raise
    except Exception:
        return None


def _svc_dupli_clear_range_fill(ptr_s: Any, r1: int, c1: int, r2: int, c2: int) -> None:
    """ui_dupli._dupli_clear_range_fill と同等の Interior クリア（xlwings Range）。"""
    rng = ptr_s.range((r1, c1), (r2, c2))
    try:
        rng.color = None
    except Exception:
        pass
    try:
        rng.api.Interior.Pattern = -4142  # xlNone
    except Exception:
        pass
    try:
        rng.api.Interior.TintAndShade = 0
    except Exception:
        pass
    for idx in (-4105, -4146):
        try:
            rng.api.Interior.ColorIndex = idx
        except Exception:
            pass


def _svc_clear_applied_highlight_rects(
    ptr_s: Any,
    ptr_a: Any,
    quads: list[list[int]],
    prog_path: Path,
    cfg: dict[str, Any],
    seq_ref: list[int],
) -> None:
    """協力キャンセル時: 既に着色した矩形だけクリアし、進捗を更新する。"""
    if not quads:
        return
    api_app_ptr = ptr_a.api
    prev_screen = True
    prev_calc: Any = _XL_CALC_AUTOMATIC
    msg_clear = _msg(cfg, "PHASE_CLEAR")
    total = len(quads)
    try:
        prev_screen = api_app_ptr.ScreenUpdating
        prev_calc = api_app_ptr.Calculation
        api_app_ptr.ScreenUpdating = False
        api_app_ptr.Calculation = _XL_CALC_MANUAL
        for i, quad in enumerate(quads):
            if len(quad) < 4:
                continue
            r1, c1, r2, c2 = int(quad[0]), int(quad[1]), int(quad[2]), int(quad[3])
            _svc_dupli_clear_range_fill(ptr_s, r1, c1, r2, c2)
            if (i + 1) % 40 == 0 or (i + 1) == total:
                seq_ref[0] += 1
                done = i + 1
                _progress_write(
                    prog_path,
                    {
                        "status": "RUN",
                        "hide_cancel_button": True,
                        "phase": msg_clear,
                        "msg": msg_clear,
                        "pct": int(100 * done / max(total, 1)),
                        "done": done,
                        "total": total,
                        "seq": seq_ref[0],
                    },
                )
    finally:
        try:
            api_app_ptr.Calculation = prev_calc
        except Exception:
            pass
        try:
            api_app_ptr.ScreenUpdating = prev_screen
        except Exception:
            pass


def _handle_dupli_cancelled(
    ptr_w: Any,
    ptr_s: Any,
    ptr_a: Any,
    prog_path: Path,
    cfg: dict[str, Any],
    hl_applied: list[list[int]],
    ph: int,
    seq_ref: list[int],
    t_flow: float,
) -> None:
    logger.info("[DUPLI] cooperative cancel hl_applied=%s", len(hl_applied))
    _perf_dupli("cancelled", t_flow, hl_applied=len(hl_applied))
    _dupli_trace("cancelled", t_flow, hl_applied=len(hl_applied))
    if hl_applied:
        seq_ref[0] += 1
        _progress_write(
            prog_path,
            {
                "status": "RUN",
                "hide_cancel_button": True,
                "phase": _msg(cfg, "PHASE_CLEAR"),
                "msg": _msg(cfg, "PHASE_CLEAR"),
                "pct": 0,
                "done": 0,
                "total": len(hl_applied),
                "seq": seq_ref[0],
            },
        )
        _svc_clear_applied_highlight_rects(ptr_s, ptr_a, hl_applied, prog_path, cfg, seq_ref)
    done_payload: dict[str, Any] = {
        "status": "DONE",
        "show_done_dialog": False,
        "phase": _msg(cfg, "CANCELLED_DONE"),
        "seq": 999,
    }
    if not hl_applied:
        done_payload["done_delay_ms"] = int(_CANCEL_DONE_DELAY_MS_NO_HIGHLIGHT)
    _progress_write(prog_path, done_payload)
    _status_bar_set(ptr_w, _msg(cfg, "STATUS_CANCELLED"))


def _highlight_bgr(cfg: dict[str, Any]) -> int:
    """
    重複セルの背景色を Excel の BGR 整数で返す。
    HIGHLIGHT.USE_CORE_ERR_BG が真なら core_cst.ERR_BG_COLOR、でなければ HIGHLIGHT.RGB または既定の薄赤。
    """
    hi = cfg.get("HIGHLIGHT") or {}
    if hi.get("USE_CORE_ERR_BG") and cst is not None:
        t = getattr(cst, "ERR_BG_COLOR", (255, 200, 200))
        if isinstance(t, (list, tuple)) and len(t) >= 3:
            r, g, b = int(t[0]), int(t[1]), int(t[2])
            return r + (g * 256) + (b * 65536)
    rgb = hi.get("RGB")
    if isinstance(rgb, (list, tuple)) and len(rgb) >= 3:
        r, g, b = int(rgb[0]), int(rgb[1]), int(rgb[2])
        return r + (g * 256) + (b * 65536)
    return 180 + (200 * 256) + (255 * 65536)


def _sheet_id_resolve(ptr_s: Any, sheet_id: str) -> str:
    """
    進捗・レポート用のシート識別子を返す。
    sheet_id が空の場合はシートの HC_GUID_B64 を取得し、無ければオブジェクト id ベースのフォールバックを使う。
    """
    s = str(sheet_id or "").strip()
    if s:
        return s
    if core_xlc_mod is not None:
        try:
            g = core_xlc_mod.get_sheet_prop(ptr_s, "HC_GUID_B64")
            if g:
                return str(g)[:48]
        except Exception:
            pass
    return f"dupli_{abs(id(ptr_s))}"


def check_duplicates(
    target_hwnd: Optional[int] = None,
    sheet_id: str = "",
    selection_areas: Optional[list[str]] = None,
    selection_count_large: Optional[int] = None,
    sheet_cells_count_large: Optional[int] = None,
) -> None:
    """
    【概要】
        指定 HWND の Excel ブック・シートにおいて、選択範囲の重複を検出し、着色とレポート表示を行う。
    【補足】
        進捗・完了通知・レポートは ui_server 経由で ui_dupli に依頼する。設定は config/ui_dupli.json。
        bridge から渡す selection_areas（External 付き Area ごとの文字列）があれば、まずそれで交差矩形を構築する。
        selection_count_large / sheet_cells_count_large は VBA がリボン内で読んだ CountLarge（全シート判定用・任意）。
    """
    t_flow = time.perf_counter()
    _perf_dupli("enter", t_flow)
    _dupli_trace("enter", t_flow)

    if core_xlc_mod is None:
        logger.error("[DUPLI] core_xlc not available")
        _perf_dupli("abort_no_core_xlc", t_flow)
        _dupli_trace("abort_no_core_xlc", t_flow)
        return
    ctx = core_xlc_mod.get_excel_context_from_hwnd(int(target_hwnd or 0), sheet_id)
    if ctx is None:
        logger.error("[DUPLI] Excel context not available (xlwings + HWND)")
        _perf_dupli("abort_no_context", t_flow)
        _dupli_trace("abort_no_context", t_flow)
        return

    ptr_a, ptr_w, ptr_s, ph = ctx
    logger.info("[DUPLI] 開始")
    _perf_dupli("after_context", t_flow, hwnd=ph)
    _dupli_trace("after_context", t_flow, hwnd=ph)
    cfg = _cfg()
    try:
        _vp0 = cfg.get("VIEWPORT_HIGHLIGHT") or {}
        _stale = float(_vp0.get("SIDECAR_STALE_SEC", 86400) or 86400)
        _cleanup_stale_dupli_hl_rect_sidecars(max_age_sec=max(3600.0, _stale))
    except Exception:
        _cleanup_stale_dupli_hl_rect_sidecars()
    saved_status = _status_bar_save(ptr_w)  # 終了時に必ず復元
    sid = _sheet_id_resolve(ptr_s, sheet_id)
    prog_path = _progress_path(sid)
    cancel_path = _cancel_request_path(sid)
    progress_closed_path = _progress_closed_ack_path(sid)
    _reset_cancel_path(progress_closed_path)
    seq = [0]  # 進捗の表示順序用
    skip_status_restore = False
    hl_applied: list[list[int]] = []

    def _submit_done_after_progress_close(message: str, title: str) -> None:
        _wait_progress_closed_ack(progress_closed_path)
        _submit_done_ui(ph, sid, message, title)

    def _upd(pct: int, phase: str, cur: str) -> None:
        seq[0] += 1
        _progress_write(
            prog_path,
            {
                "status": "RUN",
                "phase": phase,
                "msg": phase,
                "current_file": cur,
                "pct": max(0, min(100, pct)),
                "done": pct,
                "total": 100,
                "seq": seq[0],
            },
        )

    sa_normalized: Optional[list[str]] = None
    if selection_areas is not None and isinstance(selection_areas, list):
        sa_normalized = [str(x).strip() for x in selection_areas if str(x).strip()]
        if not sa_normalized:
            sa_normalized = None

    sel_cl: Optional[int] = None
    sheet_cl: Optional[int] = None
    if selection_count_large is not None:
        try:
            n = int(selection_count_large)
            sel_cl = n if n >= 0 else None
        except (TypeError, ValueError):
            sel_cl = None
    if sheet_cells_count_large is not None:
        try:
            n = int(sheet_cells_count_large)
            sheet_cl = n if n >= 0 else None
        except (TypeError, ValueError):
            sheet_cl = None

    rects_source = "com"
    rects: Optional[list[tuple[int, int, int, int]]] = None
    try:
        if sa_normalized:
            rects = _selection_used_rects_from_bridge_addresses(ptr_w, ptr_s, ptr_a, sa_normalized)
            if rects:
                rects_source = "bridge"
            else:
                try:
                    _dupli_diag.info(
                        "[DUPLI_SEL] bridge selection_areas produced no rects; falling back to COM selection"
                    )
                except Exception:
                    pass
                rects = _selection_used_rects(ptr_a, ptr_s)
                rects_source = "com_fallback"
        else:
            rects = _selection_used_rects(ptr_a, ptr_s)

        mode_b_full_sheet_no_rects = False
        corner_fs_precalc: Optional[tuple[bool, str]] = None
        if not rects:
            cfs, cwhy = _corner_full_sheet_flags(ptr_a, ptr_s, sel_cl, sheet_cl)
            corner_fs_precalc = (cfs, cwhy)
            if cfs:
                mode_b_full_sheet_no_rects = True
                logger.info(
                    "[DUPLI] UsedRange と交差なしだが全シート選択相当 → モードB（%s）",
                    cwhy,
                )
                _dupli_sel_log(
                    "full_sheet_no_intersection_proceed_b",
                    t_flow,
                    had_bridge_areas=bool(sa_normalized),
                    **_dupli_selection_diag_snapshot(ptr_a, ptr_s),
                )
            else:
                logger.info("[DUPLI] 重複チェック範囲 範囲なし")
                _dupli_sel_log(
                    "early_no_valid_range",
                    t_flow,
                    had_bridge_areas=bool(sa_normalized),
                    **_dupli_selection_diag_snapshot(ptr_a, ptr_s),
                )
                _perf_dupli("early_no_valid_range", t_flow)
                _dupli_trace("early_no_valid_range", t_flow)
                _status_bar_set(ptr_w, _msg(cfg, "NO_VALID_RANGE"))
                done_cfg = (cfg.get("SCREENS") or {}).get("DONE") or {}
                _submit_done_after_progress_close(_msg(cfg, "NO_VALID_RANGE"), str(done_cfg.get("TITLE") or "重複チェック"))
                return

        def _read_used_bounds() -> tuple[Any, int, int, int, int]:
            u = ptr_s.used_range
            return u, int(u.row), int(u.column), int(u.rows.count), int(u.columns.count)

        try:
            ptr_usd, uy1, ux1, uyn, uxn = _excel_busy_retry(_read_used_bounds)
        except Exception as exc:
            logger.warning("[DUPLI] UsedRange 取得失敗（リトライ後）: %s", exc)
            _perf_dupli("abort_used_range_busy", t_flow)
            _dupli_trace("abort_used_range_busy", t_flow, exc=repr(exc))
            _progress_write(prog_path, {"status": "DONE", "seq": 999})
            _status_bar_set(ptr_w, _msg(cfg, "READ_FAILED"))
            done_cfg = (cfg.get("SCREENS") or {}).get("DONE") or {}
            _submit_done_after_progress_close(_msg(cfg, "READ_FAILED"), str(done_cfg.get("TITLE") or "重複チェック"))
            return

        bridge_cnt_fs, _bridge_cnt_why = _full_sheet_from_bridge_count_large(sel_cl, sheet_cl)
        if corner_fs_precalc is not None:
            corner_fs, corner_why = corner_fs_precalc
        else:
            corner_fs, corner_why = _corner_full_sheet_flags(ptr_a, ptr_s, sel_cl, sheet_cl)
        bridge_fs, bridge_why = False, ""
        if rects_source == "bridge" and rects:
            bridge_fs, bridge_why = _full_sheet_hint_from_bridge_rects(ptr_s, rects)
        mode_b_full_sheet = mode_b_full_sheet_no_rects or corner_fs or bridge_fs
        mode_b = mode_b_full_sheet
        if bridge_cnt_fs:
            _fs_why = _bridge_cnt_why
        elif corner_fs:
            _fs_why = corner_why
        elif bridge_fs:
            _fs_why = bridge_why
        else:
            _fs_why = corner_why
        logger.info(
            "[DUPLI] mode=%s selection_rects=%s used y1=%s x1=%s yn=%s xn=%s full_sheet=%s full_sheet_no_rects=%s bridge_cnt=%s corner_sel=%s bridge_grid=%s full_sheet_why=%s",
            "B" if mode_b else "A",
            rects,
            uy1,
            ux1,
            uyn,
            uxn,
            mode_b_full_sheet,
            mode_b_full_sheet_no_rects,
            bridge_cnt_fs,
            corner_fs,
            bridge_fs,
            _fs_why,
        )
        if rects:
            _ix_extra = dict(_dupli_intersection_log_extra(ptr_a, ptr_s, rects_source, rects))
            _ix_extra.pop("full_sheet", None)
            _ix_extra.pop("full_sheet_why", None)
            _dupli_sel_log(
                "after_intersection_ok",
                t_flow,
                mode_b=mode_b,
                full_sheet=mode_b_full_sheet,
                full_sheet_why=_fs_why,
                **_ix_extra,
            )
        else:
            _ix_extra_b = dict(_dupli_intersection_log_extra(ptr_a, ptr_s, rects_source, rects))
            _ix_extra_b.pop("full_sheet_why", None)
            _dupli_sel_log(
                "after_intersection_empty_mode_b_path",
                t_flow,
                mode_b_full_sheet_no_rects=mode_b_full_sheet_no_rects,
                full_sheet_why=_fs_why,
                **_ix_extra_b,
            )
        _perf_dupli(
            "after_intersection",
            t_flow,
            mode_b=mode_b,
            areas=len(rects) if rects else 0,
        )
        _dupli_trace(
            "after_intersection",
            t_flow,
            mode_b=mode_b,
            areas=len(rects) if rects else 0,
        )

        _reset_cancel_path(cancel_path)
        _progress_write(
            prog_path,
            {
                "status": "RUN",
                "phase": _msg(cfg, "PHASE_READ"),
                "pct": 0,
                "done": 0,
                "total": 100,
                "seq": 0,
            },
        )
        _submit_progress_ui(
            ph,
            sid,
            prog_path,
            3,
            cancel_request_path=cancel_path,
            progress_closed_path=progress_closed_path,
        )
        time.sleep(0.25)
        _perf_dupli("after_progress_ui_submit", t_flow)
        _dupli_trace("after_progress_ui_submit", t_flow)

        _chk = _make_time_gated_cancel_check(cancel_path)

        arr_used = _read_sheet_matrix(ptr_s, uy1, ux1, uyn, uxn, _upd, cfg, cancel_check=_chk)
        if arr_used is None:
            logger.warning("[DUPLI] 読込失敗")
            _perf_dupli("abort_matrix_read_failed", t_flow)
            _dupli_trace("abort_matrix_read_failed", t_flow)
            _progress_write(prog_path, {"status": "DONE", "seq": 999})
            _status_bar_set(ptr_w, _msg(cfg, "READ_FAILED"))
            done_cfg = (cfg.get("SCREENS") or {}).get("DONE") or {}
            _submit_done_after_progress_close(_msg(cfg, "READ_FAILED"), str(done_cfg.get("TITLE") or "重複チェック"))
            return

        _perf_dupli("after_matrix_read", t_flow, yn=uyn, xn=uxn)
        _dupli_trace("after_matrix_read", t_flow, yn=uyn, xn=uxn)

        _chk(force=True)

        rep_cfg = ((cfg.get("SCREENS") or {}).get("REPORT") or {})
        count_caption_tpl = str(rep_cfg.get("COUNT_CAPTION_TEMPLATE") or "検出総数: {count} 件")
        report_title = str(rep_cfg.get("TITLE_TEMPLATE") or "重複検出レポート").strip()

        color_int = _highlight_bgr(cfg)
        rep_rows: list[list[str]] = []
        addrs: list[str] = []
        hl_rects: list[list[int]] = []
        n = 0
        scan_units = uyn
        report_intro = ""
        cols: list[dict[str, Any]] = []

        msg_w = _msg(cfg, "PHASE_WRITE")
        cw = (cfg.get("MESSAGES") or {}).get("PROGRESS_CUSTOM_WRITE") or ""

        _msg_an = _msg(cfg, "PHASE_ANALYZE")
        _custom_an = str((cfg.get("MESSAGES") or {}).get("PROGRESS_CUSTOM_ANALYZE") or "")
        _upd(45, _msg_an, _custom_an)
        _emit = _DupliAnalyzeProgressEmitter(_upd, _msg_an, _custom_an, start_pct=45)

        if mode_b:
            vy1, vx1, vyn, vxn, arr_val = _trim_used_matrix(arr_used, uy1, ux1)
            scan_units = max(1, vyn * vxn)
            if vyn < 1 or vxn < 1:
                _progress_write(prog_path, {"status": "DONE", "seq": 999})
                logger.info("[DUPLI] モードB 有効データ領域なし")
                _status_bar_set(ptr_w, _msg(cfg, "NO_VALID_RANGE"))
                done_cfg = (cfg.get("SCREENS") or {}).get("DONE") or {}
                _submit_done_after_progress_close(_msg(cfg, "NO_VALID_RANGE"), str(done_cfg.get("TITLE") or "重複チェック"))
                return

            keyed: list[tuple[int, int, str]] = []
            total_cells = max(1, vyn * vxn)
            seen_c = 0
            _emit.set_stride(_DUPLI_ANALYZE_STRIDE_CELLS)
            for ir in range(vyn):
                for ic in range(vxn):
                    _chk()
                    k = _norm_dup_cell(arr_val[ir][ic])
                    keyed.append((vy1 + ir, vx1 + ic, k))
                    seen_c += 1
                    _emit.tick(seen_c, total_cells, 46, 72)
            _chk(force=True)
            cnt = Counter(t[2] for t in keyed)
            dup_keys = {k for k, v in cnt.items() if v >= 2}
            cells_dup = [(r, c, k) for r, c, k in keyed if k in dup_keys]
            n = len(cells_dup)
            logger.info("[DUPLI] モードB 重複有り無し=%s 重複セル数=%s", "重複有り" if n else "重複無し", n)
            _perf_dupli("after_analyze", t_flow, dup_count=n, mode="B", cells=vyn * vxn)
            _dupli_trace("after_analyze", t_flow, dup_count=n, mode="B")

            if n == 0:
                _progress_write(prog_path, {"status": "DONE", "seq": 999})
                logger.info("[DUPLI] 完了 モードB 走査=%s 重複=0", scan_units)
                _status_bar_set(ptr_w, _msg(cfg, "NO_DUPLICATES"))
                done_cfg = (cfg.get("SCREENS") or {}).get("DONE") or {}
                _submit_done_after_progress_close(_msg(cfg, "NO_DUPLICATES"), str(done_cfg.get("TITLE") or "重複チェック"))
                _perf_dupli("early_no_duplicates", t_flow)
                _dupli_trace("early_no_duplicates", t_flow)
                return

            groups: dict[str, list[tuple[int, int]]] = defaultdict(list)
            ndup = len(cells_dup)
            _emit.set_stride(_DUPLI_ANALYZE_STRIDE_CELLS)
            for i_cell, (r, c, k) in enumerate(cells_dup, start=1):
                _chk()
                groups[k].append((r, c))
                _emit.tick(i_cell, ndup, 72, 78)
            parts: list[tuple[tuple[int, int], str, int]] = []
            nkeys = len(groups)
            _emit.set_stride(_DUPLI_ANALYZE_STRIDE_GROUPS)
            for ik, (k, pos) in enumerate(groups.items(), start=1):
                _chk()
                pos.sort()
                parts.append((pos[0], k, len(pos)))
                _emit.tick(ik, max(1, nkeys), 78, 82)
            parts.sort(key=lambda t: (t[0][0], t[0][1]))
            nparts = len(parts)
            _emit.set_stride(_DUPLI_ANALYZE_STRIDE_ROWS)
            for irp, (first, k, ln) in enumerate(parts, start=1):
                _chk()
                r, c = first
                cell = ptr_s.range((r, c))
                addr = cell.address.replace("$", "")
                rep_rows.append([k, str(ln), addr])
                addrs.append(addr)
                _emit.tick(irp, max(1, nparts), 82, 86)
            _chk(force=True)
            _emit.set_stride(_DUPLI_ANALYZE_STRIDE_HL)
            hl_rects = _highlight_rects_mode_b(
                [(r, c) for r, c, _ in cells_dup],
                cancel_check=_chk,
                analyze_emit=_emit,
                emit_lo=86,
                emit_hi=92,
            )
            cols = list(
                rep_cfg.get("MODE_B_COLUMNS")
                or [
                    {"key": "value", "label": "値（比較キー）", "width": 0},
                    {"key": "ncells", "label": "該当セル数", "width": 0},
                    {"key": "addr", "label": "代表座標", "width": 0},
                ]
            )
            report_intro = str(rep_cfg.get("REPORT_INTRO_MODE_B") or rep_cfg.get("REPORT_INTRO") or "").strip()
        else:
            rb1, cb1, ryn, cxn = _bbox_union_rects(rects)
            sub = _submatrix_from_used(arr_used, uy1, ux1, rb1, cb1, ryn, cxn)
            sel_pr = _sel_cols_per_row_from_rects(rects)
            all_rows = sorted(sel_pr.keys())
            if not all_rows:
                _progress_write(prog_path, {"status": "DONE", "seq": 999})
                _status_bar_set(ptr_w, _msg(cfg, "NO_VALID_RANGE"))
                done_cfg = (cfg.get("SCREENS") or {}).get("DONE") or {}
                _submit_done_after_progress_close(_msg(cfg, "NO_VALID_RANGE"), str(done_cfg.get("TITLE") or "重複チェック"))
                return

            recs: list[tuple[str, ...]] = []
            srows: list[int] = []
            n_all = len(all_rows)
            _emit.set_stride(_DUPLI_ANALYZE_STRIDE_ROWS)
            for i_r, r in enumerate(all_rows, start=1):
                _chk()
                cols_sel = sel_pr[r]
                tup = tuple(_norm_dup_cell(sub[r - rb1][c - cb1]) for c in cols_sel)
                recs.append(tup)
                srows.append(r)
                _emit.tick(i_r, max(1, n_all), 46, 61)
            _chk(force=True)
            df = pd.DataFrame({"rec": recs, "sheet_row": srows})
            ser_dup = df.duplicated(subset=["rec"], keep=False)
            _chk(force=True)
            n = int(ser_dup.sum())
            scan_units = len(all_rows)
            logger.info("[DUPLI] モードA 重複有り無し=%s 重複行数=%s", "重複有り" if n else "重複無し", n)
            _perf_dupli("after_analyze", t_flow, dup_count=n, mode="A", rows=scan_units)
            _dupli_trace("after_analyze", t_flow, dup_count=n, mode="A")

            if n == 0:
                _progress_write(prog_path, {"status": "DONE", "seq": 999})
                logger.info("[DUPLI] 完了 モードA 走査=%s 重複=0", scan_units)
                _status_bar_set(ptr_w, _msg(cfg, "NO_DUPLICATES"))
                done_cfg = (cfg.get("SCREENS") or {}).get("DONE") or {}
                _submit_done_after_progress_close(_msg(cfg, "NO_DUPLICATES"), str(done_cfg.get("TITLE") or "重複チェック"))
                _perf_dupli("early_no_duplicates", t_flow)
                _dupli_trace("early_no_duplicates", t_flow)
                return

            _emit.set_stride(_DUPLI_ANALYZE_STRIDE_GROUPS)
            idx_ordered = _dup_sheet_rows_ordered_mode_a(
                df,
                ser_dup,
                cancel_check=_chk,
                analyze_emit=_emit,
                emit_lo=62,
                emit_hi=74,
            )
            dup_set = set(int(x) for x in df.loc[ser_dup, "sheet_row"].tolist())
            cols = list(
                rep_cfg.get("COLUMNS")
                or [
                    {"key": "row", "label": "行", "width": 100},
                    {"key": "addr", "label": "座標", "width": 120},
                    {"key": "summary", "label": "内容", "width": 400},
                ]
            )
            report_intro = str(rep_cfg.get("REPORT_INTRO") or "").strip()
            n_idx = len(idx_ordered)
            _emit.set_stride(_DUPLI_ANALYZE_STRIDE_ROWS)
            for ij, sr in enumerate(idx_ordered, start=1):
                _chk()
                cols_sel = sel_pr[sr]
                raw3 = [sub[sr - rb1][c - cb1] for c in cols_sel[:3]]
                summ = " | ".join(
                    "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v) for v in raw3
                )
                c0 = cols_sel[0]
                addr = ptr_s.range((sr, c0)).address.replace("$", "")
                rep_rows.append([f"{sr}行目", addr, summ])
                addrs.append(addr)
                _emit.tick(ij, max(1, n_idx), 75, 85)
            dup_sorted = sorted(dup_set)
            _emit.set_stride(_DUPLI_ANALYZE_STRIDE_ROWS)
            for jh, r in enumerate(dup_sorted, start=1):
                _chk()
                hl_rects.extend(_merge_cols_to_row_rects(r, sel_pr[r]))
                _emit.tick(jh, max(1, len(dup_sorted)), 86, 92)

        _chk(force=True)

        sidecar_path = _write_dupli_hl_rects_sidecar(sid, hl_rects)
        _upd(94, msg_w, cw)

        _perf_dupli(
            "after_duplicate_highlight",
            t_flow,
            dup_count=n,
            viewport_sidecar=1,
            hl_rects_n=len(hl_rects),
        )
        _dupli_trace(
            "after_duplicate_highlight",
            t_flow,
            dup_count=n,
            viewport_sidecar=1,
            hl_rects_n=len(hl_rects),
        )

        _progress_write(
            prog_path,
            {
                "status": "DONE",
                "show_done_dialog": False,
                "seq": 999,
            },
        )

        highlight_clear: dict[str, Any] = {
            "rects_path": str(sidecar_path),
            "rects_count": len(hl_rects),
            "viewport_follow": True,
            "fill_bgr": int(color_int),
            "sheet_name": _sheet_name_for_highlight(ptr_s),
            "book_name": _book_name_for_highlight(ptr_s),
        }
        logger.info("[DUPLI] 完了 走査単位=%s 重複カウント=%s", scan_units, n)
        _status_bar_set(ptr_w, _msg(cfg, "STATUS_FINAL", count=n))
        done_cfg = (cfg.get("SCREENS") or {}).get("DONE") or {}
        _wait_progress_closed_ack(progress_closed_path)
        _submit_done_ui(
            ph,
            sid,
            _msg(cfg, "STATUS_FINAL", count=n),
            str(done_cfg.get("TITLE") or "重複チェック"),
        )
        _submit_report_ui(
            ph,
            sid,
            report_title,
            cols,
            rep_rows,
            addrs,
            report_intro=report_intro,
            dup_count=n,
            count_caption_template=count_caption_tpl,
            highlight_clear=highlight_clear,
            link_col=2 if mode_b else 1,
        )
        _perf_dupli("after_report_submit", t_flow, dup_count=n)
        _dupli_trace("after_report_submit", t_flow, dup_count=n)

    except _DupliCancelled:
        skip_status_restore = True
        try:
            _handle_dupli_cancelled(ptr_w, ptr_s, ptr_a, prog_path, cfg, hl_applied, ph, seq, t_flow)
        except Exception as exc:
            logger.warning("[DUPLI] cancel handler failed: %s", exc)
            try:
                _progress_write(
                    prog_path,
                    {"status": "DONE", "show_done_dialog": False, "phase": _msg(cfg, "CANCELLED_DONE"), "seq": 999},
                )
            except Exception:
                pass
        return

    except Exception as ex:
        logger.exception("[DUPLI] %s", ex)
        try:
            _status_bar_set(ptr_w, f"{_msg(cfg, 'ERROR_PREFIX')}: {ex}")
        except Exception:
            pass
    finally:
        # 正常・異常問わず: Interactive 復元、進捗完了、ステータスバー復元、Excel を前面に
        try:
            ptr_a.api.Interactive = True
        except Exception:
            pass
        try:
            _progress_write(prog_path, {"status": "DONE", "seq": 999})
        except Exception:
            pass
        try:
            if not skip_status_restore:
                _status_bar_restore(ptr_w, saved_status)
        except Exception:
            pass
        try:
            from core import core_w32

            core_w32.bring_to_front(ph)
        except Exception:
            pass
        _perf_dupli("flow_end", t_flow)
        _dupli_trace("flow_end", t_flow)
