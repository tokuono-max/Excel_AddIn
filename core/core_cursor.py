# -*- coding: utf-8 -*-
"""
Pythonバージョン: 3.12
モジュール名: core_cursor
作成日: 2026-02-12
更新日: 2026-06-06
バージョン: 0.3.5
概要:
    Excelアドインの「砂時計（Application.Cursor=xlWait）」を、Python側（UI表示完了）で解除する共通モジュール。
    VBA側で起動した保険タイマ（Application.OnTime）も、VBAマクロ呼び出しにより停止する。
    ※Polling（Application.OnTimeの定周期監視）は副作用（STA占有）を招きやすいため、本方式では採用しない。

改訂履歴:
    0.3.5: 2026-06-06 notify_wait_form_ready を .ready 合図ファイル方式に変更（COM Application.Run 廃止）。notify_ui_ready から WaitForm COM 解除を削除。
    0.3.4: 2026-06-06 data_agg/csv_tool の svc 側砂時計 API を削除（進捗は ProgressDialog のみ）。
    0.3.3: 2026-06-06 進捗砂時計は ProgressDialog の表示開始/終了のみ（ForceCursorOnProgress・保険タイマなし）。tick 系は互換のため no-op。
    0.3.2: 2026-06-04 progress_dialog_wait_cursor_on/tick/off（進捗ダイアログ表示中の砂時計を全機能で共通化）。
    0.3.1: 2026-06-04 csv_tool_wait_cursor_on/tick/off（CSV 保存・結合・分割・読込の処理中砂時計）。
    0.3.0: 2026-04-06 notify_wait_form_ready 追加（VBA WaitForm 解除）。notify_ui_ready 成功時も同時に呼ぶ。
    0.2.0: 2026-02-12 core.hc_log.get_logger に完全準拠（フォールバック削除）。ログ出力を統一。
    0.1.0: 2026-02-12 初版。
"""

from __future__ import annotations

from dataclasses import dataclass
import time
from typing import Any, Optional


# ==============================================================================
# Imports (project)
# ==============================================================================
# ルール:
#   - core.hc_log は必須依存（ユーザー指定: フォールバック不要）
#   - ロガー名は「モジュール階層」を含め、ログ解析で識別しやすくする
from core.core_log import get_logger  # noqa: E402


# ==============================================================================
# Constants
# ==============================================================================
# Excel 定数: xlDefault = -4143, xlWait = -4112
_XL_CURSOR_DEFAULT: int = -4143
_XL_CURSOR_WAIT: int = -4112


# ==============================================================================
# Logger
# ==============================================================================
_LOG = get_logger("core.hc_cursor")

_VBA_CURSOR_GUARD_START: str = "Main.StartCursorGuardTimer"
_VBA_CURSOR_FORCE_ON: str = "Main.ForceCursorOn"
_VBA_CURSOR_PROGRESS_ON: str = "Main.ForceCursorOnProgress"

# ==============================================================================
# Data models
# ==============================================================================
@dataclass(frozen=True)
class CursorGuardResult:
    """notify_ui_ready の実行結果。

    Attributes:
        ok: True=目的（Cursor OFF と保険タイマ停止）の両方または片方が達成できた
        cursor_off_ok: Cursor OFF の送信に成功
        timer_cancel_ok: VBA保険タイマ停止（Excel.Run）の呼び出しに成功
        elapsed_ms: COM処理を含む実行時間（目安）
        error: 失敗時の例外文字列（成功時は空文字）
    """

    ok: bool
    cursor_off_ok: bool
    timer_cancel_ok: bool
    elapsed_ms: float
    error: str


# ==============================================================================
# Public API
# ==============================================================================
def notify_excel_wait_cursor_on(
    *,
    sheet_id: str = "",
    vba_force_on_macro: str = _VBA_CURSOR_FORCE_ON,
    vba_guard_macro: str = _VBA_CURSOR_GUARD_START,
    try_get_excel: int = 20,
    get_excel_interval_ms: int = 50,
) -> None:
    """Excel 砂時計（xlWait）を ON にし、VBA 保険タイマを再武装する（ベストエフォート）。

    外部プロセスからの Application.Cursor 直書きは Excel 側で拒否されることがあるため、
    原則 VBA の ForceCursorOn（Excel スレッド）を Application.Run で呼ぶ。
    """
    sid = str(sheet_id or "batch")
    try:
        import pythoncom  # noqa: WPS433
        import win32com.client  # noqa: WPS433
    except Exception as ex:
        _LOG.warning("CURSOR_WAIT_ON: pywin32 import failed: %s", str(ex))
        return

    pythoncom.CoInitialize()
    try:
        excel = _get_excel_app(
            win32com_client=win32com.client,
            try_count=try_get_excel,
            interval_ms=get_excel_interval_ms,
        )
        if excel is None:
            _LOG.warning("CURSOR_WAIT_ON: GetActiveObject(Excel.Application) failed")
            return
        try:
            excel.Run(vba_force_on_macro, sid)
            _LOG.info("CURSOR_WAIT_ON: Run(%s) ok sheet_id=%s", vba_force_on_macro, sid)
            return
        except Exception as ex:
            _LOG.warning(
                "CURSOR_WAIT_ON: Run(%s) failed: %s; fallback COM/guard",
                vba_force_on_macro,
                str(ex),
            )
        try:
            excel.Cursor = _XL_CURSOR_WAIT
            _LOG.info("CURSOR_WAIT_ON: Cursor xlWait (COM fallback) sheet_id=%s", sid)
        except Exception as ex:
            _LOG.warning("CURSOR_WAIT_ON: Cursor ON failed: %s", str(ex))
        if str(vba_guard_macro or "").strip():
            try:
                excel.Run(vba_guard_macro, sid)
                _LOG.info("CURSOR_WAIT_ON: Run(%s) ok", vba_guard_macro)
            except Exception as ex:
                _LOG.warning("CURSOR_WAIT_ON: guard timer Run failed: %s", str(ex))
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def progress_dialog_wait_cursor_on(sheet_id: str = "") -> None:
    """進捗ダイアログ表示開始: 砂時計 ON（保険タイマなし・明示指示のみ）。"""
    notify_excel_wait_cursor_on(
        sheet_id=sheet_id or "progress",
        vba_force_on_macro=_VBA_CURSOR_PROGRESS_ON,
        vba_guard_macro="",
    )


def progress_dialog_wait_cursor_tick(
    sheet_id: str = "",
    *,
    min_interval_sec: float = 7.0,
) -> None:
    """互換用 no-op。進捗中の再武装は行わない（表示開始/終了のみ制御）。"""
    del sheet_id, min_interval_sec


def progress_dialog_wait_cursor_off(
    *,
    cancel_reason: str = "progress_dialog_done",
    timeout_sec: float = 2.5,
) -> None:
    """進捗ダイアログ終了時: 砂時計 OFF + WaitForm 解除（Excel COM はタイムアウト付き）。"""
    if timeout_sec <= 0:
        notify_ui_ready(cancel_reason=cancel_reason)
        return
    import threading

    finished = threading.Event()

    def _run() -> None:
        try:
            notify_ui_ready(cancel_reason=cancel_reason)
        finally:
            finished.set()

    threading.Thread(target=_run, daemon=True).start()
    if not finished.wait(timeout=max(0.1, float(timeout_sec))):
        _LOG.warning(
            "UI_READY: progress_dialog_wait_cursor_off timed out after %.1fs reason=%s",
            timeout_sec,
            cancel_reason,
        )


def notify_wait_form_ready(
    *,
    parent_hwnd: int = 0,
    book: Any = None,
) -> None:
    """VBA WaitForUiReadySignal 向けに .ready 合図ファイルを書く（ベストエフォート）。

    parent_hwnd / book.app.hwnd / HC_EXCEL_HWND の順で HWND を解決する。
    """
    hwnd = int(parent_hwnd or 0)
    if hwnd <= 0 and book is not None:
        try:
            hwnd = int(getattr(book.app, "hwnd", 0) or 0)
        except Exception:
            hwnd = 0
    if hwnd <= 0:
        try:
            import os

            from core import core_env

            hwnd = int(os.environ.get(core_env.ENV_EXCEL_HWND, 0) or 0)
        except Exception:
            hwnd = 0
    if hwnd <= 0:
        _LOG.warning("WAITFORM_READY: skip (parent_hwnd unresolved)")
        return
    try:
        from ui_qt.ipc_file import write_waitform_ready_signal

        write_waitform_ready_signal(hwnd)
        _LOG.info("WAITFORM_READY: signal written hwnd=%s", hwnd)
    except Exception as ex:
        _LOG.warning("WAITFORM_READY: write failed: %s", str(ex))


def notify_ui_ready(
    *,
    vba_cancel_macro: str = "Main.CancelCursorGuardTimer",
    cancel_reason: str = "python_ui_ready",
    also_force_off_macro: Optional[str] = "Main.ForceCursorOff",
    try_get_excel: int = 20,
    get_excel_interval_ms: int = 50,
) -> CursorGuardResult:
    """UI表示完了をトリガに、砂時計解除とVBA保険タイマ停止を行う。

    設計意図（重要）:
      - ExcelはCOM STAで動作するため、外部スレッドからのCOM呼び出しは「通らない/遅延する」ことがある。
      - しかし本方式は「UI表示完了時の 1 回だけ通知」に限定する。
        これにより、VBA側の高頻度Polling（OnTime連打）でExcelメインスレッドを占有する事態を回避する。
      - 失敗してもVBA側の保険（OnTimeで ForceCursorOff）が残る前提で、UXを破壊しない。

    Args:
        vba_cancel_macro: VBA側の保険タイマ停止マクロ（モジュール名.メソッド名）
        cancel_reason: ログ解析用の理由文字列（VBAログに残す）
        also_force_off_macro: 念押しでVBA側の ForceCursorOff を呼ぶ場合のマクロ名（None で無効）
        try_get_excel: Excel.Application 取得リトライ回数（短時間・少回数に限定）
        get_excel_interval_ms: 取得リトライ間隔（ms）

    Returns:
        CursorGuardResult: 成功/失敗、所要時間、例外文字列
    """
    t0 = time.perf_counter()
    cursor_off_ok = False
    timer_cancel_ok = False
    err = ""

    _LOG.info(
        "UI_READY: start vba_cancel_macro=%s reason=%s", vba_cancel_macro, cancel_reason
    )

    try:
        import pythoncom  # noqa: WPS433
        import win32com.client  # noqa: WPS433
    except Exception as ex:
        err = "pywin32 import failed: " + str(ex)
        return _finish(t0, cursor_off_ok, timer_cancel_ok, err)

    pythoncom.CoInitialize()
    try:
        excel = _get_excel_app(
            win32com_client=win32com.client,
            try_count=try_get_excel,
            interval_ms=get_excel_interval_ms,
        )
        if excel is None:
            err = "GetActiveObject(Excel.Application) failed"
            return _finish(t0, cursor_off_ok, timer_cancel_ok, err)

        _LOG.info("UI_READY: Excel acquired")

        # 1) 砂時計解除（外部COMで直接変更）
        try:
            excel.Cursor = _XL_CURSOR_DEFAULT
            cursor_off_ok = True
            _LOG.info("UI_READY: Cursor OFF sent (xlDefault)")
        except Exception as ex:
            if not err:
                err = "Cursor OFF failed: " + str(ex)
            _LOG.error("UI_READY: Cursor OFF failed: %s", str(ex))

        # 2) VBA保険タイマ停止（VBA内で OnTime Schedule:=False を実行）
        try:
            excel.Run(vba_cancel_macro, cancel_reason)
            timer_cancel_ok = True
            _LOG.info("UI_READY: Guard timer cancel requested (%s)", vba_cancel_macro)
        except Exception as ex:
            if not err:
                err = "VBA cancel timer failed: " + str(ex)
            _LOG.error("UI_READY: VBA cancel timer failed: %s", str(ex))

        # 3) 念押し（任意）: VBA側のForceCursorOffも呼び、内部フラグ等を揃える
        if also_force_off_macro is not None:
            try:
                excel.Run(also_force_off_macro)
                _LOG.debug("UI_READY: VBA ForceCursorOff executed (optional)")
            except Exception as ex:
                _LOG.warning(
                    "UI_READY: VBA ForceCursorOff failed (ignored): %s", str(ex)
                )

    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            _LOG.warning("UI_READY: CoUninitialize failed (ignored)")

    return _finish(t0, cursor_off_ok, timer_cancel_ok, err)


# ==============================================================================
# Internal helpers
# ==============================================================================
def _get_excel_app(*, win32com_client, try_count: int, interval_ms: int):
    """Excel.Application を取得する（短時間リトライ）。

    なぜ短時間リトライか:
      - Excel起動直後や一時的なビジー状態で GetActiveObject が失敗することがある。
      - 長い待ちは体感を悪化させるため、短い間隔で少回数だけ試す。

    Args:
        win32com_client: win32com.client（注入可能にしてテストを容易化）
        try_count: リトライ回数
        interval_ms: リトライ間隔（ms）

    Returns:
        Excel.Application オブジェクト（取得失敗なら None）
    """
    i = 0
    max_try = max(1, int(try_count))
    sleep_sec = max(0.0, float(interval_ms) / 1000.0)

    while i < max_try:
        try:
            return win32com_client.GetActiveObject("Excel.Application")
        except Exception as ex:
            if i == 0:
                _LOG.info("UI_READY: GetActiveObject retry start: %s", str(ex))
            time.sleep(sleep_sec)
            i = i + 1

    return None


def _finish(
    t0: float, cursor_off_ok: bool, timer_cancel_ok: bool, err: str
) -> CursorGuardResult:
    """戻り値を整形し、最終ログを出す。"""
    elapsed_ms = (time.perf_counter() - t0) * 1000.0
    ok = cursor_off_ok or timer_cancel_ok

    if ok:
        _LOG.info(
            "UI_READY: done ok elapsed_ms=%.1f cursor_off_ok=%s timer_cancel_ok=%s",
            elapsed_ms,
            str(cursor_off_ok),
            str(timer_cancel_ok),
        )
    else:
        _LOG.error(
            "UI_READY: done ng elapsed_ms=%.1f cursor_off_ok=%s timer_cancel_ok=%s err=%s",
            elapsed_ms,
            str(cursor_off_ok),
            str(timer_cancel_ok),
            err,
        )

    return CursorGuardResult(
        ok=ok,
        cursor_off_ok=cursor_off_ok,
        timer_cancel_ok=timer_cancel_ok,
        elapsed_ms=elapsed_ms,
        error=err,
    )
