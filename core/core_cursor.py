# -*- coding: utf-8 -*-
"""
Pythonバージョン: 3.12
モジュール名: core_cursor
作成日: 2026-02-12
更新日: 2026-04-06
バージョン: 0.3.0
概要:
    Excelアドインの「砂時計（Application.Cursor=xlWait）」を、Python側（UI表示完了）で解除する共通モジュール。
    VBA側で起動した保険タイマ（Application.OnTime）も、VBAマクロ呼び出しにより停止する。
    ※Polling（Application.OnTimeの定周期監視）は副作用（STA占有）を招きやすいため、本方式では採用しない。

改訂履歴:
    0.3.0: 2026-04-06 notify_wait_form_ready 追加（VBA WaitForm 解除）。notify_ui_ready 成功時も同時に呼ぶ。
    0.2.0: 2026-02-12 core.hc_log.get_logger に完全準拠（フォールバック削除）。ログ出力を統一。
    0.1.0: 2026-02-12 初版。
"""

from __future__ import annotations

from dataclasses import dataclass
import time
from typing import Optional


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

# VBA: HC_WaitForm.NotifyUiReady（リボン待機 UserForm を閉じる）
_VBA_WAITFORM_NOTIFY_MACRO: str = "HC_WaitForm.NotifyUiReady"
_VBA_CURSOR_GUARD_START: str = "Main.StartCursorGuardTimer"
_VBA_CURSOR_FORCE_ON: str = "Main.ForceCursorOn"

_data_agg_batch_cursor_last_rearm: float = 0.0


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


def data_agg_batch_cursor_on(sheet_id: str = "") -> None:
    """本番一括: 砂時計 ON（進捗表示直前・再開時も呼ぶ）。"""
    global _data_agg_batch_cursor_last_rearm
    _data_agg_batch_cursor_last_rearm = time.perf_counter()
    notify_excel_wait_cursor_on(sheet_id=sheet_id)


def data_agg_batch_cursor_tick(
    sheet_id: str = "",
    *,
    min_interval_sec: float = 7.0,
) -> None:
    """本番一括: 10 秒保険タイマ切れ前に砂時計を再武装する。"""
    global _data_agg_batch_cursor_last_rearm
    now = time.perf_counter()
    if now - _data_agg_batch_cursor_last_rearm < float(min_interval_sec):
        return
    data_agg_batch_cursor_on(sheet_id)


def data_agg_batch_cursor_off(*, cancel_reason: str = "data_agg_batch_done") -> None:
    """本番一括完了・失敗・キャンセル時: 砂時計 OFF + WaitForm 解除。"""
    notify_ui_ready(cancel_reason=cancel_reason)


def notify_wait_form_ready(
    *,
    vba_macro: str = _VBA_WAITFORM_NOTIFY_MACRO,
    try_get_excel: int = 20,
    get_excel_interval_ms: int = 50,
) -> None:
    """Excel.Application.Run で WaitForm を閉じる VBA を実行する（ベストエフォート）。

    COM 取得・Run 失敗はログのみ。UI 表示完了など任意のタイミングから呼べる。
    """
    try:
        import pythoncom  # noqa: WPS433
        import win32com.client  # noqa: WPS433
    except Exception as ex:
        _LOG.warning("WAITFORM_NOTIFY: pywin32 import failed: %s", str(ex))
        return

    pythoncom.CoInitialize()
    try:
        excel = _get_excel_app(
            win32com_client=win32com.client,
            try_count=try_get_excel,
            interval_ms=get_excel_interval_ms,
        )
        if excel is None:
            _LOG.warning("WAITFORM_NOTIFY: GetActiveObject(Excel.Application) failed")
            return
        try:
            excel.Run(vba_macro)
            _LOG.info("WAITFORM_NOTIFY: Run(%s) ok", vba_macro)
        except Exception as ex:
            _LOG.warning("WAITFORM_NOTIFY: Run failed: %s", str(ex))
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


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

        try:
            excel.Run(_VBA_WAITFORM_NOTIFY_MACRO)
            _LOG.info("UI_READY: WaitForm dismiss (%s)", _VBA_WAITFORM_NOTIFY_MACRO)
        except Exception as ex:
            _LOG.warning("UI_READY: WaitForm dismiss failed (ignored): %s", str(ex))

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
