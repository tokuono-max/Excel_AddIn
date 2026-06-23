# -*- coding: utf-8 -*-
"""
svc_server の Excel COM セッション方針（B+）。

不変条件:
  1. 常駐 svc_server は HWND ごとに Book キャッシュ（_book_cache_by_hwnd）でマルチ Excel を扱う。
  2. COM を触る handler の成功後はプロセスを維持する（warmup / handler キャッシュを再利用）。
  3. COM セッション汚染と判定した失敗時のみ svc_server を終了（com_recycle）する。
  4. リクエスト前の事前再起動は行わない（ensure が死んでいれば spawn のみ）。
  5. Book 取得は get_excel_context_from_hwnd（xlc 経路）を優先する。

update_check は Excel COM に触れないため recycle 対象外。
新しい svc action を追加するときは本モジュールの action 集合を更新すること。
"""
from __future__ import annotations

# svc_server._ACTION_MAP のうち _attach_book 経路で Book を渡す action
SVC_ATTACH_BOOK_ACTIONS: frozenset[str] = frozenset(
    {
        "csv_mg",
        "csv_ld",
        "csv_sv",
        "csv_sp",
        "hd_in",
        "hd_nr",
        "undo",
    }
)

# handler 内で Excel COM / xlwings を使用する action（update_check を除く）
SVC_COM_TOUCHING_ACTIONS: frozenset[str] = frozenset(
    {
        *SVC_ATTACH_BOOK_ACTIONS,
        "dupli",
        "row_dl",
        "col_dl",
        "dt_ymd",
        "dt_hm",
        "trm_ex",
        "help",
        "data_agg",
    }
)

_COM_STALE_MARKERS: tuple[str, ...] = (
    "COM stale",
    "Excel window not found",
    "オブジェクトをサーバーに接続できません",
    "リモート プロシージャ コール",
    "RPC",
    "そのインターフェイスは認識されません",
    "Workbook not found",
)


def action_uses_attach_book(action: str) -> bool:
    """svc_server が _attach_book で Book を解決して handler に渡す action か。"""
    return str(action or "").strip() in SVC_ATTACH_BOOK_ACTIONS


def action_touches_excel_com(action: str) -> bool:
    """handler 実行により svc_server プロセス内で Excel COM が使われる action か。"""
    return str(action or "").strip() in SVC_COM_TOUCHING_ACTIONS


def is_com_session_error(exc: BaseException) -> bool:
    """COM セッション汚染とみなし svc_server recycle を促す例外か。"""
    try:
        import pywintypes

        if isinstance(exc, pywintypes.com_error):
            return True
    except Exception:
        pass
    if type(exc).__name__ in ("com_error", "COMError"):
        return True
    msg = str(exc)
    return any(marker in msg for marker in _COM_STALE_MARKERS)


def should_schedule_com_recycle_after_handler(
    action: str,
    *,
    handler_ok: bool,
    exc: BaseException | None = None,
) -> bool:
    """handler 後に svc_server の COM recycle（自終了）を予約すべきか。

    成功時は常駐を維持する。COM 汚染とみなす失敗時のみ recycle する。
    """
    if not action_touches_excel_com(action):
        return False
    if handler_ok:
        return False
    return exc is not None and is_com_session_error(exc)


def prepare_com_session_before_request(target_hwnd: int) -> bool:
    """リクエスト前の COM セッション準備（B+: 常駐維持のため事前再起動は行わない）。"""
    _ = int(target_hwnd or 0)
    return False


def record_com_session_hwnd(hwnd: int) -> None:
    """COM 接続成功時: 診断用に HWND を IPC へ記録する。"""
    from svc.svc_host import write_last_svc_com_hwnd

    write_last_svc_com_hwnd(int(hwnd or 0))


def read_last_com_session_hwnd() -> int:
    """前回 svc_server が COM 接続した Excel HWND（診断用）。"""
    from svc.svc_host import read_last_svc_com_hwnd

    return read_last_svc_com_hwnd()
