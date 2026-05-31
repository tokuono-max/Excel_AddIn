# -*- coding: utf-8 -*-
"""通知ダイアログ表示時のお知らせ音（Windows MessageBeep）。"""
from __future__ import annotations

import ctypes
import os
from typing import Any, Literal

NotificationKind = Literal["done", "info", "error"]

_MB_OK = 0x0000
_MB_ICONHAND = 0x0010
_MB_ICONEXCLAMATION = 0x0030
_MB_ICONASTERISK = 0x0040

_KIND_TO_BEEP: dict[str, int] = {
    "done": _MB_OK,
    "info": _MB_ICONASTERISK,
    "error": _MB_ICONHAND,
}


def notification_sound_enabled(kind: str) -> bool:
    """HC_NOTIFICATION_SOUND=0 で全体オフ。HC_NOTIFICATION_SOUND_DONE 等で種別ごとに制御。"""
    from core.core_env import truthy

    master = os.environ.get("HC_NOTIFICATION_SOUND")
    if master is not None and str(master).strip() != "":
        if not truthy(master, empty_means_false=False):
            return False
    key = f"HC_NOTIFICATION_SOUND_{str(kind or '').strip().upper()}"
    raw = os.environ.get(key)
    if raw is not None and str(raw).strip() != "":
        return truthy(raw, empty_means_false=False)
    return True


def notification_kind_from_icon(icon_key: str) -> NotificationKind:
    k = str(icon_key or "").strip().lower()
    if k in ("critical", "error", "stop"):
        return "error"
    return "info"


def play_notification_sound(kind: str = "done") -> None:
    """終了=done / お知らせ=info / エラー=error。"""
    k = str(kind or "done").strip().lower()
    if k not in _KIND_TO_BEEP:
        k = "done"
    if not notification_sound_enabled(k):
        return
    if os.name != "nt":
        return
    try:
        ctypes.windll.user32.MessageBeep(_KIND_TO_BEEP[k])  # type: ignore[attr-defined]
    except Exception:
        pass


def play_notification_on_widget(widget: Any) -> None:
    """ダイアログの _hc_notification_sound_kind（未設定時 done）に従って鳴らす。"""
    kind = str(getattr(widget, "_hc_notification_sound_kind", "") or "done").strip().lower()
    play_notification_sound(kind)


def play_notification_for_icon(
    icon_key: str,
    *,
    default_kind: NotificationKind = "info",
) -> None:
    k = str(icon_key or "").strip().lower()
    if not k:
        play_notification_sound(default_kind)
        return
    if k in ("information", "info"):
        play_notification_sound(default_kind)
        return
    play_notification_sound(notification_kind_from_icon(k))
