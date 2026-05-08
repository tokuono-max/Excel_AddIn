# -*- coding: utf-8 -*-
"""ribbon_public_to_svc の値が svc_server._ACTION_MAP と一致すること。"""

from __future__ import annotations

from core.ribbon_public_to_svc import (
    RIBBON_ACTIONS_DEFER_WAITFORM_DISMISS_TO_UI,
    RIBBON_ACTIONS_READY_UI_CLOSES_WAITFORM,
    RIBBON_INVOKE_ACTION_KEYS,
    RIBBON_INVOKE_FINALLY_NOTIFY_WAITFORM,
    RIBBON_PUBLIC_TO_SVC_ACTION,
    SVC_ACTIONS_NOTIFY_WAITFORM_AFTER_HANDLER,
)
from svc.svc_server import SVC_SERVER_ACTION_KEYS


def test_ribbon_target_actions_exist_in_svc_server() -> None:
    for pub, svc in RIBBON_PUBLIC_TO_SVC_ACTION.items():
        assert svc in SVC_SERVER_ACTION_KEYS, f"{pub!r} -> {svc!r} missing in svc_server"


def test_ribbon_covers_all_mapped_svc_actions() -> None:
    mapped = set(RIBBON_PUBLIC_TO_SVC_ACTION.values())
    assert mapped <= set(SVC_SERVER_ACTION_KEYS)
    assert len(RIBBON_PUBLIC_TO_SVC_ACTION) == len(mapped), "duplicate svc action in ribbon map"


def test_waitform_notify_partition_matches_all_invoke_actions() -> None:
    a = RIBBON_INVOKE_FINALLY_NOTIFY_WAITFORM
    b = RIBBON_ACTIONS_READY_UI_CLOSES_WAITFORM
    c = RIBBON_ACTIONS_DEFER_WAITFORM_DISMISS_TO_UI
    assert not (a & b)
    assert not (a & c)
    assert not (b & c)
    assert a | b | c == RIBBON_INVOKE_ACTION_KEYS


def test_svc_notify_after_handler_excludes_data_agg_includes_hd_in() -> None:
    assert SVC_ACTIONS_NOTIFY_WAITFORM_AFTER_HANDLER <= SVC_SERVER_ACTION_KEYS
    assert "data_agg" not in SVC_ACTIONS_NOTIFY_WAITFORM_AFTER_HANDLER
    assert "hd_in" in SVC_ACTIONS_NOTIFY_WAITFORM_AFTER_HANDLER
