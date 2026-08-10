# -*- coding: utf-8 -*-
"""core.excel_com_session（B+ COM セッション方針）の単体テスト。"""
from __future__ import annotations

import svc.svc_server as svc_server
from core import excel_com_session as ecs


def test_svc_server_actions_com_coverage() -> None:
    """svc_server の全 action が COM 方針に登録されている（update_check のみ除外）。"""
    all_actions = set(svc_server.SVC_SERVER_ACTION_KEYS)
    touching = set(ecs.SVC_COM_TOUCHING_ACTIONS)
    assert all_actions - touching == {"update_check"}
    assert touching <= all_actions


def test_attach_book_actions_subset_of_com_touching() -> None:
    assert ecs.SVC_ATTACH_BOOK_ACTIONS <= ecs.SVC_COM_TOUCHING_ACTIONS


def test_action_uses_attach_book() -> None:
    assert ecs.action_uses_attach_book("csv_ld") is True
    assert ecs.action_uses_attach_book("data_agg") is False
    assert ecs.action_uses_attach_book("update_check") is False


def test_action_touches_excel_com() -> None:
    assert ecs.action_touches_excel_com("dupli") is True
    assert ecs.action_touches_excel_com("data_agg") is True
    assert ecs.action_touches_excel_com("update_check") is False


def test_is_com_session_error_detects_stale_and_rpc() -> None:
    assert ecs.is_com_session_error(RuntimeError("Workbook COM stale (hwnd=1)")) is True
    assert ecs.is_com_session_error(
        RuntimeError("リモート プロシージャ コール (RPC) で内部エラー")
    ) is True
    assert ecs.is_com_session_error(RuntimeError("file not found")) is False
    assert ecs.is_com_session_error(None) is False


def test_is_com_session_error_detects_system_error_with_com_cause() -> None:
    class com_error(Exception):
        pass

    root = com_error(-2147220995, "オブジェクトをサーバーに接続できません", None, None)
    wrapped = SystemError(
        "<class 'logging.LogRecord'> returned a result with an exception set"
    )
    wrapped.__cause__ = root
    assert ecs.is_com_session_error(wrapped) is True


def test_action_attach_book_fresh_resolve() -> None:
    assert ecs.action_attach_book_fresh_resolve("csv_ld") is True
    assert ecs.action_attach_book_fresh_resolve("data_agg") is False


def test_should_not_schedule_recycle_on_success_for_com_action() -> None:
    assert (
        ecs.should_schedule_com_recycle_after_handler("csv_ld", handler_ok=True) is False
    )
    assert (
        ecs.should_schedule_com_recycle_after_handler(
            "update_check", handler_ok=True
        )
        is False
    )


def test_should_schedule_recycle_on_com_error() -> None:
    ex = RuntimeError("Workbook COM stale (hwnd=99)")
    assert (
        ecs.should_schedule_com_recycle_after_handler(
            "data_agg", handler_ok=False, exc=ex
        )
        is True
    )
    assert (
        ecs.should_schedule_com_recycle_after_handler(
            "data_agg",
            handler_ok=False,
            exc=RuntimeError("config missing"),
        )
        is False
    )


def test_should_schedule_recycle_on_system_error_wrapping_com() -> None:
    class com_error(Exception):
        pass

    root = com_error(-2147220995, "server", None, None)
    wrapped = SystemError("LogRecord returned a result with an exception set")
    wrapped.__cause__ = root
    assert (
        ecs.should_schedule_com_recycle_after_handler(
            "csv_ld", handler_ok=False, exc=wrapped
        )
        is True
    )


def test_prepare_com_session_is_noop() -> None:
    assert ecs.prepare_com_session_before_request(12345) is False
    assert ecs.prepare_com_session_before_request(0) is False


def test_record_and_read_com_session_hwnd(tmp_path, monkeypatch) -> None:
    import svc.svc_host as svc_host

    monkeypatch.setattr(svc_host, "_control_dir", lambda: tmp_path)
    ecs.record_com_session_hwnd(777)
    assert ecs.read_last_com_session_hwnd() == 777
