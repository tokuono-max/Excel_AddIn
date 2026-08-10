# -*- coding: utf-8 -*-
"""svc_server._attach_book 補助（マルチ Excel COM 解放・HWND 解決）の単体テスト。"""
from __future__ import annotations

from types import SimpleNamespace

import svc.svc_server as svc_server


class _FakeImpl:
    def __init__(self, *, hwnd: int = 0, xl_hwnd: int = 0) -> None:
        self._hwnd = int(hwnd or 0)
        self._xl = SimpleNamespace(Hwnd=int(xl_hwnd)) if xl_hwnd else None


class _FakeBooks:
    def __init__(self, *, active: object | None = None, fail: bool = False) -> None:
        self._active = active
        self._fail = fail

    @property
    def active(self):
        if self._fail:
            raise RuntimeError("com broken")
        return self._active


class _FakeApp:
    def __init__(self, impl: _FakeImpl, *, books_fail: bool = False) -> None:
        self.impl = impl
        self.books = _FakeBooks(active=SimpleNamespace(name="Book1"), fail=books_fail)


def test_app_impl_hwnd_prefers_lazy_hwnd_without_xl() -> None:
    impl = _FakeImpl(hwnd=855586)
    assert svc_server._app_impl_hwnd(impl) == 855586


def test_app_impl_hwnd_falls_back_to_xl_hwnd() -> None:
    impl = _FakeImpl(hwnd=0, xl_hwnd=1707224)
    assert svc_server._app_impl_hwnd(impl) == 1707224


def test_is_com_broken_detects_com_error_type() -> None:
    class com_error(Exception):
        pass

    assert svc_server._is_com_broken(com_error("rpc")) is True
    assert svc_server._is_com_broken(RuntimeError("other")) is False


def test_validate_book_alive_and_label() -> None:
    class _Book:
        name = "Book1"

    class _DeadBook:
        @property
        def name(self) -> str:
            raise RuntimeError("com dead")

    assert svc_server._validate_book_alive(_Book()) is True
    assert svc_server._validate_book_alive(_DeadBook()) is False
    assert svc_server._book_label(_Book()) == "Book1"
    assert svc_server._book_label(_DeadBook()) == "?"


def test_release_other_excel_com_bindings_clears_foreign_xl(monkeypatch) -> None:
    keep_impl = _FakeImpl(hwnd=100, xl_hwnd=100)
    other_impl = _FakeImpl(hwnd=200, xl_hwnd=200)
    keep_app = SimpleNamespace(impl=keep_impl)
    other_app = SimpleNamespace(impl=other_impl)

    class _FakeApps:
        def __iter__(self):
            yield keep_app
            yield other_app

    import xlwings as xw

    monkeypatch.setattr(xw, "apps", _FakeApps(), raising=False)
    monkeypatch.setattr(svc_server, "_excel_hwnd_is_live", lambda _h: True)

    with svc_server._book_cache_lock:
        svc_server._book_cache_by_hwnd[100] = object()
        svc_server._book_cache_by_hwnd[200] = object()

    svc_server._release_other_excel_com_bindings(100)

    assert keep_impl._xl is not None
    assert other_impl._xl is None
    assert other_impl._hwnd == 200
    with svc_server._book_cache_lock:
        assert 100 in svc_server._book_cache_by_hwnd
        assert 200 not in svc_server._book_cache_by_hwnd


def test_release_other_excel_com_bindings_purges_dead_hwnd(monkeypatch) -> None:
    dead_impl = _FakeImpl(hwnd=200, xl_hwnd=200)
    dead_app = SimpleNamespace(impl=dead_impl)

    class _FakeApps:
        def __iter__(self):
            yield dead_app

    import xlwings as xw

    monkeypatch.setattr(xw, "apps", _FakeApps(), raising=False)
    monkeypatch.setattr(svc_server, "_excel_hwnd_is_live", lambda h: int(h) == 100)

    with svc_server._book_cache_lock:
        svc_server._book_cache_by_hwnd[200] = object()

    svc_server._release_other_excel_com_bindings(100)

    assert dead_impl._xl is None
    assert dead_impl._hwnd == 0
    with svc_server._book_cache_lock:
        assert 200 not in svc_server._book_cache_by_hwnd


def test_purge_dead_excel_app_shells(monkeypatch) -> None:
    dead_impl = _FakeImpl(hwnd=6164446, xl_hwnd=6164446)
    live_impl = _FakeImpl(hwnd=1314602, xl_hwnd=1314602)
    dead_app = SimpleNamespace(impl=dead_impl)
    live_app = SimpleNamespace(impl=live_impl)

    class _FakeApps:
        def __iter__(self):
            yield dead_app
            yield live_app

    import xlwings as xw

    monkeypatch.setattr(xw, "apps", _FakeApps(), raising=False)
    monkeypatch.setattr(
        svc_server,
        "_excel_hwnd_is_live",
        lambda h: int(h) == 1314602,
    )

    with svc_server._book_cache_lock:
        svc_server._book_cache_by_hwnd[6164446] = object()
        svc_server._book_cache_by_hwnd[1314602] = object()

    svc_server._purge_dead_excel_app_shells()

    assert dead_impl._xl is None
    assert dead_impl._hwnd == 0
    assert live_impl._xl is not None
    with svc_server._book_cache_lock:
        assert 6164446 not in svc_server._book_cache_by_hwnd
        assert 1314602 in svc_server._book_cache_by_hwnd


def test_reset_app_binding_for_hwnd(monkeypatch) -> None:
    impl = _FakeImpl(hwnd=1314602, xl_hwnd=1314602)
    app = SimpleNamespace(impl=impl)

    class _FakeApps:
        def __iter__(self):
            yield app

    import xlwings as xw

    monkeypatch.setattr(xw, "apps", _FakeApps(), raising=False)

    with svc_server._book_cache_lock:
        svc_server._book_cache_by_hwnd[1314602] = object()

    svc_server._reset_app_binding_for_hwnd(1314602)

    assert impl._xl is None
    assert impl._hwnd == 1314602
    with svc_server._book_cache_lock:
        assert 1314602 not in svc_server._book_cache_by_hwnd


def test_cached_book_if_alive_drops_stale_entry(monkeypatch) -> None:
    class _DeadBook:
        @property
        def name(self) -> str:
            raise RuntimeError("com dead")

    monkeypatch.setattr(svc_server, "_excel_hwnd_is_live", lambda _h: True)

    with svc_server._book_cache_lock:
        svc_server._book_cache_by_hwnd[999] = _DeadBook()

    assert svc_server._cached_book_if_alive(999) is None
    with svc_server._book_cache_lock:
        assert 999 not in svc_server._book_cache_by_hwnd


def test_cached_book_if_alive_drops_dead_hwnd(monkeypatch) -> None:
    monkeypatch.setattr(svc_server, "_excel_hwnd_is_live", lambda _h: False)

    with svc_server._book_cache_lock:
        svc_server._book_cache_by_hwnd[999] = SimpleNamespace(name="Book1")

    assert svc_server._cached_book_if_alive(999) is None
    with svc_server._book_cache_lock:
        assert 999 not in svc_server._book_cache_by_hwnd


def test_find_or_create_app_for_hwnd_uses_apps_scan(monkeypatch) -> None:
    hit_impl = _FakeImpl(hwnd=555)
    hit_app = _FakeApp(hit_impl)

    class _FakeApps:
        def __iter__(self):
            yield hit_app

    import xlwings as xw

    monkeypatch.setattr(xw, "apps", _FakeApps(), raising=False)
    monkeypatch.setattr(svc_server, "_excel_hwnd_is_live", lambda _h: True)

    created: list[object] = []

    def _fake_app(*, impl=None):
        created.append(impl)
        return SimpleNamespace(impl=impl)

    monkeypatch.setattr(xw, "App", _fake_app, raising=False)

    app = svc_server._find_or_create_app_for_hwnd(555)
    assert app is hit_app
    assert created == []


def test_find_or_create_app_for_hwnd_recreates_stale_shell(monkeypatch) -> None:
    stale_impl = _FakeImpl(hwnd=555)
    stale_app = _FakeApp(stale_impl, books_fail=True)

    class _FakeApps:
        def __iter__(self):
            yield stale_app

    import xlwings as xw
    from xlwings._xlwindows import App as WinApp

    monkeypatch.setattr(xw, "apps", _FakeApps(), raising=False)
    monkeypatch.setattr(svc_server, "_excel_hwnd_is_live", lambda _h: True)

    created: list[object] = []

    def _fake_app(*, impl=None):
        created.append(impl)
        return SimpleNamespace(impl=impl)

    monkeypatch.setattr(xw, "App", _fake_app, raising=False)

    app = svc_server._find_or_create_app_for_hwnd(555)
    assert app.impl is created[0]
    assert isinstance(created[0], WinApp)
    assert stale_impl._xl is None


def test_hard_reset_all_excel_com_bindings_clears_everything(monkeypatch) -> None:
    impl_a = _FakeImpl(hwnd=100, xl_hwnd=100)
    impl_b = _FakeImpl(hwnd=200, xl_hwnd=200)
    app_a = SimpleNamespace(impl=impl_a)
    app_b = SimpleNamespace(impl=impl_b)

    class _FakeApps:
        def __iter__(self):
            yield app_a
            yield app_b

    import xlwings as xw

    monkeypatch.setattr(xw, "apps", _FakeApps(), raising=False)

    with svc_server._book_cache_lock:
        svc_server._book_cache_by_hwnd[100] = object()
        svc_server._book_cache_by_hwnd[200] = object()

    svc_server._hard_reset_all_excel_com_bindings()

    assert impl_a._xl is None
    assert impl_a._hwnd == 0
    assert impl_b._xl is None
    assert impl_b._hwnd == 0
    with svc_server._book_cache_lock:
        assert svc_server._book_cache_by_hwnd == {}


def test_attach_book_via_xlc_context_returns_book(monkeypatch) -> None:
    class _Book:
        name = "Book1"

    import core.core_xlc as core_xlc

    monkeypatch.setattr(
        core_xlc,
        "get_excel_context_from_hwnd",
        lambda _h, _s: (object(), _Book(), object(), 856336),
        raising=False,
    )

    book = svc_server._attach_book_via_xlc_context(
        856336, book_fullname="Book1", book_name="Book1"
    )
    assert book is not None
    assert book.name == "Book1"


def test_attach_book_prefers_xlc_context(monkeypatch) -> None:
    class _Book:
        name = "Book1"

    import core.core_xlc as core_xlc

    monkeypatch.setattr(
        core_xlc,
        "get_excel_context_from_hwnd",
        lambda _h, _s: (object(), _Book(), object(), 856336),
        raising=False,
    )
    monkeypatch.setattr(svc_server, "_purge_dead_excel_app_shells", lambda: None)
    monkeypatch.setattr(svc_server, "_cached_book_if_alive", lambda _h: None)
    monkeypatch.setattr(svc_server, "_excel_hwnd_is_live", lambda _h: True)
    monkeypatch.setattr(svc_server, "_release_other_excel_com_bindings", lambda _h: None)

    def _fail_xlwings(*_a, **_k):
        raise AssertionError("xlwings path should not run when xlc succeeds")

    monkeypatch.setattr(svc_server, "_find_or_create_app_for_hwnd", _fail_xlwings)
    monkeypatch.setattr(
        "svc.svc_host.write_last_svc_com_hwnd",
        lambda _h: None,
    )
    book = svc_server._attach_book(856336, "Book1", "Book1")
    assert book.name == "Book1"
    assert svc_server._last_attached_hwnd == 856336


def test_attach_book_skip_cache_bypasses_cached_book(monkeypatch) -> None:
    class _StaleBook:
        name = "Stale"

    stale = _StaleBook()
    with svc_server._book_cache_lock:
        svc_server._book_cache_by_hwnd[856336] = stale

    class _FreshBook:
        name = "Fresh"

    import core.core_xlc as core_xlc

    monkeypatch.setattr(
        core_xlc,
        "get_excel_context_from_hwnd",
        lambda _h, _s: (object(), _FreshBook(), object(), 856336),
        raising=False,
    )
    monkeypatch.setattr(svc_server, "_purge_dead_excel_app_shells", lambda: None)
    monkeypatch.setattr(svc_server, "_excel_hwnd_is_live", lambda _h: True)
    monkeypatch.setattr(svc_server, "_release_other_excel_com_bindings", lambda _h: None)
    monkeypatch.setattr(
        "svc.svc_host.write_last_svc_com_hwnd",
        lambda _h: None,
    )

    book = svc_server._attach_book(856336, "Book1", "Book1", skip_cache=True)
    assert book.name == "Fresh"
    with svc_server._book_cache_lock:
        assert svc_server._book_cache_by_hwnd[856336] is book


def test_invalidate_attached_book_cache(monkeypatch) -> None:
    monkeypatch.setattr(svc_server, "_reset_app_binding_for_hwnd", lambda _h: None)
    with svc_server._book_cache_lock:
        svc_server._book_cache_by_hwnd[111] = object()
    svc_server.invalidate_attached_book_cache(111)
    with svc_server._book_cache_lock:
        assert 111 not in svc_server._book_cache_by_hwnd


def test_resolve_fresh_book_after_ui_wait_uses_skip_cache(monkeypatch) -> None:
    class _StaleBook:
        name = "Stale"

        class app:
            hwnd = 856336

    class _FreshBook:
        name = "Fresh"

    calls: list[dict] = []

    def _fake_attach(**kwargs):
        calls.append(kwargs)
        return _FreshBook()

    monkeypatch.setattr(svc_server, "_attach_book", _fake_attach)

    out = svc_server.resolve_fresh_book_after_ui_wait(
        _StaleBook(),
        excel_hwnd=856336,
        book_fullname="Book1",
        book_name="Book1",
        log_prefix="TEST",
    )
    assert out.name == "Fresh"
    assert len(calls) == 1
    assert calls[0]["skip_cache"] is True
    assert calls[0]["excel_hwnd"] == 856336


def test_resolve_fresh_book_after_ui_wait_returns_original_on_failure(monkeypatch) -> None:
    class _Book:
        name = "Only"

    def _fail_attach(**_kwargs):
        raise RuntimeError("com dead")

    monkeypatch.setattr(svc_server, "_attach_book", _fail_attach)
    book = _Book()
    assert (
        svc_server.resolve_fresh_book_after_ui_wait(
            book, excel_hwnd=1, log_prefix="TEST"
        )
        is book
    )


def test_find_or_create_app_for_hwnd_force_fresh_skips_scan(monkeypatch) -> None:
    hit_impl = _FakeImpl(hwnd=555)
    hit_app = _FakeApp(hit_impl)

    class _FakeApps:
        def __iter__(self):
            yield hit_app

    import xlwings as xw
    from xlwings._xlwindows import App as WinApp

    monkeypatch.setattr(xw, "apps", _FakeApps(), raising=False)
    monkeypatch.setattr(svc_server, "_excel_hwnd_is_live", lambda _h: True)

    created: list[object] = []

    def _fake_app(*, impl=None):
        created.append(impl)
        return SimpleNamespace(impl=impl)

    monkeypatch.setattr(xw, "App", _fake_app, raising=False)

    app = svc_server._find_or_create_app_for_hwnd(555, force_fresh=True)
    assert app is not hit_app
    assert isinstance(created[0], WinApp)
