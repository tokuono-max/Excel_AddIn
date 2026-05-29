# -*- coding: utf-8 -*-
"""更新適用失敗通知・updater 結果ファイルのユニットテスト。"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from unittest.mock import patch

_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from bootstrap import update_bootstrap as ub
from core import packaged_update as pu


def test_notify_apply_result_if_failed_skips_deferred_to_updater(tmp_path: Path) -> None:
    root = tmp_path / "app"
    root.mkdir()
    (root / "config").mkdir()
    (root / "config" / "ui_update_check.json").write_text(
        json.dumps({"MESSAGES": {"UPDATER_ERROR_TEMPLATE": "ERR {error}"}}),
        encoding="utf-8",
    )
    with patch.object(ub, "_notify_apply_failure") as mock_notify:
        ub._notify_apply_result_if_failed(
            root,
            {"ok": True, "applied": False, "deferred_to_updater": True},
        )
        mock_notify.assert_not_called()


def test_notify_apply_result_if_failed_calls_on_error(tmp_path: Path) -> None:
    root = tmp_path / "app"
    root.mkdir()
    (root / "config").mkdir()
    (root / "config" / "ui_update_check.json").write_text(
        json.dumps({"MESSAGES": {"UPDATER_ERROR_TEMPLATE": "ERR {error}"}}),
        encoding="utf-8",
    )
    with patch.object(ub, "_notify_apply_failure") as mock_notify:
        ub._notify_apply_result_if_failed(
            root,
            {"ok": False, "applied": False, "error": "test failure"},
        )
        mock_notify.assert_called_once()
        assert "test failure" in mock_notify.call_args[0][1]


def test_maybe_show_updater_result_consumes_file(tmp_path: Path) -> None:
    root = tmp_path / "app"
    root.mkdir()
    (root / "config").mkdir()
    (root / "config" / "ui_update_check.json").write_text(
        json.dumps({"MESSAGES": {"UPDATER_ERROR_TEMPLATE": "ERR {error}"}}),
        encoding="utf-8",
    )
    p = pu.updater_result_path(root)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps({"ok": False, "error": "updater boom"}), encoding="utf-8")
    with patch.object(pu, "_message_box") as mock_mb:
        pu.maybe_show_updater_result_from_previous_run(root, owner_hwnd=0, sheet_id="_")
        mock_mb.assert_called_once()
    assert not p.is_file()


def test_write_updater_result(tmp_path: Path) -> None:
    from hc_updater import _write_updater_result

    p = tmp_path / "updater_last_result.json"
    _write_updater_result(
        p,
        ok=True,
        target_bin_version="1.2.0",
        display_version="1.2.0.1",
    )
    raw = json.loads(p.read_text(encoding="utf-8"))
    assert raw["ok"] is True
    assert raw["target_bin_version"] == "1.2.0"
    assert raw["display_version"] == "1.2.0.1"


def test_maybe_show_updater_result_success_skips_notify(tmp_path: Path) -> None:
    root = tmp_path / "app"
    root.mkdir()
    (root / "config").mkdir()
    (root / "config" / "ui_update_check.json").write_text("{}", encoding="utf-8")
    p = pu.updater_result_path(root)
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(
        json.dumps({"ok": True, "target_bin_version": "2.0.0", "display_version": "2.0.0.0"}),
        encoding="utf-8",
    )
    with patch.object(pu, "_message_box") as mock_mb:
        pu.maybe_show_updater_result_from_previous_run(root, owner_hwnd=0, sheet_id="_")
        mock_mb.assert_not_called()
    assert not p.is_file()


def test_resolve_require_admin_interactive_skips_app_dialog(tmp_path: Path) -> None:
    root = tmp_path / "inst"
    root.mkdir()
    with patch.object(pu, "_message_box") as mock_mb:
        assert (
            pu._resolve_require_admin_for_bin_prompt(
                root, "all", interactive_apply_now=True
            )
            is True
        )
        assert (
            pu._resolve_require_admin_for_bin_prompt(
                root, "current", interactive_apply_now=True
            )
            is False
        )
        mock_mb.assert_not_called()


def test_updater_result_path_under_install_root(tmp_path: Path) -> None:
    root = tmp_path / "inst"
    root.mkdir()
    assert pu.updater_result_path(root).name == "updater_last_result.json"
    assert "update" in pu.updater_result_path(root).parts


def test_notify_installed_apps_list_changed_calls_shchangenotify_on_nt() -> None:
    with patch.object(pu.os, "name", "nt"):
        with patch("ctypes.windll") as mock_windll:
            pu.notify_installed_apps_list_changed()
            mock_windll.shell32.SHChangeNotify.assert_called_once_with(
                pu.SHCNE_ASSOCCHANGED,
                pu.SHCNF_IDLIST,
                None,
                None,
            )


def test_apply_succeeded_for_interactive() -> None:
    assert pu._apply_succeeded_for_interactive({"ok": True, "applied": True})
    assert pu._apply_succeeded_for_interactive(
        {"ok": True, "applied": False, "deferred_to_updater": True}
    )
    assert not pu._apply_succeeded_for_interactive(
        {"ok": True, "applied": False, "skipped": "concurrent_apply"}
    )
    assert not pu._apply_succeeded_for_interactive(
        {"ok": True, "applied": False, "deferred": True}
    )
    assert not pu._apply_succeeded_for_interactive({"ok": False, "applied": False})


def test_apply_pending_with_retry_on_concurrent(tmp_path: Path) -> None:
    root = tmp_path / "app"
    root.mkdir()
    calls: list[int] = []

    def fake_apply(_root: Path) -> dict:
        calls.append(1)
        if len(calls) < 3:
            return {"ok": True, "applied": False, "skipped": "concurrent_apply"}
        return {"ok": True, "applied": False, "deferred_to_updater": True}

    with patch.object(pu.time, "sleep"):
        with patch(
            "bootstrap.update_bootstrap.apply_pending_update",
            side_effect=fake_apply,
        ):
            res = pu._apply_pending_update_with_retry(root, source="test")
    assert res.get("deferred_to_updater") is True
    assert len(calls) == 3


def test_run_interactive_bin_apply_now_failure_clears_pending(tmp_path: Path) -> None:
    root = tmp_path / "app"
    root.mkdir()
    (root / "update").mkdir(parents=True)
    pending_path = root / "update" / "pending.json"
    pending_path.write_text('{"state":"downloaded"}', encoding="utf-8")
    (root / "config").mkdir()
    (root / "config" / "ui_update_check.json").write_text(
        json.dumps(
            {
                "MESSAGES": {
                    "UPDATER_ERROR_TEMPLATE": "ERR {error} log={log_path}",
                }
            }
        ),
        encoding="utf-8",
    )
    with patch.object(
        pu,
        "_apply_pending_update_with_retry",
        return_value={"ok": True, "applied": False, "deferred": True},
    ):
        with patch.object(pu, "_message_box") as mock_mb:
            ok = pu.run_interactive_bin_apply_now(
                root, owner_hwnd=0, sheet_id="_", source="test"
            )
    assert ok is False
    mock_mb.assert_called_once()
    assert not pending_path.is_file()


def test_run_interactive_bin_apply_now_success_deferred(tmp_path: Path) -> None:
    root = tmp_path / "app"
    root.mkdir()
    with patch.object(
        pu,
        "_apply_pending_update_with_retry",
        return_value={"ok": True, "applied": False, "deferred_to_updater": True},
    ):
        with patch.object(pu, "_message_box") as mock_mb:
            ok = pu.run_interactive_bin_apply_now(
                root, owner_hwnd=0, sheet_id="_", source="ribbon"
            )
    assert ok is True
    mock_mb.assert_not_called()


def test_notify_installed_apps_list_changed_skips_non_nt() -> None:
    with patch.object(pu.os, "name", "posix"):
        with patch("ctypes.windll") as mock_windll:
            pu.notify_installed_apps_list_changed()
            mock_windll.shell32.SHChangeNotify.assert_not_called()


if __name__ == "__main__":
    import sys

    tests = [
        test_apply_succeeded_for_interactive,
        test_apply_pending_with_retry_on_concurrent,
        test_run_interactive_bin_apply_now_failure_clears_pending,
        test_run_interactive_bin_apply_now_success_deferred,
        test_notify_apply_result_if_failed_skips_deferred_to_updater,
        test_notify_apply_result_if_failed_calls_on_error,
        test_maybe_show_updater_result_consumes_file,
        test_maybe_show_updater_result_success_skips_notify,
        test_write_updater_result,
        test_updater_result_path_under_install_root,
        test_notify_installed_apps_list_changed_calls_shchangenotify_on_nt,
        test_notify_installed_apps_list_changed_skips_non_nt,
    ]
    for t in tests:
        import tempfile

        with tempfile.TemporaryDirectory() as td:
            t(Path(td))
        print("ok", t.__name__)
    print("all passed")
    sys.exit(0)
