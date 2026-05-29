# -*- coding: utf-8 -*-
"""update_process_cleanup mutex gate for interactive defer from svc_server."""

from __future__ import annotations

import sys
from pathlib import Path
from unittest.mock import patch

_ROOT = Path(__file__).resolve().parents[1]
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))

from core import update_process_cleanup as upc


def test_mutex_blocks_pending_apply_strict_svc() -> None:
    snap = {"main": False, "main_legacy": False, "svc": True, "ui": False}
    assert upc.mutex_blocks_pending_apply(snap, relax_svc_self=False)
    assert not upc.mutex_blocks_pending_apply(snap, relax_svc_self=True)


def test_mutex_blocks_pending_apply_main_ui() -> None:
    assert upc.mutex_blocks_pending_apply(
        {"main": True, "svc": False, "main_legacy": False, "ui": False},
        relax_svc_self=True,
    )
    assert upc.mutex_blocks_pending_apply(
        {"main": False, "svc": False, "main_legacy": False, "ui": True},
        relax_svc_self=True,
    )


def test_should_relax_svc_mutex_for_interactive_defer() -> None:
    pending = {"skip_apply_confirm": True}
    with patch.dict("os.environ", {}, clear=False):
        with patch.object(upc, "is_hc_svc_server_process", return_value=True):
            assert upc.should_relax_svc_mutex_for_interactive_defer(pending)
    with patch.object(upc, "is_hc_svc_server_process", return_value=False):
        assert not upc.should_relax_svc_mutex_for_interactive_defer(pending)
    assert not upc.should_relax_svc_mutex_for_interactive_defer({})
    with patch.dict("os.environ", {"CSV_TOOL_APPLY_PENDING_INLINE_BIN": "1"}):
        with patch.object(upc, "is_hc_svc_server_process", return_value=True):
            assert not upc.should_relax_svc_mutex_for_interactive_defer(pending)


if __name__ == "__main__":
    test_mutex_blocks_pending_apply_strict_svc()
    test_mutex_blocks_pending_apply_main_ui()
    test_should_relax_svc_mutex_for_interactive_defer()
    print("all passed")
