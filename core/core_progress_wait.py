# -*- coding: utf-8 -*-
"""進捗 DONE 書込後、UI が pickle を受理する猶予を与える（ScreenUpdating 復帰前）。"""
from __future__ import annotations

import time

try:
    from core import core_xlc as xlc
except Exception:  # pragma: no cover
    xlc = None  # type: ignore


def wait_after_progress_done(*, min_sec: float = 1.0) -> None:
    """DONE pickle 書込後、進捗 UI がクローズ処理に入る猶予（ScreenUpdating 復帰前）。"""
    deadline = time.monotonic() + max(0.2, float(min_sec))
    while time.monotonic() < deadline:
        if xlc is not None:
            try:
                xlc.yield_to_excel()
            except Exception:
                pass
        time.sleep(0.05)
