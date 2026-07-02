# -*- coding: utf-8 -*-
"""ファイル確定直後に同一 ui_server プロセスで進捗を即表示する共通ヘルパ。"""
from __future__ import annotations

import logging
import time
from pathlib import Path
from typing import Any

from core.csv_tool_progress_ui import enrich_progress_req_dict, initial_run_progress_pickle
from ui_qt import ipc_file

logger = logging.getLogger(__name__)


def try_show_immediate_progress_after_pick(
    *,
    feature: str,
    mod: Any,
    parent_hwnd: int,
    sheet_id: str,
    phase_total: int,
    phase_label: str,
    detail: str = "",
    done_delay_ms: int | None = None,
    progress_closed_path: str | None = None,
    get_window_rect,
) -> bool:
    """保存先/読込ファイル確定直後に進捗を即表示。成功時 True。"""
    sid = str(sheet_id or "_").strip() or "_"
    feat = str(feature or "").strip().lower()
    if not feat or mod is None:
        return False
    try:
        root = Path(str(ipc_file.get_ipc_root()))
        progress_path = root / "progress" / f"progress_{feat}_{sid}.pkl"
        progress_path.parent.mkdir(parents=True, exist_ok=True)
        try:
            from ui_qt.ui_dialog_progress import dismiss_progress_dialogs_for_path

            dismiss_progress_dialogs_for_path(str(progress_path))
        except Exception:
            pass
        ipc_file.write_pickle(
            progress_path,
            initial_run_progress_pickle(
                phase_total=int(phase_total),
                phase_label=phase_label,
                detail=detail,
            ),
        )
        progress_req_dict = enrich_progress_req_dict(
            {
                "action": "progress",
                "progress_path": str(progress_path),
                "phase_total": int(phase_total),
                "excel_lock": True,
                "bring_excel_first": False,
                "refront_on_run": feat == "ld",
                # comdlg32 直後: show 前 opacity 0 → 前面化 → reveal（Excel COM 砂時計は reveal 後）
                "opacity_reveal_before_show": True,
            },
            done_delay_ms=done_delay_ms,
            no_native_window=True,
        )
        if progress_closed_path:
            progress_req_dict["progress_closed_path"] = str(progress_closed_path)
        excel_rect = get_window_rect(int(parent_hwnd or 0))
        if excel_rect is not None:
            progress_req_dict["excel_rect"] = excel_rect
        progress_dlg = mod.create_dialog(progress_req_dict, int(parent_hwnd or 0), sid)
        if hasattr(progress_dlg, "show"):
            progress_dlg.show()
            if feat == "ld":
                try:
                    from ui_qt.ui_dialog_progress import ensure_progress_dialog_front

                    ensure_progress_dialog_front(progress_dlg)
                except Exception:
                    pass
            try:
                from PySide6.QtWidgets import QApplication

                app = QApplication.instance()
                if app is not None:
                    app.processEvents()
            except Exception:
                pass
            logger.debug(
                "[CSV_TOOL] immediate progress shown feature=%s sheet_id=%s t=%.3f",
                feat,
                sid,
                time.time(),
            )
            return True
    except Exception as exc:
        logger.warning(
            "[CSV_TOOL] immediate progress show failed feature=%s sheet_id=%s err=%s",
            feat,
            sid,
            exc,
        )
    return False
