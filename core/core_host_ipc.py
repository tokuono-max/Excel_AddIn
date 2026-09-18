# -*- coding: utf-8 -*-
"""ホスト IPC のファイル操作（core 側）。svc プロセス生成はここへ移さない。"""
from __future__ import annotations

import tempfile
from pathlib import Path

_SVC_SHUTDOWN_FLAG = "svc_shutdown.flag"
_SVC_LAST_COM_HWND_FILE = "svc_last_com_hwnd.txt"


def host_control_dir() -> Path:
    """svc / UI 共通の control ディレクトリ。取得失敗時は一時フォルダへ退避。"""
    try:
        from ui_qt.ipc_file import get_ipc_root  # noqa: WPS433

        root = Path(get_ipc_root())
    except (ImportError, AttributeError, OSError, RuntimeError, TypeError, ValueError):
        root = Path(tempfile.gettempdir()) / "csv_tool"
    d = root / "control"
    d.mkdir(parents=True, exist_ok=True)
    return d


def write_svc_shutdown_flag(control_dir: Path) -> None:
    """svc 停止要求フラグを書く。書けなくても例外は外へ出さない。"""
    p = Path(control_dir) / _SVC_SHUTDOWN_FLAG
    try:
        p.write_text("shutdown", encoding="utf-8")
    except OSError:
        try:
            p.open("a").close()
        except OSError:
            pass


def read_last_svc_com_hwnd(control_dir: Path) -> int:
    """svc_server が最後に COM 接続した Excel HWND。無い・壊れているときは 0。"""
    try:
        p = Path(control_dir) / _SVC_LAST_COM_HWND_FILE
        if not p.exists():
            return 0
        return int((p.read_text(encoding="utf-8") or "0").strip() or "0")
    except (OSError, ValueError, UnicodeError, TypeError):
        return 0


def write_last_svc_com_hwnd(control_dir: Path, hwnd: int) -> None:
    """COM 接続先 HWND を記録する。0 以下は削除。"""
    ph = int(hwnd or 0)
    p = Path(control_dir) / _SVC_LAST_COM_HWND_FILE
    p.parent.mkdir(parents=True, exist_ok=True)
    try:
        if ph <= 0:
            p.unlink(missing_ok=True)
        else:
            p.write_text(str(ph), encoding="utf-8")
    except OSError:
        pass
