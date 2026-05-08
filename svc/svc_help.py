# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: svc/svc_help.py
Created: 2026-03-19
Version: 1.0.0
Purpose:
  操作マニュアル表示。本文テンプレは config/ui_help.json の MESSAGES.HELP_BODY_TEMPLATE
  （1 文字列、または改行で連結する文字列の配列。{version},{bootstrap_version} を format）。
  core.version_txt で解決した VERSION.txt の版を差し込む。Info.txt は参照しない。
  IPC で ui_server にヘルプ窓表示を依頼。モーダルで閉じるまでポーリングし、
  閉じたらステータスバー復元と Excel 前面化を行う。画面は ui_qt.ui_help + config/ui_help.json。
"""
from __future__ import annotations

import os
import sys
import time
from pathlib import Path
from typing import Any, Optional

_path_svc = os.path.dirname(os.path.abspath(__file__))
_root = os.path.dirname(_path_svc)
if _root not in sys.path:
    sys.path.insert(0, _root)

from core.core_log import get_logger  # noqa: E402
from ui_qt.ipc_file import get_ipc_root, get_request_dir, read_pickle, write_pickle  # noqa: E402

logger = get_logger(__name__)
__version__ = "1.0.0"

try:
    from core import core_cst as cst
except Exception:
    cst = None  # type: ignore

try:
    from core import core_xlc as core_xlc_mod
except Exception:
    core_xlc_mod = None  # type: ignore

def _status_bar_save(book: Any) -> str:
    """現在の Excel ステータスバー文言を退避する。"""
    try:
        return str(book.app.api.StatusBar or "")
    except Exception:
        return ""


def _status_bar_restore(book: Any, saved: str) -> None:
    """ステータスバーを退避した文言に戻す。"""
    try:
        book.app.api.StatusBar = saved
    except Exception:
        pass


def _cfg() -> dict[str, Any]:
    """操作マニュアル用の画面設定を config/ui_help.json から読み込む。"""
    if cst is None:
        return {}
    return cst.get_ui_config_from_file_required("help")


def _msg(cfg: dict[str, Any], key: str, **fmt: Any) -> str:
    """設定の MESSAGES からキーに対応する文言を取得し、任意でフォーマットする。"""
    m = (cfg.get("MESSAGES") or {}).get(key) or key
    try:
        return str(m).format(**fmt)
    except Exception:
        return str(m)


def _help_body_template_raw(messages: dict[str, Any] | None) -> str:
    """
    MESSAGES.HELP_BODY_TEMPLATE を 1 本のテキストに正規化する。
    ・str: 前後の空白のみ除去
    ・list: 各要素を str にし \\n で連結（JSON 上で複数行に分けて保守しやすくする）
    """
    m = messages or {}
    raw: Any = m.get("HELP_BODY_TEMPLATE")
    if raw is None:
        return ""
    if isinstance(raw, str):
        return raw.strip()
    if isinstance(raw, list):
        parts: list[str] = []
        for item in raw:
            if item is None:
                continue
            parts.append(str(item))
        return "\n".join(parts).strip()
    return str(raw).strip()


def _load_help_content(cfg: dict[str, Any]) -> str:
    """ui_help.json のテンプレに bin+config と bootstrap 版を差し込んだ本文を返す。"""
    try:
        from core.packaged_update import (
            read_installed_bin_version,
            read_installed_bootstrap_version,
            read_installed_config_version,
        )
        from core.version_txt import candidate_version_txt_paths

        template = _help_body_template_raw(cfg.get("MESSAGES") or {})
        if not template:
            logger.warning("[HELP] HELP_BODY_TEMPLATE missing or empty in ui_help.json")
            return _msg(cfg, "FILE_NOT_FOUND")

        def _candidate_roots() -> list[Path]:
            out: list[Path] = []
            for p in candidate_version_txt_paths():
                out.append(p.parent)
            uniq: list[Path] = []
            seen: set[str] = set()
            for r in out:
                key = str(r)
                if key in seen:
                    continue
                seen.add(key)
                uniq.append(r)
            return uniq

        def _help_version_text() -> str:
            for root in _candidate_roots():
                if not root.is_dir():
                    continue
                bin_v = read_installed_bin_version(root)
                if not bin_v:
                    continue
                cfg_v = read_installed_config_version(root)
                if cfg_v:
                    return f"{bin_v}.{cfg_v}"
                return str(bin_v)
            return ""

        def _help_bootstrap_version_text() -> str:
            for root in _candidate_roots():
                if not root.is_dir():
                    continue
                bsv = read_installed_bootstrap_version(root)
                if bsv:
                    return str(bsv)
            return ""

        ver = _help_version_text()
        bs_ver = _help_bootstrap_version_text()
        vdisp = ver if ver else _msg(cfg, "HELP_VERSION_MISSING_BIN_CONFIG")
        bdisp = bs_ver if bs_ver else _msg(cfg, "HELP_VERSION_MISSING_BOOTSTRAP")
        return str(template.format(version=vdisp, bootstrap_version=bdisp)).strip() + "\n"
    except Exception as e:
        logger.warning("[HELP] embedded body failed: %s", e)
        return _msg(cfg, "FILE_NOT_FOUND")


def _get_window_rect(hwnd: int) -> tuple[int, int, int, int] | None:
    """Excel HWND の GetWindowRect。UI の excel_rect 用。"""
    if not int(hwnd or 0) or os.name != "nt":
        return None
    try:
        import ctypes
        from ctypes import wintypes

        r = wintypes.RECT()
        if ctypes.windll.user32.GetWindowRect(int(hwnd), ctypes.byref(r)):
            return (int(r.left), int(r.top), int(r.right), int(r.bottom))
    except Exception:
        pass
    return None


def _submit_help_ui(parent_hwnd: int, sheet_id: str, result_path: Path, content: str) -> None:
    """ヘルプ窓の表示を ui_server に依頼する。content を req_dict に含め、result_path に結果が書かれるまで呼び出し元でポーリングする。"""
    try:
        from svc.svc_host import ensure_ui_server

        ensure_ui_server()
        req_dict: dict[str, Any] = {
            "action": "help_show",
            "content": str(content),
            "modeless": False,
        }
        er_h = _get_window_rect(int(parent_hwnd or 0))
        if er_h is not None:
            req_dict["excel_rect"] = list(er_h)
        payload: dict[str, Any] = {
            "parent_hwnd": int(parent_hwnd),
            "result_path": str(result_path),
            "ready_path": "",
            "sheet_id": str(sheet_id),
            "log_path": "",
            "action": "help",
            "module": "ui_qt.ui_help",
            "req_dict": req_dict,
        }
        ts_ms = int(time.time() * 1000)
        req_path = get_request_dir() / f"req_help_{ts_ms}_{os.getpid()}.pkl"
        write_pickle(req_path, payload)
        logger.info("[HELP] help UI request: %s", req_path)
    except Exception as exc:
        logger.warning("[HELP] help UI request failed: %s", exc)


def _poll_result(result_path: Path, timeout_sec: float = 120.0) -> dict[str, Any] | None:
    """ヘルプ窓が閉じると result_path に結果が書かれるまでポーリングする。"""
    t0 = time.time()
    while (time.time() - t0) < timeout_sec:
        if result_path.exists() and result_path.stat().st_size > 0:
            try:
                return read_pickle(result_path)
            except Exception:
                pass
        time.sleep(0.05)
    return None


def show_help(target_hwnd: Optional[int] = None, sheet_id: str = "") -> None:
    """
    [VBA 呼出] 操作マニュアル（埋め込み＋VERSION.txt）をモーダルでヘルプ窓に表示する。
    窓を閉じたあとステータスバーを復元し、Excel を前面に戻す。
    """
    if core_xlc_mod is None:
        logger.error("[HELP] core_xlc not available")
        return
    ctx = core_xlc_mod.get_excel_context_from_hwnd(int(target_hwnd or 0), sheet_id)
    if ctx is None:
        logger.error("[HELP] Excel context not available (xlwings + HWND)")
        return

    ptr_a, ptr_w, ptr_s, ph = ctx
    logger.info("[HELP] 開始")
    cfg = _cfg()
    saved_status = _status_bar_save(ptr_w)

    try:
        content = _load_help_content(cfg)
        res_dir = Path(get_ipc_root()) / "result"
        res_dir.mkdir(parents=True, exist_ok=True)
        ts_ms = int(time.time() * 1000)
        result_path = res_dir / f"res_help_{ts_ms}_{os.getpid()}.pkl"
        sid = str(sheet_id or "").strip() or f"help_{abs(id(ptr_s))}"
        _submit_help_ui(ph, sid, result_path, content)
        _poll_result(result_path)
        logger.info("[HELP] 運用ログ ヘルプ表示 閉じるまで待機完了")
    except Exception as ex:
        logger.exception("[HELP] %s", ex)
    finally:
        try:
            _status_bar_restore(ptr_w, saved_status)
        except Exception:
            pass
        try:
            from core import core_w32

            core_w32.bring_to_front(ph)
        except Exception:
            pass
