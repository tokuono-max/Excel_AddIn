# -*- coding: utf-8 -*-
"""
Python: 3.12
Module: core/core_w32.py
Created: 2026-01-30
Updated: 2026-06-29
Version: 2.2.0
Purpose:
  Win32 API の薄いラッパ（UI非依存 / Tk 禁止）。
  - HWND の owner 設定（タスクバー抑止 / 裏回り防止）
  - EnableWindow による入力抑止（Excel 実質モーダルの基盤）
  - 位置/矩形取得、前面化、TOPMOST 制御（ベストエフォート）

Design:
  - 本モジュールは Tk/Qt を一切 import しない（core の UI 依存禁止）。
  - 失敗しても例外を投げない（Excelロック解除漏れの方が致命）。

History (latest 3):
  - 2.2.0 (2026-06-29) win32_get_open_file_name / win32_get_save_file_name: comdlg32 直叩き（hwndOwner=Excel）。qt_name_filter_to_win32。
  - 2.1.9 (2026-06-29) enum_visible_top_level_windows_for_pid: ネイティブファイルダイアログ前面化の HWND 探索用。
  - 2.1.8 (2026-05-03) get_process_image_path_for_diag: PID から QueryFullProcessImageNameW で exe パス短縮（EXCEL_FRONT_FOLLOW 診断の前景 PID 特定用）。
  - 2.1.7 (2026-04-13) WS_EX_TOOLWINDOW は既定で付与しない（タイトルバー min/max 表示のため）。付与は HC_USE_WS_EX_TOOLWINDOW_FOR_TASKBAR=1 またはウィジェット属性。get_window_exstyle_toolwindow 追加。set_owner から拡張スタイル適用を除去（_set_owner_hwnd の遅延経路のみ widget 付きで判定）。
  - 2.1.6 (2026-04-13) get_window_caption_style_summary（GWL_STYLE・最小化/最大化ボックス bit、HC_UI_WINDOW_CAPTION_DIAG 用）。
  - 2.1.5 (2026-04-12) set_foreground_window_attach_input / nudge_top_level_to_foreground（モーダル終了後の CMD 前面化対策）。
  - 2.1.4 (2026-04-10) 診断用: get_window_class_name / get_owner_hwnd / is_window_visible / set_foreground_window_result / format_ui_fg_diag_line（HC_UI_FG_DIAG）。
  - 2.1.3 (2026-04-08) set_owner 後に apply_taskbar_hiding_extended_style: WS_EX_TOOLWINDOW + WS_EX_APPWINDOW 解除でタスクバー非表示を強化。HC_SKIP_WS_EX_TOOLWINDOW=1 で無効化。
  - 2.1.2 (2026-03-01) 追加: set_window_style_remove_min_max（最小化・最大化ボタンをタイトルバーから削除）。
  - 2.1.1 (2026-03-01) set_owner から WS_EX_TOOLWINDOW を削除。Windows11 タイトルバー×赤表示を解消。
"""

from __future__ import annotations

import ctypes
import os
from ctypes import wintypes
from typing import Any, Iterable, Optional

_user32 = ctypes.windll.user32
_kernel32 = ctypes.windll.kernel32


# ==============================================================================
# Basic types
# ==============================================================================
class RECT(wintypes.RECT):
    """GetWindowRect 等で使用する RECT。"""


# ==============================================================================
# Small helpers (never raise)
# ==============================================================================
def enable_window(hwnd: int, enable: bool) -> None:
    """指定 HWND の入力可否を切り替える（失敗しても例外にしない）。

    Args:
        hwnd: 対象 HWND
        enable: True=有効化, False=無効化
    """
    try:
        _user32.EnableWindow(int(hwnd), bool(enable))
    except Exception:
        pass


# 変数: バージョン情報
__version__ = "2.2.0"

_comdlg32 = ctypes.windll.comdlg32

_OFN_FILEMUSTEXIST = 0x00001000
_OFN_PATHMUSTEXIST = 0x00000800
_OFN_HIDEREADONLY = 0x00000004
_OFN_OVERWRITEPROMPT = 0x00000002
_OFN_EXPLORER = 0x00080000
_OFN_ENABLESIZING = 0x00800000
_OFN_NOCHANGEDIR = 0x00000008
_WIN32_FILE_BUFFER_CHARS = 65536


class _OPENFILENAMEW(ctypes.Structure):
    _fields_ = [
        ("lStructSize", wintypes.DWORD),
        ("hwndOwner", wintypes.HWND),
        ("hInstance", wintypes.HINSTANCE),
        ("lpstrFilter", wintypes.LPCWSTR),
        ("lpstrCustomFilter", wintypes.LPWSTR),
        ("nMaxCustFilter", wintypes.DWORD),
        ("nFilterIndex", wintypes.DWORD),
        ("lpstrFile", wintypes.LPWSTR),
        ("nMaxFile", wintypes.DWORD),
        ("lpstrFileTitle", wintypes.LPWSTR),
        ("nMaxFileTitle", wintypes.DWORD),
        ("lpstrInitialDir", wintypes.LPCWSTR),
        ("lpstrTitle", wintypes.LPCWSTR),
        ("Flags", wintypes.DWORD),
        ("nFileOffset", wintypes.WORD),
        ("nFileExtension", wintypes.WORD),
        ("lpstrDefExt", wintypes.LPCWSTR),
        ("lCustData", wintypes.LPARAM),
        ("lpfnHook", wintypes.LPVOID),
        ("lpTemplateName", wintypes.LPCWSTR),
        ("pvReserved", wintypes.LPVOID),
        ("dwReserved", wintypes.DWORD),
        ("FlagsEx", wintypes.DWORD),
    ]


def qt_name_filter_to_win32(filter_str: str) -> str:
    """Qt の名前フィルタ（;; 区切り）を Win32 lpstrFilter 形式へ変換する。"""
    pairs: list[str] = []
    for seg in (filter_str or "").split(";;"):
        seg = seg.strip()
        if not seg:
            continue
        if "(" in seg and seg.endswith(")"):
            i = seg.rfind("(")
            desc = seg[:i].strip() or seg
            pat = seg[i + 1 : -1].strip() or "*.*"
        else:
            desc = seg
            pat = "*.*"
        pairs.extend([desc, pat])
    if not pairs:
        pairs = ["すべてのファイル", "*.*"]
    return "\0".join(pairs) + "\0\0"


def _first_def_ext_from_win32_filter(win32_filter: str) -> str:
    """Win32 フィルタの先頭パターンから拡張子（ドットなし）を推定する。"""
    try:
        parts = (win32_filter or "").split("\0")
        if len(parts) < 2:
            return ""
        pat = (parts[1] or "").strip()
        if not pat or pat == "*.*":
            return ""
        for token in pat.replace(",", " ").split():
            t = token.strip()
            if t.startswith("*.") and len(t) > 2:
                return t[2:].lstrip(".")
    except Exception:
        pass
    return ""


def win32_get_open_file_name(
    owner_hwnd: int,
    title: str,
    initial_dir: str,
    filter_str: str,
) -> str:
    """Win32 GetOpenFileNameW。hwndOwner に Excel を渡し Qt ホストを出さない。"""
    try:
        owner = int(owner_hwnd or 0)
        if not owner:
            return ""
        win32_filter = qt_name_filter_to_win32(filter_str)
        file_buf = ctypes.create_unicode_buffer(_WIN32_FILE_BUFFER_CHARS)
        ofn = _OPENFILENAMEW()
        ofn.lStructSize = ctypes.sizeof(_OPENFILENAMEW)
        ofn.hwndOwner = owner
        ofn.lpstrFilter = win32_filter
        ofn.lpstrFile = ctypes.cast(file_buf, wintypes.LPWSTR)
        ofn.nMaxFile = _WIN32_FILE_BUFFER_CHARS
        init_dir = (initial_dir or "").strip()
        if init_dir:
            ofn.lpstrInitialDir = init_dir
        dlg_title = (title or "").strip()
        if dlg_title:
            ofn.lpstrTitle = dlg_title
        ofn.Flags = (
            _OFN_FILEMUSTEXIST
            | _OFN_PATHMUSTEXIST
            | _OFN_HIDEREADONLY
            | _OFN_EXPLORER
            | _OFN_ENABLESIZING
            | _OFN_NOCHANGEDIR
        )
        if not _comdlg32.GetOpenFileNameW(ctypes.byref(ofn)):
            return ""
        return (file_buf.value or "").strip()
    except Exception:
        return ""


def win32_get_save_file_name(
    owner_hwnd: int,
    title: str,
    initial_path: str,
    filter_str: str,
) -> str:
    """Win32 GetSaveFileNameW。hwndOwner に Excel を渡す。"""
    try:
        owner = int(owner_hwnd or 0)
        if not owner:
            return ""
        win32_filter = qt_name_filter_to_win32(filter_str)
        ip = (initial_path or "").strip()
        init_dir = ""
        init_file = ""
        if ip:
            init_dir = os.path.dirname(ip) or ""
            init_file = os.path.basename(ip) or ""
        file_buf = ctypes.create_unicode_buffer(_WIN32_FILE_BUFFER_CHARS)
        if init_file:
            file_buf.value = init_file
        ofn = _OPENFILENAMEW()
        ofn.lStructSize = ctypes.sizeof(_OPENFILENAMEW)
        ofn.hwndOwner = owner
        ofn.lpstrFilter = win32_filter
        ofn.lpstrFile = ctypes.cast(file_buf, wintypes.LPWSTR)
        ofn.nMaxFile = _WIN32_FILE_BUFFER_CHARS
        if init_dir:
            ofn.lpstrInitialDir = init_dir
        dlg_title = (title or "").strip()
        if dlg_title:
            ofn.lpstrTitle = dlg_title
        def_ext = _first_def_ext_from_win32_filter(win32_filter)
        if def_ext:
            ofn.lpstrDefExt = def_ext
        ofn.Flags = (
            _OFN_OVERWRITEPROMPT
            | _OFN_PATHMUSTEXIST
            | _OFN_EXPLORER
            | _OFN_ENABLESIZING
            | _OFN_NOCHANGEDIR
        )
        if not _comdlg32.GetSaveFileNameW(ctypes.byref(ofn)):
            return ""
        return (file_buf.value or "").strip()
    except Exception:
        return ""


def enum_visible_top_level_windows_for_pid(pid: int) -> list[int]:
    """指定 PID の可視トップレベル HWND を列挙する（EnumWindows。失敗時 []）。"""
    target = int(pid or 0)
    if not target:
        return []
    out: list[int] = []

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def _cb(hwnd: int, _lparam: int) -> bool:
        try:
            h = int(hwnd or 0)
            if not h or not is_window_visible(h):
                return True
            if int(get_window_pid(h) or 0) != target:
                return True
            out.append(h)
        except Exception:
            pass
        return True

    try:
        _user32.EnumWindows(_cb, 0)
    except Exception:
        return []
    return out


def get_foreground_window() -> int:
    """現在のフォアグラウンド HWND を返す（失敗時 0）。"""
    try:
        return int(_user32.GetForegroundWindow())
    except Exception:
        return 0


def get_root_window(hwnd: int) -> int:
    """指定 HWND のルート（トップレベル）ウィンドウを返す。オーナー設定はルートにするとタスクバー非表示が安定する。

    Args:
        hwnd: 任意のウィンドウ HWND（子ウィンドウ可）

    Returns:
        ルート HWND。失敗時は引数をそのまま返す。
    """
    try:
        GA_ROOT = 2
        root = _user32.GetAncestor(int(hwnd), GA_ROOT)
        return int(root) if root else int(hwnd)
    except Exception:
        return int(hwnd)


def get_window_rect(hwnd: int) -> Optional[tuple[int, int, int, int]]:
    """ウィンドウ矩形（物理ピクセル）を取得する。

    Args:
        hwnd: 対象 HWND

    Returns:
        (left, top, right, bottom) / 取得失敗時 None
    """
    try:
        rect = RECT()
        ok = _user32.GetWindowRect(int(hwnd), ctypes.byref(rect))
        if not ok:
            return None
        return (int(rect.left), int(rect.top), int(rect.right), int(rect.bottom))
    except Exception:
        return None


def move_window(
    hwnd: int, x: int, y: int, w: int, h: int, repaint: bool = True
) -> None:
    """MoveWindow を実行する（失敗しても例外にしない）。"""
    try:
        _user32.MoveWindow(int(hwnd), int(x), int(y), int(w), int(h), bool(repaint))
    except Exception:
        pass


def center_to_owner(hwnd_child: int, hwnd_owner: int) -> None:
    """子ウィンドウを owner の中央へ配置する（物理ピクセル）。

    Args:
        hwnd_child: 対象ダイアログ HWND
        hwnd_owner: owner（Excel HWND）
    """
    try:
        r_owner = get_window_rect(hwnd_owner)
        r_child = get_window_rect(hwnd_child)
        if r_owner is None or r_child is None:
            return

        ol, ot, or_, ob = r_owner
        cl, ct, cr, cb = r_child
        cw = max(1, cr - cl)
        ch = max(1, cb - ct)

        ocx = ol + (or_ - ol) // 2
        ocy = ot + (ob - ot) // 2

        x = int(ocx - cw // 2)
        y = int(ocy - ch // 2)
        move_window(hwnd_child, x, y, cw, ch, repaint=True)
    except Exception:
        pass


def _env_skip_toolwindow_exstyle() -> bool:
    v = (os.environ.get("HC_SKIP_WS_EX_TOOLWINDOW") or "").strip().lower()
    return v in ("1", "true", "yes", "on")


def _env_use_ws_ex_toolwindow_for_taskbar() -> bool:
    """従来のタスクバー抑止（WS_EX_TOOLWINDOW）を明示的に有効にするか。"""
    v = (os.environ.get("HC_USE_WS_EX_TOOLWINDOW_FOR_TASKBAR") or "").strip().lower()
    return v in ("1", "true", "yes", "on")


def apply_taskbar_hiding_extended_style(hwnd: int, widget: Any = None) -> None:
    """タスクバーにボタンが出にくくする拡張スタイルを付与する（ベストエフォート）。

    Excel とは別プロセスの Qt ウィンドウは GWLP_HWNDPARENT のみではタスクバーに残ることがある。
    WS_EX_TOOLWINDOW に加え WS_EX_APPWINDOW を外すと、タイトルバーに最小化・最大化ボタンが
    描画されない環境がある（GWL_STYLE のビットは残る）。そのため既定では本スタイルを付与しない。

    付与条件（いずれか）:
      - 環境変数 HC_USE_WS_EX_TOOLWINDOW_FOR_TASKBAR=1
      - Qt ウィジェットにプロパティ _hc_use_ws_ex_toolwindow_for_taskbar が True（JSON WINDOW の
        USE_WS_EX_TOOLWINDOW_FOR_TASKBAR で設定）

    無効化（最優先）: HC_SKIP_WS_EX_TOOLWINDOW=1
    """
    if _env_skip_toolwindow_exstyle():
        return
    use_tool = _env_use_ws_ex_toolwindow_for_taskbar()
    if widget is not None:
        try:
            p = widget.property("_hc_use_ws_ex_toolwindow_for_taskbar")
            if p is not None:
                use_tool = bool(p)
        except Exception:
            pass
    if not use_tool:
        return
    try:
        GWL_EXSTYLE = -20
        WS_EX_TOOLWINDOW = 0x00000080
        WS_EX_APPWINDOW = 0x00040000
        SWP_NOMOVE = 0x0002
        SWP_NOSIZE = 0x0001
        SWP_NOZORDER = 0x0004
        SWP_FRAMECHANGED = 0x0020
        get_long = getattr(_user32, "GetWindowLongPtrW", None) or _user32.GetWindowLongW
        set_long = getattr(_user32, "SetWindowLongPtrW", None) or _user32.SetWindowLongW
        h = int(hwnd)
        if not h:
            return
        ex = int(get_long(h, GWL_EXSTYLE))
        ex = (ex | WS_EX_TOOLWINDOW) & ~WS_EX_APPWINDOW
        set_long(h, GWL_EXSTYLE, ex)
        _user32.SetWindowPos(
            h, 0, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED
        )
    except Exception:
        pass


def set_owner(hwnd_child: int, hwnd_owner: int) -> None:
    """HWND の owner（GWLP_HWNDPARENT）を設定する。

    Args:
        hwnd_child: 子（ダイアログ）HWND
        hwnd_owner: owner（Excel）HWND

    Note:
        WS_EX_TOOLWINDOW は ui_qt.ui_common._set_owner_hwnd の遅延経路で、オプトイン時のみ
        apply_taskbar_hiding_extended_style に委ねる（既定では owner のみでタイトルバーを保全）。
    """
    try:
        GWLP_HWNDPARENT = -8

        set_long = getattr(_user32, "SetWindowLongPtrW", None)
        if set_long is None:
            set_long = _user32.SetWindowLongW

        # 1) owner 設定（Z オーダー連動・親子関係）
        set_long(int(hwnd_child), GWLP_HWNDPARENT, int(hwnd_owner))

        # 2) フレーム更新
        SWP_NOMOVE = 0x0002
        SWP_NOSIZE = 0x0001
        SWP_NOZORDER = 0x0004
        SWP_FRAMECHANGED = 0x0020
        _user32.SetWindowPos(
            int(hwnd_child),
            0,
            0,
            0,
            0,
            0,
            SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED,
        )
    except Exception:
        pass


def bring_to_front(hwnd: int) -> None:
    """前面化（ベストエフォート）。"""
    try:
        _user32.SetForegroundWindow(int(hwnd))
    except Exception:
        pass


def set_foreground_window_attach_input(target_hwnd: int) -> bool:
    """AttachThreadInput で現在スレッドとフォアグラウンドスレッドを結び、SetForegroundWindow を再試行する。

    モーダル終了直後など、単独の SetForegroundWindow が無視される環境向け。失敗時 False。
    """
    try:
        target = int(target_hwnd or 0)
        if not target:
            return False
        cur_tid = int(_kernel32.GetCurrentThreadId())
        fg = int(_user32.GetForegroundWindow() or 0)
        fg_tid = 0
        if fg:
            pid = wintypes.DWORD(0)
            fg_tid = int(_user32.GetWindowThreadProcessId(fg, ctypes.byref(pid)) or 0)
        attached = False
        if fg_tid and fg_tid != cur_tid:
            try:
                attached = bool(_user32.AttachThreadInput(cur_tid, fg_tid, True))
            except Exception:
                attached = False
        try:
            return bool(_user32.SetForegroundWindow(target))
        finally:
            if attached:
                try:
                    _user32.AttachThreadInput(cur_tid, fg_tid, False)
                except Exception:
                    pass
    except Exception:
        return False


def nudge_top_level_to_foreground(hwnd: int) -> None:
    """トップレベル HWND を Z 順・フォアグラウンドへまとめてベストエフォートで寄せる（例外なし）。"""
    try:
        h = int(hwnd or 0)
        if not h:
            return
        try:
            if bool(_user32.IsIconic(h)):
                SW_RESTORE = 9
                _user32.ShowWindow(h, SW_RESTORE)
        except Exception:
            pass
        HWND_TOP = 0
        SWP_NOMOVE = 0x0002
        SWP_NOSIZE = 0x0001
        SWP_SHOWWINDOW = 0x0040
        try:
            _user32.SetWindowPos(
                h,
                HWND_TOP,
                0,
                0,
                0,
                0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW,
            )
        except Exception:
            pass
        try:
            _user32.BringWindowToTop(h)
        except Exception:
            pass
        ok = False
        try:
            ok = bool(_user32.SetForegroundWindow(h))
        except Exception:
            pass
        if not ok:
            set_foreground_window_attach_input(h)
    except Exception:
        pass


def set_topmost(hwnd: int, enable: bool) -> None:
    """TOPMOST を設定/解除する（ベストエフォート）。"""
    try:
        HWND_TOPMOST = -1
        HWND_NOTOPMOST = -2
        SWP_NOMOVE = 0x0002
        SWP_NOSIZE = 0x0001
        flags = SWP_NOMOVE | SWP_NOSIZE
        _user32.SetWindowPos(
            int(hwnd), HWND_TOPMOST if enable else HWND_NOTOPMOST, 0, 0, 0, 0, flags
        )
    except Exception:
        pass


def get_window_exstyle_toolwindow(hwnd: int) -> tuple[int, bool]:
    """GWL_EXSTYLE と WS_EX_TOOLWINDOW の有無。失敗時 (0, False)。"""
    try:
        GWL_EXSTYLE = -20
        WS_EX_TOOLWINDOW = 0x00000080
        get_long = getattr(_user32, "GetWindowLongPtrW", None) or _user32.GetWindowLongW
        ex = int(get_long(int(hwnd), GWL_EXSTYLE))
        return (ex, bool(ex & WS_EX_TOOLWINDOW))
    except Exception:
        return (0, False)


def get_window_caption_style_summary(hwnd: int) -> tuple[int, bool, bool]:
    """GWL_STYLE と WS_MINIMIZEBOX / WS_MAXIMIZEBOX の有無。失敗時 (0, False, False)。"""
    try:
        GWL_STYLE = -16
        WS_MINIMIZEBOX = 0x00020000
        WS_MAXIMIZEBOX = 0x00010000
        get_long = getattr(_user32, "GetWindowLongPtrW", None) or _user32.GetWindowLongW
        style = int(get_long(int(hwnd), GWL_STYLE))
        return (style, bool(style & WS_MINIMIZEBOX), bool(style & WS_MAXIMIZEBOX))
    except Exception:
        return (0, False, False)


def set_window_style_remove_min_max(hwnd: int) -> None:
    """タイトルバーから最小化・最大化ボタンを削除する（Win32 スタイルで強制）。"""
    try:
        GWL_STYLE = -16
        WS_MINIMIZEBOX = 0x00020000
        WS_MAXIMIZEBOX = 0x00010000
        SWP_NOMOVE = 0x0002
        SWP_NOSIZE = 0x0001
        SWP_FRAMECHANGED = 0x0020
        get_long = getattr(_user32, "GetWindowLongPtrW", None) or _user32.GetWindowLongW
        set_long = getattr(_user32, "SetWindowLongPtrW", None) or _user32.SetWindowLongW
        style = int(get_long(int(hwnd), GWL_STYLE))
        style &= ~(WS_MINIMIZEBOX | WS_MAXIMIZEBOX)
        set_long(int(hwnd), GWL_STYLE, style)
        _user32.SetWindowPos(int(hwnd), 0, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_FRAMECHANGED)
    except Exception:
        pass


def enum_child_windows(hwnd_parent: int) -> list[int]:
    """EnumChildWindows で子HWNDを列挙する（失敗時は空）。"""
    try:
        result: list[int] = []

        WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)

        def _cb(hwnd: int, lparam: int) -> bool:  # noqa: ANN001
            result.append(int(hwnd))
            return True

        _user32.EnumChildWindows(int(hwnd_parent), WNDENUMPROC(_cb), 0)
        return result
    except Exception:
        return []


# ==============================================================================
# Process / foreground hook
# ==============================================================================
def get_window_pid(hwnd: int) -> int:
    """Get process id of the window (best-effort).

    Args:
        hwnd: Target window handle.

    Returns:
        PID on success, otherwise 0.
    """
    try:
        pid = wintypes.DWORD(0)
        _user32.GetWindowThreadProcessId(int(hwnd), ctypes.byref(pid))
        return int(pid.value)
    except Exception:
        return 0


def is_process_alive(pid: int) -> bool:
    """PID のプロセスが生存していれば True。

    Windows では GetExitCodeProcess を使う（os.kill(0) は ACCESS_DENIED で誤判定しやすい）。
    判定不能時は生存扱いにして Python の誤終了を防ぐ。
    """
    pid_i = int(pid or 0)
    if pid_i <= 0:
        return False
    if os.name != "nt":
        try:
            os.kill(pid_i, 0)
            return True
        except OSError:
            return False
    STILL_ACTIVE = 259
    PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
    ERROR_ACCESS_DENIED = 5
    try:
        h = _kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid_i)
        if not h:
            if int(_kernel32.GetLastError() or 0) == ERROR_ACCESS_DENIED:
                return True
            return False
        code = wintypes.DWORD()
        ok = _kernel32.GetExitCodeProcess(h, ctypes.byref(code))
        _kernel32.CloseHandle(h)
        if not ok:
            return True
        return int(code.value) == STILL_ACTIVE
    except Exception:
        return True


def is_window(hwnd: int) -> bool:
    """HWND が有効なウィンドウなら True（Excel 終了後は False）。"""
    try:
        return bool(_user32.IsWindow(int(hwnd or 0)))
    except Exception:
        return False


def get_process_image_path_for_diag(pid: int, max_chars: int = 96) -> str:
    """PID に紐づく実行ファイルパスを短く返す（ログ用。失敗時は空文字）。"""
    try:
        pid_i = int(pid or 0)
        if not pid_i:
            return ""
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        hproc = _kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, 0, pid_i)
        if not hproc:
            return ""
        try:
            buf = ctypes.create_unicode_buffer(2048)
            dw = wintypes.DWORD(ctypes.sizeof(buf) // ctypes.sizeof(ctypes.c_wchar))
            ok = _kernel32.QueryFullProcessImageNameW(hproc, 0, buf, ctypes.byref(dw))
            if not ok:
                return ""
            raw = str(buf.value or "").strip()
            if not raw:
                return ""
            return _sanitize_diag_token(raw, int(max_chars) if max_chars else 96)
        finally:
            _kernel32.CloseHandle(hproc)
    except Exception:
        return ""


def get_window_class_name(hwnd: int) -> str:
    """GetWindowClassNameW（失敗時は空文字）。"""
    try:
        h = int(hwnd or 0)
        if not h:
            return ""
        buf = ctypes.create_unicode_buffer(256)
        n = int(_user32.GetClassNameW(h, buf, 256) or 0)
        if n > 0:
            return str(buf.value or "")
    except Exception:
        pass
    return ""


def get_owner_hwnd(hwnd: int) -> int:
    """GetWindow(GW_OWNER)。オーナーなし時は 0。"""
    try:
        GW_OWNER = 4
        r = _user32.GetWindow(int(hwnd or 0), GW_OWNER)
        return int(r) if r else 0
    except Exception:
        return 0


def is_window_visible(hwnd: int) -> bool:
    try:
        return bool(_user32.IsWindowVisible(int(hwnd or 0)))
    except Exception:
        return False


def set_foreground_window_result(hwnd: int) -> bool:
    """SetForegroundWindow の成否（例外時 False）。"""
    try:
        return bool(_user32.SetForegroundWindow(int(hwnd or 0)))
    except Exception:
        return False


def _sanitize_diag_token(s: str, max_len: int = 56) -> str:
    t = "".join(c if 32 <= ord(c) < 0x110000 and c not in "\r\n\t" else "?" for c in str(s or ""))
    if len(t) > max_len:
        return t[: max_len - 3] + "..."
    return t


def format_hwnd_diag(hwnd: int) -> str:
    """1 HWND を pid / cls / vis / owner 付きで短く整形（ログ用）。"""
    h = int(hwnd or 0)
    if not h:
        return "h=0"
    cls = _sanitize_diag_token(get_window_class_name(h), 48)
    return (
        f"h={h} pid={get_window_pid(h)} cls={cls!r} "
        f"vis={int(is_window_visible(h))} owner={get_owner_hwnd(h)}"
    )


def format_ui_fg_diag_line(
    phase: str,
    parent_hwnd: int,
    dlg_hwnd: int,
    *,
    sfw_ok: Optional[bool] = None,
    extra: str = "",
) -> str:
    """
    フォアグラウンド調査用 1 行。parent=Excel 想定 HWND、dlg=Qt ダイアログ HWND。
    sfw_ok: 直前の SetForegroundWindow( dlg ) の成否が分かるときだけ渡す。
    """
    fg = int(get_foreground_window() or 0)
    ph = int(parent_hwnd or 0)
    dh = int(dlg_hwnd or 0)
    parts = [
        f"[UI_FG] phase={_sanitize_diag_token(phase, 40)!r}",
        f"fg={fg}({format_hwnd_diag(fg)})",
        f"parent={ph}({format_hwnd_diag(ph)})",
        f"dlg={dh}({format_hwnd_diag(dh)})",
        f"fg==parent={int(fg == ph)} fg==dlg={int(fg == dh)}",
    ]
    if sfw_ok is not None:
        parts.append(f"sfw_ok={int(bool(sfw_ok))}")
    if extra:
        parts.append(_sanitize_diag_token(extra, 120))
    return " ".join(parts)


_FG_HOOK: Optional[int] = None
_FG_PROC = None
_FG_CB = None


def start_foreground_hook(callback) -> None:  # noqa: ANN001
    """Start foreground change hook (best-effort).

    This installs WinEvent hook for EVENT_SYSTEM_FOREGROUND and calls `callback(hwnd)`
    when the foreground window changes.

    Notes:
        - Never raises.
        - Only one hook is kept (new call replaces previous).

    Args:
        callback: Callable that receives foreground hwnd (int).
    """
    try:
        stop_foreground_hook()

        EVENT_SYSTEM_FOREGROUND = 0x0003
        WINEVENT_OUTOFCONTEXT = 0x0000
        WINEVENT_SKIPOWNPROCESS = 0x0002

        WinEventProc = ctypes.WINFUNCTYPE(
            None,
            wintypes.HANDLE,
            wintypes.DWORD,
            wintypes.HWND,
            wintypes.LONG,
            wintypes.LONG,
            wintypes.DWORD,
            wintypes.DWORD,
        )

        def _proc(hook, event, hwnd, obj_id, child_id, event_thread, event_time):  # noqa: ANN001
            try:
                if callable(_FG_CB):
                    _FG_CB(int(hwnd))
            except Exception:
                return

        global _FG_PROC, _FG_CB, _FG_HOOK
        _FG_CB = callback
        _FG_PROC = WinEventProc(_proc)

        hook = _user32.SetWinEventHook(
            EVENT_SYSTEM_FOREGROUND,
            EVENT_SYSTEM_FOREGROUND,
            0,
            _FG_PROC,
            0,
            0,
            WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS,
        )
        _FG_HOOK = int(hook or 0)
    except Exception:
        return


def stop_foreground_hook() -> None:
    """Stop foreground change hook (best-effort)."""
    try:
        global _FG_HOOK, _FG_PROC, _FG_CB
        if _FG_HOOK:
            try:
                _user32.UnhookWinEvent(wintypes.HANDLE(_FG_HOOK))
            except Exception:
                pass
        _FG_HOOK = None
        _FG_PROC = None
        _FG_CB = None
    except Exception:
        return


def nudge_to_front(hwnd: int) -> None:
    """Bring window to front by temporary TOPMOST toggle (best-effort).

    This is useful when SetForegroundWindow alone is ignored.
    """
    try:
        HWND_TOPMOST = -1
        HWND_NOTOPMOST = -2
        SWP_NOMOVE = 0x0002
        SWP_NOSIZE = 0x0001
        SWP_SHOWWINDOW = 0x0040
        flags = SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW

        _user32.SetWindowPos(int(hwnd), HWND_TOPMOST, 0, 0, 0, 0, flags)
        _user32.SetWindowPos(int(hwnd), HWND_NOTOPMOST, 0, 0, 0, 0, flags)
        _user32.SetForegroundWindow(int(hwnd))
    except Exception:
        return


def enable_windows(hwnds: Iterable[int], enable: bool) -> None:
    """複数HWNDをまとめて EnableWindow する。"""
    for h in hwnds:
        enable_window(int(h), enable)
