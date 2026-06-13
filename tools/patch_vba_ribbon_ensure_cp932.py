# -*- coding: utf-8 -*-
"""マルチ Excel 初回リボン対策パッチ（cp932 厳守）。"""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAIN = ROOT / "VBA" / "Main.bas"

ENSURE_SUB = """\
' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: EnsurePythonHostsReady
' 改版番号および履歴: 1.1.0 (2026-06-13) hwnd を渡して register_book。マルチ Excel 初回リボン対策。
' プロシージャの動作概要: ensure_python_hosts_ready(hwnd) を 1 回 RunPython で呼ぶ。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub EnsurePythonHostsReady()
    On Error Resume Next
    RunPython "from svc.svc_host import ensure_python_hosts_ready; ensure_python_hosts_ready(" & CStr(Application.hwnd) & ")"
    On Error GoTo 0
End Sub


"""

VERSION_219 = (
    "'   2.19.0 (2026-06-13) [起動] EnsurePythonHostsReady に hwnd 渡し・WaitForm を RunPython 前へ移動。\n"
)

VERSION_218 = (
    "'   2.18.0 (2026-06-13) [起動] リボン操作前に EnsurePythonHostsReady（Python 生存確認）。\n"
)

OLD_ENSURE_RUN = (
    '    RunPython "from svc.svc_host import ensure_python_hosts_ready; ensure_python_hosts_ready()"\n'
)
NEW_ENSURE_RUN = (
    '    RunPython "from svc.svc_host import ensure_python_hosts_ready; ensure_python_hosts_ready(" & CStr(Application.hwnd) & ")"\n'
)

RIBBON_OLD = (
    "    Call Main.EnsurePythonHostsReady\n"
    "    Call HC_WaitForm.BeginWaitForRibbon(control.ID, act)\n"
)
RIBBON_NEW = (
    "    Call HC_WaitForm.BeginWaitForRibbon(control.ID, act)\n"
    "    Call Main.EnsurePythonHostsReady\n"
)


def _read_cp932(path: Path) -> str:
    return path.read_text(encoding="cp932")


def _write_cp932(path: Path, text: str) -> None:
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    path.write_bytes(normalized.encode("cp932").replace(b"\n", b"\r\n"))


def main() -> None:
    text = _read_cp932(MAIN)
    norm = text.replace("\r\n", "\n")

    if "2.19.0 (2026-06-13)" not in norm:
        if "'   2.18.0 (2026-06-13)" in norm:
            norm = norm.replace(
                "'   2.18.0 (2026-06-13)",
                VERSION_219 + "'   2.18.0 (2026-06-13)",
                1,
            )
        elif "'   2.17.0 (2026-06-13)" in norm:
            norm = norm.replace(
                "'   2.17.0 (2026-06-13)",
                VERSION_219 + VERSION_218 + "'   2.17.0 (2026-06-13)",
                1,
            )
        else:
            raise SystemExit("Main.bas: version anchor not found")
        print("Main.bas: version 2.19.0 added")

    if "EnsurePythonHostsReady" not in norm:
        anchor = "End Sub\n\n\n' Python 単一引用符リテラル用エスケープ"
        if anchor not in norm:
            raise SystemExit("Main.bas: EnsurePythonHostsReady insert anchor not found")
        norm = norm.replace(
            anchor,
            "End Sub\n\n\n" + ENSURE_SUB.replace("\r\n", "\n") + "' Python 単一引用符リテラル用エスケープ",
            1,
        )
        print("Main.bas: EnsurePythonHostsReady added")
    elif OLD_ENSURE_RUN.replace("\r\n", "\n") in norm:
        norm = norm.replace(OLD_ENSURE_RUN.replace("\r\n", "\n"), NEW_ENSURE_RUN.replace("\r\n", "\n"), 1)
        norm = norm.replace(
            "' 改版番号および履歴: 1.0.0 (2026-06-13) リボン操作前に svc/ui/bridge の生存確認（死んでいれば起動）。\n"
            "' プロシージャの動作概要: ensure_svc_ui_bridge_parallel を 1 回 RunPython で呼ぶ。",
            "' 改版番号および履歴: 1.1.0 (2026-06-13) hwnd を渡して register_book。マルチ Excel 初回リボン対策。\n"
            "' プロシージャの動作概要: ensure_python_hosts_ready(hwnd) を 1 回 RunPython で呼ぶ。",
            1,
        )
        print("Main.bas: EnsurePythonHostsReady hwnd updated")

    if RIBBON_NEW.replace("\r\n", "\n") not in norm:
        if RIBBON_OLD.replace("\r\n", "\n") in norm:
            norm = norm.replace(
                RIBBON_OLD.replace("\r\n", "\n"),
                RIBBON_NEW.replace("\r\n", "\n"),
                1,
            )
            print("Main.bas: WaitForm before ensure")
        elif "Call HC_WaitForm.BeginWaitForRibbon(control.ID, act)" not in norm:
            anchor = "    Call HC_RibbonPerf.RibbonPerfMark(\"before_bridge_submit\")"
            norm = norm.replace(
                anchor,
                "    Call HC_WaitForm.BeginWaitForRibbon(control.ID, act)\n"
                "    Call Main.EnsurePythonHostsReady\n"
                "    Call HC_RibbonPerf.RibbonPerfMark(\"before_bridge_submit\")",
                1,
            )
            print("Main.bas: ribbon ensure + wait inserted")

    _write_cp932(MAIN, norm)
    print("done")


if __name__ == "__main__":
    main()
