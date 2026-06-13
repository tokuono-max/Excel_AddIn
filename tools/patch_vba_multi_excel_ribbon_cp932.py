# -*- coding: utf-8 -*-
"""マルチ Excel 初回リボン対策 + 不足分の Main 復元（cp932 厳守）。"""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
MAIN = ROOT / "VBA" / "Main.bas"

VBA_ONLY_SUB = """\
' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: ShutdownExcelUiVbaOnly
' 改版番号および履歴: 1.0.0 (2026-06-13) BeforeClose 用 VBA のみ片付け（Python 終了は lifecycle monitor に委譲）。
' プロシージャの動作概要: WaitForm / OnTime / Cursor / Interactive を復元する。RunPython は呼ばない。
' 呼出し例: Call Main.ShutdownExcelUiVbaOnly
' ---------------------------------------------------------------------------------------------------------------------
Public Sub ShutdownExcelUiVbaOnly()
    On Error Resume Next
    Call HC_WaitForm.NotifyUiReady
    Call CancelCursorGuardTimer("shutdown")
    Application.Cursor = xlDefault
    Application.Interactive = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Call HC_Log.Info("Main", "ShutdownExcelUiVbaOnly done")
    On Error GoTo 0
End Sub


"""

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

RIBBON_STARTUP_GUARD = """\
    ' # 【目的】Workbook_Open の startup_full RunPython 実行中に 2 本目の RunPython を避ける（マルチ Excel 初回リボン競合）。
    If mWorkbookOpenStartupFullStarted And Not mWorkbookOpenFullPythonDone Then
        Call HC_Log.Info("Main", "RibbonInvoke: startup_full 実行中のためスキップ（完了後に再操作してください）")
        Call HC_RibbonPerf.RibbonPerfEnd
        Exit Sub
    End If

"""

RIBBON_ACT_ANCHOR = "    act = Trim$(control.tag)\n"


def _read_cp932(path: Path) -> str:
    return path.read_text(encoding="cp932")


def _write_cp932(path: Path, text: str) -> None:
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    path.write_bytes(normalized.encode("cp932").replace(b"\n", b"\r\n"))


def _insert_versions(norm: str) -> str:
    if "2.19.0 (2026-06-13)" in norm:
        return norm
    block = (
        "'   2.20.0 (2026-06-13) [起動] startup_full 実行中はリボンをスキップ（RunPython 競合防止）。\n"
        "'   2.19.0 (2026-06-13) [起動] EnsurePythonHostsReady に hwnd 渡し・WaitForm を RunPython 前へ移動。\n"
        "'   2.18.0 (2026-06-13) [起動] リボン操作前に EnsurePythonHostsReady（Python 生存確認）。\n"
        "'   2.17.0 (2026-06-13) [終了] B+A: BeforeClose は ShutdownExcelUiVbaOnly のみ（Python は lifecycle monitor）。\n"
    )
    for anchor in (
        "'   2.16.2 (2026-06-13)",
        "'   2.16.1 (2026-06-13)",
        "'   2.16.0 (2026-06-13)",
        "'   2.15.0 (2026-06-07)",
    ):
        if anchor in norm:
            return norm.replace(anchor, block + anchor, 1)
    raise SystemExit("Main.bas: version anchor not found")


def main() -> None:
    norm = _read_cp932(MAIN).replace("\r\n", "\n")
    norm = _insert_versions(norm)
    norm = norm.replace("' 更新日: 2026-06-07\n", "' 更新日: 2026-06-13\n", 1)

    if "Public Sub ShutdownExcelUiVbaOnly" not in norm:
        anchor = "' Python 単一引用符リテラル用エスケープ"
        if anchor not in norm:
            raise SystemExit("Main.bas: PyEscSq anchor not found")
        norm = norm.replace(anchor, VBA_ONLY_SUB.replace("\r\n", "\n") + anchor, 1)
        print("Main.bas: ShutdownExcelUiVbaOnly added")

    if "Public Sub EnsurePythonHostsReady" not in norm:
        anchor = "' Python 単一引用符リテラル用エスケープ"
        norm = norm.replace(
            anchor,
            ENSURE_SUB.replace("\r\n", "\n") + anchor,
            1,
        )
        print("Main.bas: EnsurePythonHostsReady added")
    else:
        old = (
            '    RunPython "from svc.svc_host import ensure_python_hosts_ready; ensure_python_hosts_ready()"\n'
        )
        new = (
            '    RunPython "from svc.svc_host import ensure_python_hosts_ready; ensure_python_hosts_ready(" & CStr(Application.hwnd) & ")"\n'
        )
        if old in norm:
            norm = norm.replace(old, new, 1)
            print("Main.bas: EnsurePythonHostsReady hwnd updated")

    ribbon_old = (
        "    Call HC_WaitForm.BeginWaitForRibbon(control.ID, act)\n"
        "    Call HC_RibbonPerf.RibbonPerfMark(\"before_bridge_submit\")"
    )
    ribbon_new = (
        "    Call HC_WaitForm.BeginWaitForRibbon(control.ID, act)\n"
        "    Call Main.EnsurePythonHostsReady\n"
        "    Call HC_RibbonPerf.RibbonPerfMark(\"before_bridge_submit\")"
    )
    if "Call Main.EnsurePythonHostsReady" not in norm:
        if ribbon_old in norm:
            norm = norm.replace(ribbon_old, ribbon_new, 1)
            print("Main.bas: ensure inserted after WaitForm")
        else:
            raise SystemExit("Main.bas: ribbon anchor not found")
    elif (
        "    Call Main.EnsurePythonHostsReady\n"
        "    Call HC_WaitForm.BeginWaitForRibbon(control.ID, act)\n" in norm
    ):
        norm = norm.replace(
            "    Call Main.EnsurePythonHostsReady\n"
            "    Call HC_WaitForm.BeginWaitForRibbon(control.ID, act)\n",
            "    Call HC_WaitForm.BeginWaitForRibbon(control.ID, act)\n"
            "    Call Main.EnsurePythonHostsReady\n",
            1,
        )
        print("Main.bas: WaitForm moved before ensure")

    if "2.19.0 (2026-06-13)" in norm and "2.20.0 (2026-06-13)" not in norm:
        norm = norm.replace(
            "'   2.19.0 (2026-06-13)",
            "'   2.20.0 (2026-06-13) [起動] startup_full 実行中はリボンをスキップ（RunPython 競合防止）。\n"
            "'   2.19.0 (2026-06-13)",
            1,
        )
        print("Main.bas: version 2.20.0 added")

    if "startup_full 実行中のためスキップ" not in norm:
        anchor = RIBBON_ACT_ANCHOR.replace("\r\n", "\n")
        if anchor not in norm:
            raise SystemExit("Main.bas: RibbonInvoke act anchor not found")
        norm = norm.replace(
            anchor,
            anchor + RIBBON_STARTUP_GUARD.replace("\r\n", "\n"),
            1,
        )
        print("Main.bas: startup_full ribbon guard added")

    _write_cp932(MAIN, norm)
    print("done")


if __name__ == "__main__":
    main()
