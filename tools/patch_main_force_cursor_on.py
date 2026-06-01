# -*- coding: utf-8 -*-
"""Insert Main.ForceCursorOn into VBA/Main.bas (binary-safe)."""
from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
TARGET = ROOT / "VBA" / "Main.bas"
NEEDLE = b"Public Sub ForceCursorOff()"
INSERT = (
    b"Public Sub ForceCursorOn(Optional ByVal sId As String = \"batch\")\r\n"
    b"    On Error Resume Next\r\n"
    b"    If Len(sId) = 0 Then sId = \"batch\"\r\n"
    b"    m_cursorReleased = False\r\n"
    b"    Application.Cursor = xlWait\r\n"
    b'    Call HC_Log.Diag("Main", "Application.Cursor: ON (ForceCursorOn)")\r\n'
    b"    Call StartCursorGuardTimer(sId)\r\n"
    b"    On Error GoTo 0\r\n"
    b"End Sub\r\n"
    b"\r\n"
)


def main() -> None:
    data = TARGET.read_bytes()
    if b"Public Sub ForceCursorOn" in data:
        print("already patched:", TARGET)
        return
    if NEEDLE not in data:
        raise SystemExit("needle not found in " + str(TARGET))
    TARGET.write_bytes(data.replace(NEEDLE, INSERT + NEEDLE, 1))
    print("patched:", TARGET)


if __name__ == "__main__":
    main()
