# -*- coding: utf-8 -*-
"""One-shot: patch VBA/xlwings.bas (CP932) for packaged RunPython via xlwings_short_runner."""

from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
VBA_PATH = ROOT / "VBA" / "xlwings.bas"
ENC = "cp932"

MARKER_FUNC = "Private Function HcTruthyPackagedRunPython() As Boolean"
OLD_ELSE = (
 "        Else\n"
    "            RunPython = ExecuteWindows(False, pythonCommand, interpreter, PYTHONPATH)\n"
    "        End If"
)

NEW_ELSE = r"""        Else
            If HcTruthyPackagedRunPython() Then
                If Len(Trim$(Environ$("HC_INSTALL_ROOT"))) = 0 Then
                    MsgBox "HC_INSTALL_ROOT is not set. Packaged RunPython requires install root.", vbCritical, "xlwings"
                    RunPython = -1
                ElseIf Not FileExists(interpreter) Then
                    MsgBox "Interpreter not found: " & interpreter, vbCritical, "xlwings"
                    RunPython = -1
                Else
                    Dim scriptPath As String
                    Dim fullPy As String
                    Dim frozenArgs As String
                    fullPy = "import os, sys" & vbLf
                    fullPy = fullPy & "install = (os.environ.get('HC_INSTALL_ROOT') or '').strip()" & vbLf
                    fullPy = fullPy & "if install:" & vbLf
                    fullPy = fullPy & "    if install not in sys.path:" & vbLf
                    fullPy = fullPy & "        sys.path.insert(0, install)" & vbLf
                    fullPy = fullPy & "    try:" & vbLf
                    fullPy = fullPy & "        os.chdir(install)" & vbLf
                    fullPy = fullPy & "    except OSError:" & vbLf
                    fullPy = fullPy & "        pass" & vbLf
                    fullPy = fullPy & "else:" & vbLf
                    fullPy = fullPy & "    root = r'" & Replace(ThisWorkbook.Path, "'", "''") & "'" & vbLf
                    fullPy = fullPy & "    cand = os.path.dirname(root)" & vbLf
                    fullPy = fullPy & "    hc = os.path.join(root, 'hc_main.py')" & vbLf
                    fullPy = fullPy & "    hc2 = os.path.join(cand, 'hc_main.py')" & vbLf
                    fullPy = fullPy & "    if (not os.path.exists(hc)) and os.path.exists(hc2):" & vbLf
                    fullPy = fullPy & "        root = cand" & vbLf
                    fullPy = fullPy & "    if root and root not in sys.path:" & vbLf
                    fullPy = fullPy & "        sys.path.insert(0, root)" & vbLf
                    fullPy = fullPy & pythonCommand
                    scriptPath = GetConfig("TEMP DIR", Environ("Temp")) & "\xlwings-snippet-" & CreateGUID() & ".py"
                    Call HcWriteUtf8File(scriptPath, fullPy)
                    frozenArgs = "--script-file=" & Chr(34) & scriptPath & Chr(34)
                    RunPython = ExecuteWindows(True, interpreter, ParentFolder(interpreter), PYTHONPATH, frozenArgs)
                    On Error Resume Next
                    Kill scriptPath
                    On Error GoTo 0
                End If
            Else
                RunPython = ExecuteWindows(False, pythonCommand, interpreter, PYTHONPATH)
            End If
        End If"""

HELPERS = r"""

Private Function HcTruthyPackagedRunPython() As Boolean
    Dim s As String
    Dim e As String
    s = LCase$(Trim$(CStr(GetConfig("USE_PACKAGED_RUNPYTHON", "False"))))
    HcTruthyPackagedRunPython = (s = "true" Or s = "1" Or s = "yes")
    e = LCase$(Trim$(Environ$("HC_PACKAGED_DEPLOYMENT")))
    If (e = "1" Or e = "true" Or e = "yes") Then
        HcTruthyPackagedRunPython = True
    End If
End Function

Private Sub HcWriteUtf8File(ByVal filePath As String, ByVal content As String)
    Const AD_TYPE_TEXT As Long = 2
    Const AD_SAVE_OVERWRITE As Long = 2
    Const AD_WRITE_CHAR As Long = 0
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = AD_TYPE_TEXT
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText content, AD_WRITE_CHAR
    stm.SaveToFile filePath, AD_SAVE_OVERWRITE
    stm.Close
    Set stm = Nothing
End Sub
"""


def main() -> None:
    text = VBA_PATH.read_text(encoding=ENC)
    if MARKER_FUNC in text:
        print("skip helpers: already patched")
    else:
        needle = "Function ExecuteWindows(IsFrozen As Boolean"
        pos = text.find(needle)
        if pos < 0:
            raise SystemExit("anchor not found: ExecuteWindows")
        text = text[:pos] + HELPERS + text[pos:]
        print("inserted helpers before ExecuteWindows")

    if OLD_ELSE not in text:
        raise SystemExit("OLD_ELSE block not found (already replaced?)")
    text = text.replace(OLD_ELSE, NEW_ELSE, 1)
    print("replaced RunPython ExecuteWindows branch")

    VBA_PATH.write_text(text, encoding=ENC)
    print("wrote", VBA_PATH)


if __name__ == "__main__":
    main()
