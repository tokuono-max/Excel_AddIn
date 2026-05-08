Attribute VB_Name = "excel_selection_inspect"
' Dupli-check probe: log Selection / UsedRange / Intersect right after user selects (e.g. sheet corner).
' Import in VBE: File -> Import File -> this .bas, or paste into a standard module.
'
Option Explicit

Public Sub InspectSelectionAfterCorner()
    Dim ws As Worksheet
    Dim sel As Range
    Dim ur As Range
    Dim inter As Range
    Dim msg As String
    Dim logPath As String
    Dim ff As Integer
    Dim i As Long
    Dim a As Range

    On Error GoTo EH

    Set ws = ActiveSheet
    Set sel = Selection

    msg = "=== Selection Inspect " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ===" & vbCrLf
    msg = msg & "Sheet: " & ws.Name & vbCrLf & vbCrLf

    msg = msg & "Selection.Address (A1相対): " & sel.Address(False, False) & vbCrLf

    On Error Resume Next
    msg = msg & "Selection.Address External:=True: " & sel.Address(External:=True) & vbCrLf
    On Error GoTo EH

    msg = msg & "Selection.Rows.Count: " & CStr(sel.Rows.Count) & vbCrLf
    msg = msg & "Selection.Columns.Count: " & CStr(sel.Columns.Count) & vbCrLf
    msg = msg & "Selection.Areas.Count: " & CStr(sel.Areas.Count) & vbCrLf

    On Error Resume Next
    msg = msg & "Selection.Count: " & CStr(sel.Count) & vbCrLf
    msg = msg & "Selection.CountLarge: " & CStr(sel.CountLarge) & vbCrLf
    On Error GoTo EH

    msg = msg & vbCrLf & "--- Sheet limits ---" & vbCrLf
    msg = msg & "Sheet.Rows.Count: " & CStr(ws.Rows.Count) & vbCrLf
    msg = msg & "Sheet.Columns.Count: " & CStr(ws.Columns.Count) & vbCrLf

    On Error Resume Next
    msg = msg & "Sheet.Cells.CountLarge: " & CStr(ws.Cells.CountLarge) & vbCrLf
    On Error GoTo EH

    msg = msg & vbCrLf & "--- UsedRange ---" & vbCrLf
    Set ur = ws.UsedRange
    msg = msg & "UsedRange.Address: " & ur.Address(False, False) & vbCrLf
    msg = msg & "UsedRange.Rows.Count: " & CStr(ur.Rows.Count) & vbCrLf
    msg = msg & "UsedRange.Columns.Count: " & CStr(ur.Columns.Count) & vbCrLf

    msg = msg & vbCrLf & "--- Intersect(Selection, UsedRange) ---" & vbCrLf
    On Error Resume Next
    Set inter = Application.Intersect(sel, ur)
    If inter Is Nothing Then
        msg = msg & "Intersect: Nothing" & vbCrLf
    Else
        msg = msg & "Intersect.Address: " & inter.Address(False, False) & vbCrLf
        msg = msg & "Intersect.Rows.Count: " & CStr(inter.Rows.Count) & vbCrLf
        msg = msg & "Intersect.Columns.Count: " & CStr(inter.Columns.Count) & vbCrLf
    End If
    On Error GoTo EH

    If sel.Areas.Count > 1 And sel.Areas.Count <= 20 Then
        msg = msg & vbCrLf & "--- Areas(1.." & CStr(sel.Areas.Count) & ") 要約 ---" & vbCrLf
        For i = 1 To sel.Areas.Count
            Set a = sel.Areas(i)
            msg = msg & "Areas(" & CStr(i) & ").Address: " & a.Address(False, False) & _
                  " R=" & CStr(a.Rows.Count) & " C=" & CStr(a.Columns.Count) & vbCrLf
        Next i
    End If

    Debug.Print msg

    logPath = Environ$("TEMP") & "\excel_selection_inspect.txt"
    ff = FreeFile
    Open logPath For Append As #ff
    Print #ff, msg
    Print #ff, String$(60, "-") & vbCrLf
    Close #ff

    MsgBox msg & vbCrLf & "ログ追記: " & logPath, vbInformation, "Selection Inspect"
    Exit Sub

EH:
    MsgBox "Err " & CStr(Err.Number) & " " & Err.Description, vbCritical, "Selection Inspect"
End Sub
