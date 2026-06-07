Attribute VB_Name = "HC_WaitForm"

Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' モジュール名: HC_WaitForm
' 作成日: 2026-04-06
' 更新日: 2026-06-06
' 文字コード: 本モジュールは Shift-JIS（CP932）で保存すること（日本語コメント・文字列の破損防止）。
' 改版番号および履歴:
'   1.2.0 (2026-06-06) WaitForUiReadySignal: ui_server の .ready 合図を DoEvents で待ち VBA 内で閉じる。
'   1.1.0 (2026-06-01) btnDataAgg: シナリオ画面起動時の WaitForm 表示名を「シナリオ」に統一（確認）。
'   1.0.0 (2026-04-06) 初版: リボン操作後の待機 UserForm。NotifyUiReady / 30秒タイムアウト。
' プロシージャの動作概要: リボン RunPython 前に「準備中」モーダレス表示。Python UI 表示完了で閉じる。
' 注意事項: Comm 禁止。ログは HC_Log。Python ui_server は %TEMP%\csv_tool\waitform\{hwnd}.ready を書く。NotifyUiReady は VBA 内 WaitForUiReadySignal から呼ぶ。
' ---------------------------------------------------------------------------------------------------------------------
Private Const WAIT_TIMEOUT_SEC As Long = 30
Private m_WaitOnTimeFire As Date
Private m_WaitActive As Boolean
' Ribbon button id ごとの待機表示ポリシー（CSV_Tool_xml.txt の button id と一致）

Private Type tRibbonWaitInfo
    ShowWaitForm As Boolean
    DisplayName As String
End Type

Private Function RibbonWaitInfo(ByVal btnId As String) As tRibbonWaitInfo
    Dim r As tRibbonWaitInfo
    Select Case btnId
        Case "btnDataAgg"
            r.ShowWaitForm = True
            r.DisplayName = "シナリオ"
        Case "btnLoadCSV"
            r.ShowWaitForm = True
            r.DisplayName = "CSV読込"
        Case "btnSaveCSV"
            r.ShowWaitForm = True
            r.DisplayName = "CSV保存"
        Case "btnMerge"
            r.ShowWaitForm = True
            r.DisplayName = "ファイル結合"
        Case "btnSplit"
            r.ShowWaitForm = True
            r.DisplayName = "ファイル分割"
        Case "btnSetHeader"
            r.ShowWaitForm = True
            r.DisplayName = "ヘッダ設定"
        Case "btnReleaseHeader"
            r.ShowWaitForm = True
            r.DisplayName = "ヘッダ解除"
        Case "btnAddShukaHeader"
            r.ShowWaitForm = True
            r.DisplayName = "出荷履歴項目追加"
        Case "btnCheckDup"
            r.ShowWaitForm = True
            r.DisplayName = "重複チェック"
        Case "btnDelRows"
            r.ShowWaitForm = True
            r.DisplayName = "空白行削除"
        Case "btnDelCols"
            r.ShowWaitForm = True
            r.DisplayName = "空白列削除"
        Case "btnDateYMD"
            r.ShowWaitForm = True
            r.DisplayName = "年月日変換"
        Case "btnDateYMDHM"
            r.ShowWaitForm = True
            r.DisplayName = "年月日時分変換"
        Case "btnTrimming"
            r.ShowWaitForm = True
            r.DisplayName = "トリミング"
        Case "btnUndo"
            r.ShowWaitForm = True
            r.DisplayName = "アドイン元に戻す"
        Case "btnHelp"
            r.ShowWaitForm = True
            r.DisplayName = "ヘルプ"
        Case "btnCheckUpdates"
            r.ShowWaitForm = False
            r.DisplayName = "更新確認"
        Case Else
            r.ShowWaitForm = True
            r.DisplayName = "処理"

    End Select

    RibbonWaitInfo = r

End Function

' 表示名のみ必要な箇所向け（内部は RibbonWaitInfo と同一の id 対応表）
Public Function RibbonDisplayNameFromControlId(ByVal btnId As String) As String

    RibbonDisplayNameFromControlId = RibbonWaitInfo(btnId).DisplayName

End Function

' ---------------------------------------------------------------------------------------------------------------------
' リボンから RunPython 直前に呼ぶ。連打時は文言更新とタイムアウトをリセット。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub BeginWaitForRibbon(ByVal ribbonControlId As String, ByVal actionTag As String)
    Dim wi As tRibbonWaitInfo
    Dim disp As String

    On Error GoTo CleanFail

    wi = RibbonWaitInfo(ribbonControlId)
    If Not wi.ShowWaitForm Then Exit Sub
    disp = wi.DisplayName
    If m_WaitActive Then
        On Error Resume Next
        Application.OnTime EarliestTime:=m_WaitOnTimeFire, Procedure:="HC_WaitForm.WaitFormTimeout", Schedule:=False
        On Error GoTo CleanFail
    End If

    WaitForm.Label1.Caption = disp & "を準備しています。" & vbCrLf & "しばらくお待ちください。"
    WaitForm.Label2.Caption = "このまま Excel を操作しないでください。"
    Call EnsureWaitFormReadyDir
    Call ClearStaleWaitFormReadySignal(Application.hwnd)
    WaitForm.Show vbModeless
    m_WaitActive = True
    m_WaitOnTimeFire = Now + TimeSerial(0, 0, WAIT_TIMEOUT_SEC)
    Application.OnTime EarliestTime:=m_WaitOnTimeFire, Procedure:="HC_WaitForm.WaitFormTimeout", Schedule:=True

    Exit Sub

CleanFail:

    ' # 【目的】待機フォーム表示失敗時もリボン処理を続けるため。

End Sub

' ---------------------------------------------------------------------------------------------------------------------
' ready 合図 (%TEMP%\csv_tool\waitform\{hwnd}.ready) を ui_server が書く。VBA 内 DoEvents で待ち NotifyUiReady。
' ---------------------------------------------------------------------------------------------------------------------
Private Function WaitFormReadySignalPath(ByVal hwnd As LongPtr) As String
    WaitFormReadySignalPath = Environ$("TEMP") & "\csv_tool\waitform\" & CStr(hwnd) & ".ready"
End Function

Private Sub EnsureWaitFormReadyDir()
    Dim fso As Object
    Dim baseDir As String
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    baseDir = Environ$("TEMP") & "\csv_tool"
    If Not fso.FolderExists(baseDir) Then fso.CreateFolder baseDir
    If Not fso.FolderExists(baseDir & "\waitform") Then fso.CreateFolder baseDir & "\waitform"
    On Error GoTo 0
End Sub

Private Sub ClearStaleWaitFormReadySignal(ByVal hwnd As LongPtr)
    Dim readyPath As String
    Dim fso As Object
    If hwnd = 0 Then Exit Sub
    readyPath = WaitFormReadySignalPath(hwnd)
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(readyPath) Then fso.DeleteFile readyPath, True
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: WaitForUiReadySignal
' 改版番号および履歴:
'   1.0.0 (2026-06-06) bridge 送信後、ready ファイルを DoEvents で待つ（上限 WAIT_TIMEOUT_SEC）。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub WaitForUiReadySignal(ByVal hwnd As LongPtr)
    Dim readyPath As String
    Dim deadline As Date
    Dim fso As Object

    If Not m_WaitActive Then Exit Sub
    If hwnd = 0 Then Exit Sub

    readyPath = WaitFormReadySignalPath(hwnd)
    deadline = Now + TimeSerial(0, 0, WAIT_TIMEOUT_SEC)
    Set fso = CreateObject("Scripting.FileSystemObject")

    Do While m_WaitActive And Now < deadline
        If fso.FileExists(readyPath) Then
            Call HC_Log.Diag("HC_WaitForm", "WaitForUiReadySignal: ready detected path=" & readyPath)
            Call NotifyUiReady
            On Error Resume Next
            If fso.FileExists(readyPath) Then fso.DeleteFile readyPath, True
            On Error GoTo 0
            Exit Sub
        End If
        DoEvents
    Loop

    Call HC_Log.Diag("HC_WaitForm", "WaitForUiReadySignal: no ready before deadline hwnd=" & CStr(hwnd))
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' Python 側から Application.Run で呼ぶ。WaitForm を閉じ、タイムアウトを解除。
' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: NotifyUiReady
' 改版番号および履歴: 1.0.0 (2026-04-06) Python から WaitForm を閉じ、OnTime タイムアウトを解除。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub NotifyUiReady()
    On Error Resume Next
    If m_WaitActive Then
        Application.OnTime EarliestTime:=m_WaitOnTimeFire, Procedure:="HC_WaitForm.WaitFormTimeout", Schedule:=False
    End If
    m_WaitActive = False
    WaitForm.Hide
    On Error GoTo 0

End Sub

' ---------------------------------------------------------------------------------------------------------------------
' OnTime タイムアウト（30 秒で合図が無い場合）
' ---------------------------------------------------------------------------------------------------------------------
Public Sub WaitFormTimeout()
    On Error Resume Next
    m_WaitActive = False
    WaitForm.Hide
    On Error GoTo 0

End Sub

