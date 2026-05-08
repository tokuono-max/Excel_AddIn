Attribute VB_Name = "Main"
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' モジュール名: Main (標準モジュール)
' 作成日: 2025-11-28
' 更新日: 2026-04-11
' 文字コード: 本モジュールは Shift-JIS（CP932）で保存すること（日本語コメント・文字列の破損防止）。
' 改版番号および履歴:
'   2.6.0 (2026-04-11) [整理] TerminatePython / RunPythonSafe の RunPython 文字列から hc_main を除去。core.excel_session（clear_internal_registry / invoke_action）経由。
'   2.7.0 (2026-04-11) [重複] check_duplicates: bridge JSON に selection_areas（各 Area の External アドレス）を付与。
'   2.8.0 (2026-04-11) check_duplicates: bridge JSON CountLarge (selection_count_large / sheet_cells_count_large).
'   2.5.0 (2026-04-10) [経路] リボン全 action を SubmitSvcRequestViaBridge（bridge JSON UTF-8）→ bridge_runner → svc_server。RunPythonSafe は非リボン用に残す。
'   2.4.1 (2026-04-07) [文字コード] SubmitLoadCsvViaBridge の JSON 出力を ADODB.Stream + Windows-31J(CP932) 明示へ変更。
'   2.4.0 (2026-04-06) [UX] HC_WaitForm: リボン～RunPython 前に待機 UserForm。Python から HC_WaitForm.NotifyUiReady で閉じる。
'   2.3.0 (2026-04-06) [計測] HC_RibbonPerf: リボン～RunPython 区間を hc_csv_perf.log へ（HC_LOG_PERF）。
'   2.2.0 (2026-04-06) [起動] Workbook_Open で excel_startup_workbook_open_full を 1 回のみ RunPython。
'                          InitPythonServer は成功時スキップ。Manual_Init 時は Reset 後に従来どおり RunPython。
'   2.1.0 (2026-04-06) [整理] リボンは Main.RibbonCallback_hc_main のみ。customUI の各 button に tag（action）必須。
'                          Call* 系および Id→action フォールバックを削除。hc_main は invoke のみ公開入口。
'   2.0.0 (2026-04-05) [設計] リボンは tag または control.Id から hc_main.invoke(action=...) へ集約。
'                          CSV 読込も他機能と同一経路（RunPythonSafe）。RibbonCallback_hc_main を追加。
'   1.9.9 (2026-02-03) [不具合修正] ファイル結合時の Python 呼出名を "merge_csv" に同期。
'                          全てのプロシージャヘッダを規定の「詳細版」へ完全復元。
'                          全てのコールバックに物理ログ登録用のエラー捕捉ロジックを実装。
'   1.9.8 (2026-02-02) [最終仕様確定] 通知専用キー (HC_NOTIFY_RETV) とステータスの役割分離を完遂。
'   1.9.6 (2026-02-01) VBA 主導の通知方式 (CheckAndNotifyVBA) を実装。
' プロシージャの動作概要: リボン customUI → RibbonCallback_hc_main → SubmitSvcRequestViaBridge（bridge 依頼 JSON は UTF-8）→ bridge_runner → svc_server。RunPythonSafe は非リボン経路用。
' 注意事項: Comm クラスの使用は厳禁。通知は MsgBox、ログは HC_Log を使用すること。
'           マルチステートメント（:）を禁止し、全ての論理に「# 【目的】」コメントを付帯させる。
' ---------------------------------------------------------------------------------------------------------------------

' --- 共通定数 ---
' 変数: Python 側（hc_stat.py）と物理同期させる通知キー名
' # 【目的】処理結果のメッセージを格納するプロパティ名を定義するため。
Private Const RET_NAME As String = "HC_NOTIFY_RETV"

' 砂時計監視タイマアウト（秒）
Private Const CURSOR_GUARD_SEC As Long = 10

' WaitAndInit: Application.OnTime までの待ち秒（秒精度）
Private Const WAIT_INIT_SEC As Long = 1

Private m_cursorGuardTime As Date
Private m_cursorGuardActive As Boolean
Private m_cursorReleased As Boolean

' Workbook_Open で excel_startup_workbook_open_full が成功したら True（遅延 InitPythonServer の 2 回目 RunPython を省略）
Private mWorkbookOpenFullPythonDone As Boolean


' ---------------------------------------------------------------------------------------------------------------------
' Python 単一引用符リテラル用エスケープ（sheet_id / action に \ や ' が含まれても RunPython 1 行を壊さない）
' ---------------------------------------------------------------------------------------------------------------------
Private Function PyEscSq(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, "'", "\'")
    PyEscSq = t
End Function


' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: EscapeJsonStringForBridge
' 公開: Private
' 概要: bridge JSON の selection_areas 用。バックスラッシュ・ダブルクォート・制御文字をエスケープする。
' ---------------------------------------------------------------------------------------------------------------------
Private Function EscapeJsonStringForBridge(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, Chr$(34), "\" & Chr$(34))
    t = Replace(t, Chr$(8), "\b")
    t = Replace(t, Chr$(12), "\f")
    t = Replace(t, Chr$(10), "\n")
    t = Replace(t, Chr$(13), "\r")
    t = Replace(t, Chr$(9), "\t")
    EscapeJsonStringForBridge = t
End Function

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: BuildCheckDuplicatesSelectionAreasJson
' 公開: Private
' 概要: check_duplicates 専用。Selection が Range のとき、各 Area の Address(External:=True) を JSON 配列リテラル（中身）で返す。
' ---------------------------------------------------------------------------------------------------------------------
Private Function BuildCheckDuplicatesSelectionAreasJson() As String
    On Error GoTo Fail
    If TypeOf Selection Is Range Then
        Dim rng As Range
        Dim i As Long
        Dim q As String
        Dim frag As String
        Dim oneAddr As String
        Set rng = Selection
        q = Chr$(34)
        frag = "["
        For i = 1 To rng.Areas.Count
            If i > 1 Then frag = frag & ","
            On Error Resume Next
            oneAddr = rng.Areas(i).Address(External:=True)
            On Error GoTo Fail
            frag = frag & q & EscapeJsonStringForBridge(oneAddr) & q
        Next i
        frag = frag & "]"
        BuildCheckDuplicatesSelectionAreasJson = frag
        Exit Function
    End If
Fail:
    BuildCheckDuplicatesSelectionAreasJson = vbNullString
End Function


Private Function BuildCheckDuplicatesCountLargeFragment() As String
    On Error GoTo Fail
    If TypeOf Selection Is Range Then
        Dim rng As Range
        Dim q As String
        Dim selN As Variant
        Dim sheetN As Variant
        Set rng = Selection
        q = Chr$(34)
        selN = rng.CountLarge
        sheetN = ActiveSheet.Cells.CountLarge
        BuildCheckDuplicatesCountLargeFragment = ", " & q & "selection_count_large" & q & ":" & CStr(selN) & ", " & q & "sheet_cells_count_large" & q & ":" & CStr(sheetN)
        Exit Function
    End If
Fail:
    BuildCheckDuplicatesCountLargeFragment = vbNullString
End Function

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: RibbonCallback_hc_main
' 公開: Public（customUI の onAction="Main.RibbonCallback_hc_main" からのみ呼ばれる）
' 引数: control（IRibbonControl 相当）? control.Tag に hc_main.invoke の action を必須で設定する（CSV_Tool_xml.txt）。
' 処理: アクティブシートの sheet_id を取得し、RunPythonSafe(act, sId) で Python へ委譲する。
' 備考: tag が空のときはログに記録して終了する。customUI を修正すること。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub RibbonCallback_hc_main(ByVal control As Object)
    On Error GoTo ErrorHandler
    Call HC_RibbonPerf.RibbonPerfBegin
    Call RibbonInvokeFromControl(control)
    Exit Sub
ErrorHandler:
    Call HC_WaitForm.NotifyUiReady
    Call HC_RibbonPerf.RibbonPerfEnd
    Call HC_Log.Error("Main", "RibbonCallback_hc_main failed: " & Err.Description)
End Sub


' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: RibbonInvokeFromControl
' 公開: Private（RibbonCallback_hc_main からのみ呼ぶ）
' 引数: control ? Tag プロパティが hc_main.invoke(action=...) の action と一致すること。
' ---------------------------------------------------------------------------------------------------------------------
' Packaged update: RunPython only (not svc bridge). tag=check_for_updates in CSV_Tool_xml.txt
Private Sub RunPythonCheckForUpdatesRibbon()
    On Error GoTo Fail
    Application.Cursor = xlWait
    RunPython "from core.packaged_update import check_for_updates_interactive; check_for_updates_interactive('ribbon')"
    GoTo Done
Fail:
    Call HC_Log.Error("Main", "RunPythonCheckForUpdatesRibbon: " & Err.Description)
Done:
    Application.Cursor = xlDefault
End Sub


Private Sub RibbonInvokeFromControl(ByVal control As Object)
    Dim sId As String
    Dim act As String
    On Error GoTo ErrorHandler
    act = Trim$(control.tag)
    If Len(act) = 0 Then
        Call HC_Log.Error("Main", "RibbonInvoke: tag が空です。customUI の button に hc_main.invoke と同一の action を tag で指定してください。control.Id=" & control.ID)
        Call HC_RibbonPerf.RibbonPerfEnd
        Exit Sub
    End If
    If StrComp(act, "check_for_updates", vbTextCompare) = 0 Then
        Call HC_WaitForm.BeginWaitForRibbon(control.ID, act)
        Call RunPythonCheckForUpdatesRibbon
        Call HC_WaitForm.NotifyUiReady
        Call HC_RibbonPerf.RibbonPerfEnd
        Exit Sub
    End If
    If ActiveSheet Is Nothing Then
        Call HC_RibbonPerf.RibbonPerfEnd
        Exit Sub
    End If
    sId = ExcelUtil.GetSheetIdSafe(ActiveSheet)
    ' # 【目的】リボンは全て bridge_runner → svc_server（RunPython 短命を避ける）。bridge 依頼 JSON は UTF-8。
    If ActiveWorkbook Is Nothing Then
        Call HC_Log.Info("Main", "Ribbon bridge: ActiveWorkbook が Nothing のためスキップ")
        Call HC_WaitForm.NotifyUiReady
        Call HC_RibbonPerf.RibbonPerfEnd
        Exit Sub
    End If
    Dim selAreasJson As String
    Dim dupliCf As String
    selAreasJson = vbNullString
    dupliCf = vbNullString
    If StrComp(act, "check_duplicates", vbTextCompare) = 0 Then
        selAreasJson = BuildCheckDuplicatesSelectionAreasJson()
        dupliCf = BuildCheckDuplicatesCountLargeFragment()
    End If
    Call HC_WaitForm.BeginWaitForRibbon(control.ID, act)
    Call HC_RibbonPerf.RibbonPerfMark("before_bridge_submit")
    Call Main.SubmitSvcRequestViaBridge(act, Application.hwnd, sId, ActiveWorkbook.FullName, ActiveWorkbook.Name, selAreasJson, dupliCf)
    Call HC_RibbonPerf.RibbonPerfMark("after_bridge_submit")
    Call HC_RibbonPerf.RibbonPerfEnd
    Exit Sub
ErrorHandler:
    Call HC_WaitForm.NotifyUiReady
    Call HC_RibbonPerf.RibbonPerfEnd
    Call HC_Log.Error("Main", "RibbonInvokeFromControl failed: " & Err.Description)
End Sub


' ==============================================================================
' 物理実行エンジン（xlwings RunPython → hc_main.invoke）
' ==============================================================================

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: RunPythonSafe
' 作成日: 2026-02-01
' 更新日: 2026-04-06
' 改版番号および履歴:
'   2.1.0 (2026-04-06) 引数名の意味を「hc_main.invoke の action」に明示。Ribbon 経路との対応をヘッダに記載。
'   1.9.9 (2026-02-03) 詳細ヘッダの適用。例外ログ記録の厳格化。
' プロシージャの動作概要: xlwings の RunPython で 1 行の Python を実行する。構築する命令は常に hc_main.invoke のみ。
' 引数:
'   methodName ? リボン control.Tag と同一。hc_main.invoke(action=...) に渡す action 文字列（例 "load_csv"）。
'   sId        ? 対象シート識別子（ExcelUtil.GetSheetIdSafe）。invoke の sheet_id に渡す。
' 戻り値: なし
' 呼出し例: Call Main.RunPythonSafe("merge_csv", sheetIdGuid)
' RunPython: from core.excel_session import invoke_action -> hc_main.invoke (no import hc_main in VBA).
' 事後処理: CheckAndNotifyVBA（HC_NOTIFY_RETV）, HC_Bridge.RestoreStatBar, カーソル保険タイマ
' 注意事項: RunPython 実行中は Excel がメッセージループを止め COM 待ちになる。methodName に単引用符等が含まれても PyEscSq でエスケープ済み。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub RunPythonSafe(ByVal methodName As String, ByVal sId As String)
    Dim sCmd As String                                      ' Python 命令文字列
    Dim hwnd As LongPtr                                     ' ウィンドウハンドル

    On Error GoTo ErrorHandler

    ' マウスを砂時計に設定
    Call HC_Log.Diag("Main", "Application.Cursor: Wait Cursor on")
    Application.Cursor = xlWait

    ' [変更] --- READY監視は廃止（OnTime高頻度PollingでCOMが詰まる副作用を避ける）---
    ' Call StartQtReadyPolling

    ' [追加] --- 保険：10秒後に必ず砂時計を戻す（Python通知が来ない/COMが通らない最悪ケース対策）---
    Call StartCursorGuardTimer(sId)

    ' 1. 前処理シーケンス。
    ' # 【目的】現在の Excel 親ウィンドウを特定するため。
    hwnd = Application.hwnd

    ' # 【目的】実行開始の事実を解析証跡としてログへ記録するため。
    Call HC_Log.Diag("Main", "RunPythonSafe: Start [" & methodName & "] for HWND: " & hwnd & " ID: " & sId)
    Call HC_Log.Perf("Main", "RunPythonSafe start action=" & methodName & " hwnd=" & hwnd & " sId=" & sId)

    ' # 【目的】OS のメッセージキューを整理し、描画状態を安定させるため。
    DoEvents

    ' 2. 実行命令の詳細構築。
    ' excel_session.invoke_action -> hc_main.invoke (VBA string avoids import hc_main).
    sCmd = "from core.excel_session import invoke_action; invoke_action(action='" & PyEscSq(methodName) & "', target_hwnd=" & CStr(hwnd) & ", sheet_id='" & PyEscSq(sId) & "')"

    ' 3. 物理実行の執行。
    ' # 【目的】外部 Python プロセスを同期実行するため。
    Call HC_RibbonPerf.RibbonPerfMark("before_xlwings_runpython")
    RunPython sCmd
    Call HC_RibbonPerf.RibbonPerfMark("after_xlwings_runpython")

    ' # 【目的】実行直後の OS 描画をフラッシュさせるため。
    DoEvents

    ' 4. 事後処理セクション（VBA主導型通知）。
    ' # 【目的】Python 側がシートに書き残した通知用情報を判定し表示するため。
    Call Main.CheckAndNotifyVBA(sId)

    ' # 【目的】最新のステータス情報 (HC_STATUS_INFO) をステータスバーへ同期反映するため。
    Call HC_Bridge.RestoreStatBar

    ' [注] 砂時計OFFは、
    '   - Python側がUI描画完了時に Excelへ通知して直接OFF + VBAの保険タイマ停止（推奨）
    '   - それが失敗した場合は、保険タイマ(ForceCursorOff)が10秒で必ずOFF
    ' という二段構えにする。
    Call HC_RibbonPerf.RibbonPerfEnd
    Exit Sub

ErrorHandler:
    Call HC_WaitForm.NotifyUiReady
    Call HC_RibbonPerf.RibbonPerfMark("after_xlwings_runpython_error")
    Call HC_RibbonPerf.RibbonPerfEnd
    ' # 【目的】ブリッジ実行失敗時の詳細をログへ登録し、保守性を高めるため。
    Call HC_Log.Error("Main", "RunPythonSafe execution Error Number: " & Hex$(Err.Number) & " FAILED: " & Err.Description)
    
    ' マウスの砂時計を戻す
    Call HC_Log.Diag("Main", "Application.Cursor: ErrorHandler Wait Cursor off")
    Application.Cursor = xlDefault

    ' [変更] READY監視は廃止
    ' StopQtReadyPolling

    ' [追加] 保険タイマを確実に止める（OnTime残留を防ぐ）
    Call CancelCursorGuardTimer("ErrorHandler:" & sId)

End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: CheckAndNotifyVBA
' 作成日: 2026-02-01
' 改版番号および履歴: 1.1.2 (2026-02-03) 詳細ヘッダの完遂と意図コメント強化。
' プロシージャの動作概要: シートの通知専用プロパティからデータを吸引し、VBA 独自の MsgBox を表示する。
' 引数: sId (String) - 対象シートの Base64 GUID
' 戻り値: なし
' 呼出し例: Call Main.CheckAndNotifyVBA("GUID-STRING")
' ヘルパープロシージャの親子関係: (子) ExcelUtil.FindSheetById, ExcelUtil.GetSheetProp, ExcelUtil.DeleteSheetProp
' 注意事項: 通知完了後、プロパティを物理的に削除し二重表示を防止する。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub CheckAndNotifyVBA(ByVal sId As String)
    Dim ws As Worksheet                                     ' 対象シートオブジェクト
    Dim sRaw As String                                      ' プロパティ生データ
    Dim vDat As Variant                                     ' 配列バッファ
    Dim sCap As String                                      ' メッセージ見出し
    Dim sMsg As String                                      ' メッセージ本文
    
    On Error GoTo ErrorHandler
    
    ' 1. 対象シートの特定。
    ' # 【目的】GUID を基準にブック内から実オブジェクトを物理特定するため。
    Set ws = ExcelUtil.FindSheetById(sId)
    
    ' # 【判定】シートが特定不能な場合は処理不能として離脱。
    If ws Is Nothing Then
        Exit Sub
    End If
    
    ' 2. 通知情報の詳細吸引。
    ' # 【目的】VBA 通知専用に定義されたプロパティ値を抽出するため。
    sRaw = ExcelUtil.GetSheetProp(ws, RET_NAME)
    
    ' # 【判定】通知すべきデータが存在しない場合は終了。
    If sRaw = "" Then
        Exit Sub
    End If
    
    ' 3. メッセージパケットの解析。
    ' # 【目的】Python 側で連結された「タイトル|内容」を詳細分解するため。
    vDat = Split(sRaw, "|")
    
    ' # 【判定】区切り文字に基づいた配列構造の正当性を検証。
    If UBound(vDat) >= 1 Then
        sCap = vDat(0)
        ' # 【補足】MsgBox での改行を有効にするため \n を vbCrLf へ置換。
        sMsg = Replace(vDat(1), "\n", vbCrLf)
    Else
        sCap = "CSV Tool 通知"
        sMsg = sRaw
    End If

    
    ' 4. 通知の執行。
    ' # 【目的】ユーザーへ実行結果をダイアログ形式で明示するため。
    On Error Resume Next
    Application.Activate
    AppActivate Application.Caption
    On Error GoTo ErrorHandler
    Call MsgBox(sMsg, vbInformation, sCap)
    
    ' 5. 後片付けの執行。
    ' # 【目的】プロパティを物理抹消し、次回起動時の誤通知を防止するため。
    Call ExcelUtil.DeleteSheetProp(ws, RET_NAME)
    
    Exit Sub

ErrorHandler:
    ' # 【目的】通知処理中の異常をログに捕捉するため。
    Call HC_Log.Error("Main", "CheckAndNotifyVBA encountered an error: " & Err.Description)
End Sub

' ============================================================
' Cursor Guard Timer（保険タイマ）
'   - RunPythonSafe開始時に開始
'   - Python側がUI表示完了時に CancelCursorGuardTimer を呼ぶ（Excel.Run）
'   - 失敗しても 10秒後に ForceCursorOff が必ず走る
' ============================================================
Public Sub StartCursorGuardTimer(ByVal sId As String)
    ' # 【目的】UI表示完了通知が遅延/喪失しても、一定時間で砂時計を必ず解除するため。
    ' # 【設計】OnTimeは秒精度。ms監視はしない（過負荷と副作用を避ける）。

    On Error Resume Next

    m_cursorReleased = False
    m_cursorGuardActive = True
    m_cursorGuardTime = Now + TimeSerial(0, 0, CURSOR_GUARD_SEC)

    Call HC_Log.Diag("Main", "CursorGuard: start timer +" & CStr(CURSOR_GUARD_SEC) & "s for ID: " & sId)

    Application.OnTime _
        EarliestTime:=m_cursorGuardTime, _
        Procedure:="ForceCursorOff", _
        Schedule:=True

    On Error GoTo 0
End Sub

Public Sub CancelCursorGuardTimer(Optional ByVal reason As String = "")
    ' # 【目的】Python側がUI表示完了した時点で、保険タイマを停止するため。
    ' # 【注意】既に発火済み/未登録等でも落ちないようにする（OnError Resume Next）。

    On Error Resume Next

    If m_cursorGuardActive Then
        Application.OnTime _
            EarliestTime:=m_cursorGuardTime, _
            Procedure:="ForceCursorOff", _
            Schedule:=False
    End If

    m_cursorGuardActive = False

    Call HC_Log.Diag("Main", "CursorGuard: timer cancelled. reason=" & reason)

    On Error GoTo 0
End Sub

Public Sub ForceCursorOff()
    ' # 【目的】最終保険として砂時計を解除する。
    ' # 【注意】多重実行されても安全（冪等）にする。

    On Error Resume Next

    If m_cursorReleased Then Exit Sub

    Application.Cursor = xlDefault
    m_cursorReleased = True
    m_cursorGuardActive = False

    Call HC_Log.Diag("Main", "Application.Cursor: OFF (ForceCursorOff)")

    On Error GoTo 0
End Sub



' ==============================================================================
' ライフサイクル管理
' ==============================================================================

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: TerminatePython
' 作成日: 2026-02-02
' 改版番号および履歴: 1.0.1 (2026-02-03) 規定ヘッダの適用。
' プロシージャの動作概要: Python 側のブック参照辞書を初期化し、メモリ資源を解放する。
' 引数: なし
' 戻り値: なし
' 呼出し例: ThisWorkbook の終了時
' ヘルパープロシージャの親子関係: (子) xlwings.RunPython
' ---------------------------------------------------------------------------------------------------------------------
Public Sub TerminatePython()
    On Error Resume Next
    ' # 【目的】アドイン終了時に Python 側の COM 参照をクリアするため。
    Call HC_Log.Info("Main", "TerminatePython: Clearing internal registries.")
    RunPython "from core.excel_session import clear_internal_registry; clear_internal_registry()"
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: MarkWorkbookOpenFullPythonDone
' プロシージャの動作概要: Workbook_Open 内の RunPython（startup_full）成功後に呼び、遅延初期化での再実行を抑止する。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub MarkWorkbookOpenFullPythonDone()
    mWorkbookOpenFullPythonDone = True
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: ResetWorkbookOpenFullPythonDone
' プロシージャの動作概要: Manual_Init 等で Python 側の再登録を必要とするとき、InitPythonServer で RunPython させる。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub ResetWorkbookOpenFullPythonDone()
    mWorkbookOpenFullPythonDone = False
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: RunInitEvents
' 作成日: 2026-03-xx
' 改版番号および履歴: 1.0.0 (2026-03-xx) ThisWorkbook.InitEvents を非同期実行するラッパーとして新規作成。
' プロシージャの動作概要: Application.OnTime から呼び出され、ThisWorkbook.InitEvents を安全に実行する。
'                          実行時のエラーはログには残さず、次の初期化処理へ影響を与えない。
' 引数: なし
' 戻り値: なし
' 呼出し例: Main.WaitAndInit（Application.OnTime 経由）
' ---------------------------------------------------------------------------------------------------------------------
Public Sub RunInitEvents()
    On Error Resume Next
    Call HC_StartupPerf.StartupPerfMark("run_init_events_enter")
    Call ThisWorkbook.InitEvents
    Call HC_StartupPerf.StartupPerfMark("run_init_events_exit")
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: WaitAndInit
' 作成日: 2026-02-02
' 改版番号および履歴: 1.0.1 (2026-02-03) 規定ヘッダの適用。
' プロシージャの動作概要: Excel 起動後の安定待ちを行い、非同期で初期化プロシージャを予約実行する。
' 引数: なし
' 戻り値: なし
' 呼出し例: Workbook_Open
' ヘルパープロシージャの親子関係: (子) Application.OnTime
' ---------------------------------------------------------------------------------------------------------------------
Public Sub WaitAndInit()
    ' # 【目的】Excel 起動直後の COM 不安定期を避け、WAIT_INIT_SEC 秒後に初期化を行うため。
    Call Application.OnTime(Now + TimeSerial(0, 0, WAIT_INIT_SEC), "Main.RunInitEvents")
    Call HC_Log.Info("Main", "WaitAndInit: Reserved non-blocking initialization.")
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: InitPythonServer
' 作成日: 2026-02-02
' 改版番号および履歴: 1.0.1 (2026-02-03) 規定ヘッダの適用。
' プロシージャの動作概要: 現在のブックを Python 司令塔に登録し、COM 通信の土台を構築する。
' 引数: なし
' 戻り値: なし
' 呼出し例: 初期化時
' ヘルパープロシージャの親子関係: (子) xlwings.RunPython（初回成功後はスキップ。Manual_Init で再実行）
' ---------------------------------------------------------------------------------------------------------------------
Public Sub InitPythonServer()
    On Error Resume Next
    If mWorkbookOpenFullPythonDone Then
        Call HC_Log.Info("Main", "InitPythonServer: Skipped RunPython (startup_full already ran at Workbook_Open).")
        Call HC_StartupPerf.StartupPerfMark("init_python_server_skipped_startup_full_done")
        Exit Sub
    End If
    Call HC_Log.Info("Main", "InitPythonServer: Establishing bridge connection (repair or first-run fallback).")
    Call HC_StartupPerf.StartupPerfMark("init_python_server_before_runpython")
    RunPython "from svc.svc_host import excel_startup_after_excel_idle; excel_startup_after_excel_idle(" & Application.hwnd & ")"
    Call HC_StartupPerf.StartupPerfMark("init_python_server_after_runpython")
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: SubmitSvcRequestViaBridge
' 公開: Public
' 改版番号および履歴:
'   1.0.0 (2026-04-10) リボン tag を action とした JSON を UTF-8（ADODB.Stream）で bridge_requests へ。
'   1.1.0 (2026-04-11) Optional selectionAreasJson で selection_areas を付加可能（check_duplicates）。
' プロシージャの動作概要: bridge_runner が読み取り svc_server へ転送する依頼を出力する。
' 引数 publicAction: リボン control.Tag（hc_main.invoke と同一文字列）。
' ---------------------------------------------------------------------------------------------------------------------
' 補足: selectionAreasJson は事前エスケープ済みの JSON 配列リテラル（selection_areas の値）。
Public Sub SubmitSvcRequestViaBridge(ByVal publicAction As String, ByVal hwnd As LongPtr, ByVal sId As String, ByVal bookFullName As String, ByVal bookName As String, Optional ByVal selectionAreasJson As String = "", Optional ByVal dupliCountsFragment As String = "")
    Dim baseDir As String
    Dim reqStamp As String
    Dim reqTmp As String
    Dim reqPath As String
    Dim jsonStr As String
    Dim fso As Object
    Dim stm As Object
    Dim q As String
    Dim escAct As String
    Dim escFull As String
    Dim escName As String
    Dim escId As String

    On Error GoTo ErrBridge

    q = Chr$(34)
    escAct = Replace(publicAction, q, "\" & q)
    escFull = Replace(Replace(bookFullName, "\", "\\"), q, "\" & q)
    escName = Replace(Replace(bookName, "\", "\\"), q, "\" & q)
    escId = Replace(sId, q, "\" & q)

    jsonStr = "{" _
              & q & "action" & q & ":" & q & escAct & q & "," _
              & q & "hwnd" & q & ":" & CStr(hwnd) & "," _
              & q & "sheet_id" & q & ":" & q & escId & q & "," _
              & q & "book_fullname" & q & ":" & q & escFull & q & "," _
              & q & "book_name" & q & ":" & q & escName & q & "}"
    If Len(selectionAreasJson) > 0 Or Len(dupliCountsFragment) > 0 Then
        jsonStr = Left$(jsonStr, Len(jsonStr) - 1)
        If Len(selectionAreasJson) > 0 Then
            jsonStr = jsonStr & "," & q & "selection_areas" & q & ":" & selectionAreasJson
        End If
        If Len(dupliCountsFragment) > 0 Then
            jsonStr = jsonStr & dupliCountsFragment
        End If
        jsonStr = jsonStr & "}"
    End If

    baseDir = Environ$("TEMP") & "\csv_tool\bridge_requests"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(Environ$("TEMP") & "\csv_tool") Then
        fso.CreateFolder Environ$("TEMP") & "\csv_tool"
    End If
    If Not fso.FolderExists(baseDir) Then
        fso.CreateFolder baseDir
    End If

    reqStamp = Format$(Now, "yyyymmddhhnnss")
    reqTmp = baseDir & "\req_" & reqStamp & ".tmp"
    reqPath = baseDir & "\req_" & reqStamp & ".json"

    If fso.FileExists(reqTmp) Then
        fso.DeleteFile reqTmp
    End If
    If fso.FileExists(reqPath) Then
        fso.DeleteFile reqPath
    End If

    Const AD_TYPE_TEXT As Long = 2
    Const AD_SAVE_OVERWRITE As Long = 2
    Const AD_WRITE_CHAR As Long = 0
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = AD_TYPE_TEXT
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText jsonStr, AD_WRITE_CHAR
    stm.SaveToFile reqTmp, AD_SAVE_OVERWRITE
    stm.Close
    Set stm = Nothing
    fso.MoveFile reqTmp, reqPath

    Exit Sub

ErrBridge:
    Dim errNum As Long
    Dim errSrc As String
    Dim errDesc As String
    errNum = Err.Number
    errSrc = Err.source
    errDesc = Err.Description
    Call HC_Log.Error("Main", "SubmitSvcRequestViaBridge failed: err=" & CStr(errNum) & " hex=&H" & Hex$(errNum) & " desc=" & errDesc)
    Err.Raise errNum, errSrc, errDesc
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: SubmitLoadCsvViaBridge
' 公開: Public（既存呼び出し互換）
' 改版: 2.5.0 (2026-04-10) SubmitSvcRequestViaBridge("load_csv", ...) へ委譲。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub SubmitLoadCsvViaBridge(ByVal hwnd As LongPtr, ByVal sId As String, ByVal bookFullName As String, ByVal bookName As String)
    Call SubmitSvcRequestViaBridge("load_csv", hwnd, sId, bookFullName, bookName)
End Sub

