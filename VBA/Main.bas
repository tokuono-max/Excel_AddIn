Attribute VB_Name = "Main"
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' ���W���[����: Main (�W�����W���[��)
' �쐬��: 2025-11-28
' �X�V��: 2026-05-31
' �����R�[�h: �{���W���[���� Shift-JIS�iCP932�j�ŕۑ����邱�Ɓi���{��R�����g�E������̔j���h�~�j�B
' ���Ŕԍ�����ї���:
'   2.9.0 (2026-05-30) [�I��] ShutdownExcelUiCleanup: WaitForm/OnTime/Cursor/Interactive �����iExcel �c���΍�j�B
'   2.10.0 (2026-05-31) [�I��] ShutdownExcelUiCleanup: Python EXCEL_RESTORE �Ăяo���iCOM �n���O�~�ρj�B
'   2.6.0 (2026-04-11) [����] TerminatePython / RunPythonSafe �� RunPython �����񂩂� hc_main �������Bcore.excel_session�iclear_internal_registry / invoke_action�j�o�R�B
'   2.7.0 (2026-04-11) [�d��] check_duplicates: bridge JSON �� selection_areas�i�e Area �� External �A�h���X�j��t�^�B
'   2.8.0 (2026-04-11) check_duplicates: bridge JSON CountLarge (selection_count_large / sheet_cells_count_large).
'   2.5.0 (2026-04-10) [�o�H] ���{���S action �� SubmitSvcRequestViaBridge�ibridge JSON UTF-8�j�� bridge_runner �� svc_server�BRunPythonSafe �͔񃊃{���p�Ɏc���B
'   2.4.1 (2026-04-07) [�����R�[�h] SubmitLoadCsvViaBridge �� JSON �o�͂� ADODB.Stream + Windows-31J(CP932) �����֕ύX�B
'   2.4.0 (2026-04-06) [UX] HC_WaitForm: ���{���`RunPython �O�ɑҋ@ UserForm�BPython ���� HC_WaitForm.NotifyUiReady �ŕ���B
'   2.3.0 (2026-04-06) [�v��] HC_RibbonPerf: ���{���`RunPython ��Ԃ� hc_csv_perf.log �ցiHC_LOG_PERF�j�B
'   2.2.0 (2026-04-06) [�N��] Workbook_Open �� excel_startup_workbook_open_full �� 1 ��̂� RunPython�B
'                          InitPythonServer �͐������X�L�b�v�BManual_Init ���� Reset ��ɏ]���ǂ��� RunPython�B
'   2.1.0 (2026-04-06) [����] ���{���� Main.RibbonCallback_hc_main �̂݁BcustomUI �̊e button �� tag�iaction�j�K�{�B
'                          Call* �n����� Id��action �t�H�[���o�b�N���폜�Bhc_main �� invoke �̂݌��J�����B
'   2.0.0 (2026-04-05) [�݌v] ���{���� tag �܂��� control.Id ���� hc_main.invoke(action=...) �֏W��B
'                          CSV �Ǎ������@�\�Ɠ���o�H�iRunPythonSafe�j�BRibbonCallback_hc_main ��ǉ��B
'   1.9.9 (2026-02-03) [�s��C��] �t�@�C���������� Python �ďo���� "merge_csv" �ɓ����B
'                          �S�Ẵv���V�[�W���w�b�_���K��́u�ڍהŁv�֊��S�����B
'                          �S�ẴR�[���o�b�N�ɕ������O�o�^�p�̃G���[�ߑ����W�b�N�������B
'   1.9.8 (2026-02-02) [�ŏI�d�l�m��] �ʒm��p�L�[ (HC_NOTIFY_RETV) �ƃX�e�[�^�X�̖��������������B
'   1.9.6 (2026-02-01) VBA �哱�̒ʒm���� (CheckAndNotifyVBA) �������B
' �v���V�[�W���̓���T�v: ���{�� customUI �� RibbonCallback_hc_main �� SubmitSvcRequestViaBridge�ibridge �˗� JSON �� UTF-8�j�� bridge_runner �� svc_server�BRunPythonSafe �͔񃊃{���o�H�p�B
' ���ӎ���: Comm �N���X�̎g�p�͌��ցB�ʒm�� MsgBox�A���O�� HC_Log ���g�p���邱�ƁB
'           �}���`�X�e�[�g�����g�i:�j���֎~���A�S�Ă̘_���Ɂu# �y�ړI�z�v�R�����g��t�т�����B
' ---------------------------------------------------------------------------------------------------------------------

' --- ���ʒ萔 ---
' �ϐ�: Python ���ihc_stat.py�j�ƕ�������������ʒm�L�[��
' # �y�ړI�z�������ʂ̃��b�Z�[�W���i�[����v���p�e�B�����`���邽�߁B
Private Const RET_NAME As String = "HC_NOTIFY_RETV"

' �����v�Ď��^�C�}�A�E�g�i�b�j
Private Const CURSOR_GUARD_SEC As Long = 10

' WaitAndInit: Application.OnTime �܂ł̑҂��b�i�b���x�j
Private Const WAIT_INIT_SEC As Long = 1

Private m_cursorGuardTime As Date
Private m_cursorGuardActive As Boolean
Private m_cursorReleased As Boolean

' Workbook_Open �� excel_startup_workbook_open_full ������������ True�i�x�� InitPythonServer �� 2 ��� RunPython ���ȗ��j
Private mWorkbookOpenFullPythonDone As Boolean

' Workbook_Open �� startup_full �� RunPython �𓊂������� True�i�߂�O�� OnTime ��d���s��}�~�j
Private mWorkbookOpenStartupFullStarted As Boolean


' ---------------------------------------------------------------------------------------------------------------------
' Python �P����p�����e�����p�G�X�P�[�v�isheet_id / action �� \ �� ' ���܂܂�Ă� RunPython 1 �s���󂳂Ȃ��j
' ---------------------------------------------------------------------------------------------------------------------
Private Function PyEscSq(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, "'", "\'")
    PyEscSq = t
End Function


' ---------------------------------------------------------------------------------------------------------------------
' �v���V�[�W����: EscapeJsonStringForBridge
' ���J: Private
' �T�v: bridge JSON �� selection_areas �p�B�o�b�N�X���b�V���E�_�u���N�H�[�g�E���䕶�����G�X�P�[�v����B
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
' �v���V�[�W����: BuildCheckDuplicatesSelectionAreasJson
' ���J: Private
' �T�v: check_duplicates ��p�BSelection �� Range �̂Ƃ��A�e Area �� Address(External:=True) �� JSON �z�񃊃e�����i���g�j�ŕԂ��B
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
' �v���V�[�W����: RibbonCallback_hc_main
' ���J: Public�icustomUI �� onAction="Main.RibbonCallback_hc_main" ����̂݌Ă΂��j
' ����: control�iIRibbonControl �����j? control.Tag �� hc_main.invoke �� action ��K�{�Őݒ肷��iCSV_Tool_xml.txt�j�B
' ����: �A�N�e�B�u�V�[�g�� sheet_id ���擾���ARunPythonSafe(act, sId) �� Python �ֈϏ�����B
' ���l: tag ����̂Ƃ��̓��O�ɋL�^���ďI������BcustomUI ���C�����邱�ƁB
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
' �v���V�[�W����: RibbonInvokeFromControl
' ���J: Private�iRibbonCallback_hc_main ����̂݌Ăԁj
' ����: control ? Tag �v���p�e�B�� hc_main.invoke(action=...) �� action �ƈ�v���邱�ƁB
' ---------------------------------------------------------------------------------------------------------------------


Private Sub RibbonInvokeFromControl(ByVal control As Object)
    Dim sId As String
    Dim act As String
    Dim isUpdateCheck As Boolean

    Dim bookFullName As String

    Dim bookName As String

    On Error GoTo ErrorHandler
    act = Trim$(control.tag)
    If Len(act) = 0 Then
        Call HC_Log.Error("Main", "RibbonInvoke: tag ����ł��BcustomUI �� button �� hc_main.invoke �Ɠ���� action �� tag �Ŏw�肵�Ă��������Bcontrol.Id=" & control.ID)
        Call HC_RibbonPerf.RibbonPerfEnd
        Exit Sub
    End If
    isUpdateCheck = (StrComp(act, "check_for_updates", vbTextCompare) = 0)



    If (Not isUpdateCheck) And ActiveSheet Is Nothing Then

        Call HC_RibbonPerf.RibbonPerfEnd

        Exit Sub

    End If

    If Not ActiveSheet Is Nothing Then

        sId = ExcelUtil.GetSheetIdSafe(ActiveSheet)

    Else

        sId = vbNullString

    End If

    ' # �y�ړI�z���{���͑S�� bridge_runner �� svc_server�iRunPython �Z���������j�Bbridge �˗� JSON �� UTF-8�B
    If (Not isUpdateCheck) And ActiveWorkbook Is Nothing Then

        Call HC_Log.Info("Main", "Ribbon bridge: ActiveWorkbook �� Nothing �̂��߃X�L�b�v")
        Call HC_WaitForm.NotifyUiReady

        Call HC_RibbonPerf.RibbonPerfEnd

        Exit Sub

    End If

    If Not ActiveWorkbook Is Nothing Then

        bookFullName = ActiveWorkbook.FullName

        bookName = ActiveWorkbook.Name

    Else

        bookFullName = vbNullString

        bookName = vbNullString

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
    Call Main.SubmitSvcRequestViaBridge(act, Application.hwnd, sId, bookFullName, bookName, selAreasJson, dupliCf)

    Call HC_RibbonPerf.RibbonPerfMark("after_bridge_submit")
    Call HC_RibbonPerf.RibbonPerfEnd
    Exit Sub
ErrorHandler:
    Call HC_WaitForm.NotifyUiReady
    Call HC_RibbonPerf.RibbonPerfEnd
    Call HC_Log.Error("Main", "RibbonInvokeFromControl failed: " & Err.Description)
End Sub


' ==============================================================================
' �������s�G���W���ixlwings RunPython �� hc_main.invoke�j
' ==============================================================================

' ---------------------------------------------------------------------------------------------------------------------
' �v���V�[�W����: RunPythonSafe
' �쐬��: 2026-02-01
' �X�V��: 2026-04-06
' ���Ŕԍ�����ї���:
'   2.1.0 (2026-04-06) �������̈Ӗ����uhc_main.invoke �� action�v�ɖ����BRibbon �o�H�Ƃ̑Ή����w�b�_�ɋL�ځB
'   1.9.9 (2026-02-03) �ڍ׃w�b�_�̓K�p�B��O���O�L�^�̌��i���B
' �v���V�[�W���̓���T�v: xlwings �� RunPython �� 1 �s�� Python �����s����B�\�z���閽�߂͏�� hc_main.invoke �̂݁B
' ����:
'   methodName ? ���{�� control.Tag �Ɠ���Bhc_main.invoke(action=...) �ɓn�� action ������i�� "load_csv"�j�B
'   sId        ? �ΏۃV�[�g���ʎq�iExcelUtil.GetSheetIdSafe�j�Binvoke �� sheet_id �ɓn���B
' �߂�l: �Ȃ�
' �ďo����: Call Main.RunPythonSafe("merge_csv", sheetIdGuid)
' RunPython: from core.excel_session import invoke_action -> hc_main.invoke (no import hc_main in VBA).
' ���㏈��: CheckAndNotifyVBA�iHC_NOTIFY_RETV�j, HC_Bridge.RestoreStatBar, �J�[�\���ی��^�C�}
' ���ӎ���: RunPython ���s���� Excel �����b�Z�[�W���[�v���~�� COM �҂��ɂȂ�BmethodName �ɒP���p�������܂܂�Ă� PyEscSq �ŃG�X�P�[�v�ς݁B
' ---------------------------------------------------------------------------------------------------------------------
Public Sub RunPythonSafe(ByVal methodName As String, ByVal sId As String)
    Dim sCmd As String                                      ' Python ���ߕ�����
    Dim hwnd As LongPtr                                     ' �E�B���h�E�n���h��

    On Error GoTo ErrorHandler

    ' �}�E�X�������v�ɐݒ�
    Call HC_Log.Diag("Main", "Application.Cursor: Wait Cursor on")
    Application.Cursor = xlWait

    ' [�ύX] --- READY�Ď��͔p�~�iOnTime���p�xPolling��COM���l�܂镛��p�������j---
    ' Call StartQtReadyPolling

    ' [�ǉ�] --- �ی��F10�b��ɕK�������v��߂��iPython�ʒm�����Ȃ�/COM���ʂ�Ȃ��ň��P�[�X�΍�j---
    Call StartCursorGuardTimer(sId)

    ' 1. �O�����V�[�P���X�B
    ' # �y�ړI�z���݂� Excel �e�E�B���h�E����肷�邽�߁B
    hwnd = Application.hwnd

    ' # �y�ړI�z���s�J�n�̎�������͏ؐՂƂ��ă��O�֋L�^���邽�߁B
    Call HC_Log.Diag("Main", "RunPythonSafe: Start [" & methodName & "] for HWND: " & hwnd & " ID: " & sId)
    Call HC_Log.Perf("Main", "RunPythonSafe start action=" & methodName & " hwnd=" & hwnd & " sId=" & sId)

    ' # �y�ړI�zOS �̃��b�Z�[�W�L���[�𐮗����A�`���Ԃ����肳���邽�߁B
    DoEvents

    ' 2. ���s���߂̏ڍ׍\�z�B
    ' excel_session.invoke_action -> hc_main.invoke (VBA string avoids import hc_main).
    sCmd = "from core.excel_session import invoke_action; invoke_action(action='" & PyEscSq(methodName) & "', target_hwnd=" & CStr(hwnd) & ", sheet_id='" & PyEscSq(sId) & "')"

    ' 3. �������s�̎��s�B
    ' # �y�ړI�z�O�� Python �v���Z�X�𓯊����s���邽�߁B
    Call HC_RibbonPerf.RibbonPerfMark("before_xlwings_runpython")
    RunPython sCmd
    Call HC_RibbonPerf.RibbonPerfMark("after_xlwings_runpython")

    ' # �y�ړI�z���s����� OS �`����t���b�V�������邽�߁B
    DoEvents

    ' 4. ���㏈���Z�N�V�����iVBA�哱�^�ʒm�j�B
    ' # �y�ړI�zPython �����V�[�g�ɏ����c�����ʒm�p���𔻒肵�\�����邽�߁B
    Call Main.CheckAndNotifyVBA(sId)

    ' # �y�ړI�z�ŐV�̃X�e�[�^�X��� (HC_STATUS_INFO) ���X�e�[�^�X�o�[�֓������f���邽�߁B
    Call HC_Bridge.RestoreStatBar

    ' [��] �����vOFF�́A
    '   - Python����UI�`�抮������ Excel�֒ʒm���Ē���OFF + VBA�̕ی��^�C�}��~�i�����j
    '   - ���ꂪ���s�����ꍇ�́A�ی��^�C�}(ForceCursorOff)��10�b�ŕK��OFF
    ' �Ƃ�����i�\���ɂ���B
    Call HC_RibbonPerf.RibbonPerfEnd
    Exit Sub

ErrorHandler:
    Call HC_WaitForm.NotifyUiReady
    Call HC_RibbonPerf.RibbonPerfMark("after_xlwings_runpython_error")
    Call HC_RibbonPerf.RibbonPerfEnd
    ' # �y�ړI�z�u���b�W���s���s���̏ڍׂ����O�֓o�^���A�ێ琫�����߂邽�߁B
    Call HC_Log.Error("Main", "RunPythonSafe execution Error Number: " & Hex$(Err.Number) & " FAILED: " & Err.Description)
    
    ' �}�E�X�̍����v��߂�
    Call HC_Log.Diag("Main", "Application.Cursor: ErrorHandler Wait Cursor off")
    Application.Cursor = xlDefault

    ' [�ύX] READY�Ď��͔p�~
    ' StopQtReadyPolling

    ' [�ǉ�] �ی��^�C�}���m���Ɏ~�߂�iOnTime�c����h���j
    Call CancelCursorGuardTimer("ErrorHandler:" & sId)

End Sub

' ---------------------------------------------------------------------------------------------------------------------
' �v���V�[�W����: CheckAndNotifyVBA
' �쐬��: 2026-02-01
' ���Ŕԍ�����ї���: 1.1.2 (2026-02-03) �ڍ׃w�b�_�̊����ƈӐ}�R�����g�����B
' �v���V�[�W���̓���T�v: �V�[�g�̒ʒm��p�v���p�e�B����f�[�^���z�����AVBA �Ǝ��� MsgBox ��\������B
' ����: sId (String) - �ΏۃV�[�g�� Base64 GUID
' �߂�l: �Ȃ�
' �ďo����: Call Main.CheckAndNotifyVBA("GUID-STRING")
' �w���p�[�v���V�[�W���̐e�q�֌W: (�q) ExcelUtil.FindSheetById, ExcelUtil.GetSheetProp, ExcelUtil.DeleteSheetProp
' ���ӎ���: �ʒm������A�v���p�e�B�𕨗��I�ɍ폜����d�\����h�~����B
' ---------------------------------------------------------------------------------------------------------------------
Public Sub CheckAndNotifyVBA(ByVal sId As String)
    Dim ws As Worksheet                                     ' �ΏۃV�[�g�I�u�W�F�N�g
    Dim sRaw As String                                      ' �v���p�e�B���f�[�^
    Dim vDat As Variant                                     ' �z��o�b�t�@
    Dim sCap As String                                      ' ���b�Z�[�W���o��
    Dim sMsg As String                                      ' ���b�Z�[�W�{��
    
    On Error GoTo ErrorHandler
    
    ' 1. �ΏۃV�[�g�̓���B
    ' # �y�ړI�zGUID ����Ƀu�b�N��������I�u�W�F�N�g�𕨗����肷�邽�߁B
    Set ws = ExcelUtil.FindSheetById(sId)
    
    ' # �y����z�V�[�g������s�\�ȏꍇ�͏����s�\�Ƃ��ė��E�B
    If ws Is Nothing Then
        Exit Sub
    End If
    
    ' 2. �ʒm���̏ڍ׋z���B
    ' # �y�ړI�zVBA �ʒm��p�ɒ�`���ꂽ�v���p�e�B�l�𒊏o���邽�߁B
    sRaw = ExcelUtil.GetSheetProp(ws, RET_NAME)
    
    ' # �y����z�ʒm���ׂ��f�[�^�����݂��Ȃ��ꍇ�͏I���B
    If sRaw = "" Then
        Exit Sub
    End If
    
    ' 3. ���b�Z�[�W�p�P�b�g�̉�́B
    ' # �y�ړI�zPython ���ŘA�����ꂽ�u�^�C�g��|���e�v���ڍו������邽�߁B
    vDat = Split(sRaw, "|")
    
    ' # �y����z��؂蕶���Ɋ�Â����z��\���̐����������؁B
    If UBound(vDat) >= 1 Then
        sCap = vDat(0)
        ' # �y�⑫�zMsgBox �ł̉��s��L���ɂ��邽�� \n �� vbCrLf �֒u���B
        sMsg = Replace(vDat(1), "\n", vbCrLf)
    Else
        sCap = "CSV Tool �ʒm"
        sMsg = sRaw
    End If

    
    ' 4. �ʒm�̎��s�B
    ' # �y�ړI�z���[�U�[�֎��s���ʂ��_�C�A���O�`���Ŗ������邽�߁B
    On Error Resume Next
    Application.Activate
    AppActivate Application.Caption
    On Error GoTo ErrorHandler
    Call MsgBox(sMsg, vbInformation, sCap)
    
    ' 5. ��Еt���̎��s�B
    ' # �y�ړI�z�v���p�e�B�𕨗��������A����N�����̌�ʒm��h�~���邽�߁B
    Call ExcelUtil.DeleteSheetProp(ws, RET_NAME)
    
    Exit Sub

ErrorHandler:
    ' # �y�ړI�z�ʒm�������ُ̈�����O�ɕߑ����邽�߁B
    Call HC_Log.Error("Main", "CheckAndNotifyVBA encountered an error: " & Err.Description)
End Sub

' ============================================================
' Cursor Guard Timer�i�ی��^�C�}�j
'   - RunPythonSafe�J�n���ɊJ�n
'   - Python����UI�\���������� CancelCursorGuardTimer ���ĂԁiExcel.Run�j
'   - ���s���Ă� 10�b��� ForceCursorOff ���K������
' ============================================================
Public Sub StartCursorGuardTimer(ByVal sId As String)
    ' # �y�ړI�zUI�\�������ʒm���x��/�r�����Ă��A��莞�Ԃō����v��K���������邽�߁B
    ' # �y�݌v�zOnTime�͕b���x�Bms�Ď��͂��Ȃ��i�ߕ��ׂƕ���p�������j�B

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
    ' # �y�ړI�zPython����UI�\�������������_�ŁA�ی��^�C�}���~���邽�߁B
    ' # �y���Ӂz���ɔ��΍ς�/���o�^���ł������Ȃ��悤�ɂ���iOnError Resume Next�j�B

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
    ' # �y�ړI�z�ŏI�ی��Ƃ��č����v����������B
    ' # �y���Ӂz���d���s����Ă����S�i�p���j�ɂ���B

    On Error Resume Next

    If m_cursorReleased Then Exit Sub

    Application.Cursor = xlDefault
    m_cursorReleased = True
    m_cursorGuardActive = False

    Call HC_Log.Diag("Main", "Application.Cursor: OFF (ForceCursorOff)")

    On Error GoTo 0
End Sub



' ==============================================================================
' ���C�t�T�C�N���Ǘ�
' ==============================================================================

' ---------------------------------------------------------------------------------------------------------------------
' �v���V�[�W����: TerminatePython
' �쐬��: 2026-02-02
' ���Ŕԍ�����ї���: 1.0.1 (2026-02-03) �K��w�b�_�̓K�p�B
' �v���V�[�W���̓���T�v: Python ���̃u�b�N�Q�Ǝ��������������A�������������������B
' ����: �Ȃ�
' �߂�l: �Ȃ�
' �ďo����: ThisWorkbook �̏I����
' �w���p�[�v���V�[�W���̐e�q�֌W: (�q) xlwings.RunPython
' ---------------------------------------------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------------
' �v���V�[�W����: ShutdownExcelUiCleanup
' ���Ŕԍ�����ї���:
'   1.1.0 (2026-05-31) Python restore_excel_host_ui_state �Ăяo���EEnableEvents ������ǉ��B
'   1.0.0 (2026-05-30) Excel �I����: WaitForm/OnTime/Interactive/ScreenUpdating �̕����B
' �v���V�[�W���̓���T�v: �A�h�C���I�����O�� VBA ���̑ҋ@ UI �� OnTime ���������AExcel �����Ԃ�߂��B
' �ďo����: Call Main.ShutdownExcelUiCleanup
' �w���p�[�v���V�[�W���̐e�q�֌W: (�q) HC_WaitForm.NotifyUiReady, CancelCursorGuardTimer, xlwings.RunPython
' ---------------------------------------------------------------------------------------------------------------------
Public Sub ShutdownExcelUiCleanup()
    On Error Resume Next
    Dim hwnd As LongPtr
    Dim sId As String
    Dim sCmd As String
    Call HC_WaitForm.NotifyUiReady
    Call CancelCursorGuardTimer("shutdown")
    Application.Cursor = xlDefault
    Application.Interactive = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    hwnd = Application.hwnd
    sId = vbNullString
    If Not ActiveSheet Is Nothing Then
        sId = ExcelUtil.GetSheetIdSafe(ActiveSheet)
    End If
    sCmd = "from core.excel_host_restore import restore_excel_host_ui_state; restore_excel_host_ui_state(" _
        & CStr(hwnd) & ", '" & PyEscSq(sId) & "')"
    RunPython sCmd
    Call HC_Log.Info("Main", "ShutdownExcelUiCleanup done")
    On Error GoTo 0
End Sub

Public Sub TerminatePython()
    On Error Resume Next
    ' # �y�ړI�z�A�h�C���I������ Python ���� COM �Q�Ƃ��N���A���邽�߁B
    Call HC_Log.Info("Main", "TerminatePython: Clearing internal registries.")
    RunPython "from core.excel_session import clear_internal_registry; clear_internal_registry()"
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' �v���V�[�W����: MarkWorkbookOpenFullPythonDone
' �v���V�[�W���̓���T�v: Workbook_Open ���� RunPython�istartup_full�j������ɌĂсA�x���������ł̍Ď��s��}�~����B
' ---------------------------------------------------------------------------------------------------------------------
' ---------------------------------------------------------------------------------------------------------------------
' �v���V�[�W����: MarkWorkbookOpenStartupFullStarted
' �v���V�[�W���̓���T�v: startup_full �� RunPython ����O�ɌĂсA�x�� InitPythonServer �� 2 ��ڂ�}�~����B
' ---------------------------------------------------------------------------------------------------------------------
Public Sub MarkWorkbookOpenStartupFullStarted()
    mWorkbookOpenStartupFullStarted = True
End Sub

Public Function IsWorkbookOpenFullPythonDone() As Boolean
    IsWorkbookOpenFullPythonDone = mWorkbookOpenFullPythonDone
End Function

Public Function IsWorkbookOpenStartupFullStarted() As Boolean
    IsWorkbookOpenStartupFullStarted = mWorkbookOpenStartupFullStarted
End Function

Public Sub MarkWorkbookOpenFullPythonDone()
    mWorkbookOpenFullPythonDone = True
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' �v���V�[�W����: ResetWorkbookOpenFullPythonDone
' �v���V�[�W���̓���T�v: Manual_Init ���� Python ���̍ēo�^��K�v�Ƃ���Ƃ��AInitPythonServer �� RunPython ������B
' ---------------------------------------------------------------------------------------------------------------------
Public Sub ResetWorkbookOpenFullPythonDone()
    mWorkbookOpenFullPythonDone = False
    mWorkbookOpenStartupFullStarted = False
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' �v���V�[�W����: RunInitEvents
' �쐬��: 2026-03-xx
' ���Ŕԍ�����ї���: 1.0.0 (2026-03-xx) ThisWorkbook.InitEvents ��񓯊����s���郉�b�p�[�Ƃ��ĐV�K�쐬�B
' �v���V�[�W���̓���T�v: Application.OnTime ����Ăяo����AThisWorkbook.InitEvents �����S�Ɏ��s����B
'                          ���s���̃G���[�̓��O�ɂ͎c�����A���̏����������։e����^���Ȃ��B
' ����: �Ȃ�
' �߂�l: �Ȃ�
' �ďo����: Main.WaitAndInit�iApplication.OnTime �o�R�j
' ---------------------------------------------------------------------------------------------------------------------
Public Sub RunInitEvents()
    On Error Resume Next
    Call HC_StartupPerf.StartupPerfMark("run_init_events_enter")
    Call ThisWorkbook.InitEvents
    Call HC_StartupPerf.StartupPerfMark("run_init_events_exit")
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' �v���V�[�W����: WaitAndInit
' �쐬��: 2026-02-02
' ���Ŕԍ�����ї���: 1.0.1 (2026-02-03) �K��w�b�_�̓K�p�B
' �v���V�[�W���̓���T�v: Excel �N����̈���҂����s���A�񓯊��ŏ������v���V�[�W����\����s����B
' ����: �Ȃ�
' �߂�l: �Ȃ�
' �ďo����: Workbook_Open
' �w���p�[�v���V�[�W���̐e�q�֌W: (�q) Application.OnTime
' ---------------------------------------------------------------------------------------------------------------------
Public Sub WaitAndInit()
    ' # �y�ړI�zExcel �N������� COM �s�����������AWAIT_INIT_SEC �b��ɏ��������s�����߁B
    Call Application.OnTime(Now + TimeSerial(0, 0, WAIT_INIT_SEC), "Main.RunInitEvents")
    Call HC_Log.Info("Main", "WaitAndInit: Reserved non-blocking initialization.")
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' �v���V�[�W����: InitPythonServer
' �쐬��: 2026-02-02
' ���Ŕԍ�����ї���: 1.0.1 (2026-02-03) �K��w�b�_�̓K�p�B
' �v���V�[�W���̓���T�v: ���݂̃u�b�N�� Python �i�ߓ��ɓo�^���ACOM �ʐM�̓y����\�z����B
' ����: �Ȃ�
' �߂�l: �Ȃ�
' �ďo����: ��������
' �w���p�[�v���V�[�W���̐e�q�֌W: (�q) xlwings.RunPython�i���񐬌���̓X�L�b�v�BManual_Init �ōĎ��s�j
' ---------------------------------------------------------------------------------------------------------------------
Public Sub InitPythonServer()
    On Error Resume Next
    If mWorkbookOpenFullPythonDone Or mWorkbookOpenStartupFullStarted Then
        Call HC_Log.Info("Main", "InitPythonServer: Skipped RunPython (startup_full done or in progress at Workbook_Open).")
        Call HC_StartupPerf.StartupPerfMark("init_python_server_skipped_startup_full_done")
        Exit Sub
    End If
    Call HC_Log.Info("Main", "InitPythonServer: Establishing bridge connection (repair or first-run fallback).")
    Call HC_StartupPerf.StartupPerfMark("init_python_server_before_runpython")
    RunPython "from svc.svc_host import excel_startup_after_excel_idle; excel_startup_after_excel_idle(" & Application.hwnd & ")"
    Call HC_StartupPerf.StartupPerfMark("init_python_server_after_runpython")
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' �v���V�[�W����: SubmitSvcRequestViaBridge
' ���J: Public
' ���Ŕԍ�����ї���:
'   1.0.0 (2026-04-10) ���{�� tag �� action �Ƃ��� JSON �� UTF-8�iADODB.Stream�j�� bridge_requests �ցB
'   1.1.0 (2026-04-11) Optional selectionAreasJson �� selection_areas ��t���\�icheck_duplicates�j�B
' �v���V�[�W���̓���T�v: bridge_runner ���ǂݎ�� svc_server �֓]������˗����o�͂���B
' ���� publicAction: ���{�� control.Tag�ihc_main.invoke �Ɠ��ꕶ����j�B
' ---------------------------------------------------------------------------------------------------------------------
' �⑫: selectionAreasJson �͎��O�G�X�P�[�v�ς݂� JSON �z�񃊃e�����iselection_areas �̒l�j�B
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
' �v���V�[�W����: SubmitLoadCsvViaBridge
' ���J: Public�i�����Ăяo���݊��j
' ����: 2.5.0 (2026-04-10) SubmitSvcRequestViaBridge("load_csv", ...) �ֈϏ��B
' ---------------------------------------------------------------------------------------------------------------------
Public Sub SubmitLoadCsvViaBridge(ByVal hwnd As LongPtr, ByVal sId As String, ByVal bookFullName As String, ByVal bookName As String)
    Call SubmitSvcRequestViaBridge("load_csv", hwnd, sId, bookFullName, bookName)
End Sub

