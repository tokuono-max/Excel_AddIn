VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' ���W���[����: ThisWorkbook (�N���X���W���[��)
' �쐬��: 2025-11-28
' �X�V��: 2026-06-13
' ���Ŕԍ�����ї���:
'   1.10.0 (2026-06-13) [�I��] B+A: BeforeClose �� VBA �̂ݕЕt���i�x�� shutdown �p�~�j�B
'   1.9.9 (2026-06-13) [�I��] BeforeClose �x�� shutdown�i�I���L�����Z������ Python �풓�ێ��j�B
'   1.9.8 (2026-06-07) [�N��] LogInstalledAddins �� startup_full RunPython ��֒x���B
'   1.9.7 (2026-06-01) [�ێ�] ���W���[���R�����g�E�{���� CP932 ���{���畜���iUTF-8 �����ɂ�镶�������C���j�B
'   1.9.6 (2026-05-30) [�I��] BeforeClose �� Main.ShutdownExcelUiCleanup ���Ăяo���iOnTime �c���EInteractive �����j�B
'   1.9.5 (2026-02-01) [�ŏI�d�l�m��] �N�����̊����iVer/Path�j����ёS�A�h�C���������O�@�\�����S�����B
'                          VBA�J���K��i��ȗ��A1�s1���߁A�w�b�_�����AComm�֎~�AHC_Log��{���j��S��K�p�B
'   1.8.0 (2026-02-01) �N�����ɑS�A�h�C���̓Ǎ���Ԃ𑖍������O�o�͂���@�\�������B
'   1.7.0 (2026-02-01) ���_���A�ɔ����A�h�C�����[�h�E�I���V�[�P���X�̃��O�L�^�����B
' �v���V�[�W���̓���T�v: �A�h�C���̕����I�Ȑ������ԁi���C�t�T�C�N���j���Ď��E���䂷��B
'                          Excel �N�����̊������W�A�Ď��Z���T�[�̋N���A����яI�����̃N���[���A�b�v���哱����B
' ���ӎ���: Comm �N���X�̎g�p�͌��ցB���b�Z�[�W�ʒm�͕W�� MsgBox�A���O�� HC_Log ���g�p���邱�ƁB
'           �}���`�X�e�[�g�����g�i:�j���֎~���A�S�Ă̘_���Ɂu# �y�ړI�z�v�R�����g��t�т�����B
'           �ϐ����͒Z�k�`�imSensor ���j���ێ����邱�ƁB
' ---------------------------------------------------------------------------------------------------------------------

' --- ���W���[�����x���ϐ���` ---
' # �y�ړI�zExcel �A�v���P�[�V�����S�̂̃C�x���g��ߑ��������邽�߁B
Private mSensor As AppEventClass                            ' �A�v���P�[�V�����Ď��Z���T�[�̃C���X�^���X

' ==============================================================================
' �u�b�N���x���E�C�x���g�n���h�� (�A�h�C���E�G���g���|�C���g)
' ==============================================================================

' ---------------------------------------------------------------------------------------------------------------------
' �R�[���o�b�N��: Workbook_Open
' �쐬��: 2026-02-01
' ���Ŕԍ�����ї���: 1.3.0 (2026-02-01) �����̎��W�ƃA�h�C�������̌Ăяo����ǉ��B
' �v���V�[�W���̓���T�v: �A�h�C���ǂݍ��ݎ��Ɏ������s����A�����̎��W�Ə������\����s���B
' ����: �Ȃ�
' �߂�l: �Ȃ�
' �ďo����: Excel �ɂ��A�h�C�����[�h���Ɏ������s�B
' �w���p�[�v���V�[�W���̐e�q�֌W: (�q) HC_Log.InsertBlankLine, HC_Log.Info, ThisWorkbook.LogEnvironmentInfo,
'                                  ThisWorkbook.LogInstalledAddins, Main.WaitAndInit
' ���ӎ���: Python�isvc/ui/bridge/register_book�j�� Workbook_Open ���� 1 ��� RunPython �Ŋ����B
'           �C�x���g�Z���T�[�ڑ��̂� Main.WaitAndInit �� RunInitEvents �Œx������B
' ---------------------------------------------------------------------------------------------------------------------
Private Sub Workbook_Open()
    On Error Resume Next
    
    ' 0. �N���v���iHC_LOG_PERF=1 �̂Ƃ��̂� hc_csv_perf.log�j�B
    ' # �y�ړI�zWorkbook_Open ����x�������������܂ł̌o�߂𓯈������ŒǐՂ��邽�߁B
    Call HC_StartupPerf.StartupPerfReset
    
    ' 1. �������O�̏������B
    ' # �y�ړI�z���O�t�@�C����ŐV���� Excel �Z�b�V�������J�n���ꂽ���Ƃ����o�I�ɋ�ʂ��邽�߁B
    Call HC_Log.InsertBlankLine
    Call HC_StartupPerf.StartupPerfMark("after_insert_blank_line")
    
    ' 2. �N���J�n�̋L�^�B
    ' # �y�ړI�z�A�h�C�����V�X�e���ɓǂݍ��܂ꂽ���m�Ȏ����Ǝ����𕨗����O�֎c�����߁B
    Call HC_Log.Info("ThisWorkbook", "--- STARTUP: HC_CSV Add-in loading initiated ---")
    
    ' 3. ���s�����̋L�^�B
    ' # �y�ړI�z��蔭������ Excel �o�[�W������ OS ���Ɉˑ�����s���؂蕪����ؐՂƂ��邽�߁B
    Call ThisWorkbook.LogEnvironmentInfo
    Call HC_StartupPerf.StartupPerfMark("after_log_environment_info")
    
    ' UI_SERVER�̋N�� �� ��d�N���̉\�������邽�߈ڍs�R�����g�A�E�g
    ' Dim UISERVER As String
    
    ' UISERVER = "from svc import hc_csv_mg; hc_csv_mg.ensure_qt_ui_server()"
    ' UISERVER = "from svc import hc_host; hc_host.ensure_ui_server()"
    
    ' Call HC_Log.Info("ThisWorkbook", "ui_server start:" & UISERVER)
    ' Call RunPython(UISERVER)
    
    Dim warmUp As String
    ' Python: svc / ui / bridge / register_book �� 1 ��� RunPython �Ŋ����i�v���� HC_LOG_PERF�j
    warmUp = "from svc.svc_host import excel_startup_workbook_open_full; excel_startup_workbook_open_full(" & Application.hwnd & ")"
    Call HC_Log.Diag("ThisWorkbook", "startup_full: " & warmUp)
    
    Call HC_StartupPerf.StartupPerfMark("before_runpython_startup_full")
    Call Main.MarkWorkbookOpenStartupFullStarted
    Err.Clear
    RunPython warmUp
    If Err.Number = 0 Then
        Call Main.MarkWorkbookOpenFullPythonDone
    Else
        Call HC_Log.Error("ThisWorkbook", "startup_full RunPython Err=" & CStr(Err.Number) & " (delayed InitPythonServer suppressed until Manual_Init)")
    End If
    Call HC_StartupPerf.StartupPerfMark("after_runpython_startup_full")
    
    ' 4. �C���X�g�[���ς݃A�h�C���̑����iRunPython ��E�N���u���b�N�Z�k�j�B
    ' # �y�ړI�z�N�����_�ł� Excel ���i�L���ȑ��A�h�C���̃��X�g�j���ڍׂ�c�����邽�߁B
    Call ThisWorkbook.LogInstalledAddins
    Call HC_StartupPerf.StartupPerfMark("after_log_installed_addins")
    
    
    ' 5. �񓯊��������V�[�P���X�̗\��B
    ' # �y�ړI�zExcel �{�̂� Ready ��ԁi�A�C�h���j�ɂȂ�̂�҂��Ă��珉���������s���邽�߁B
    
    ' �y�d�v�zPython �풓�v���Z�X�̋N���͏�L startup_full �̂݁B�x�����̓C�x���g�Z���T�[�{�K�v���̂� InitPythonServer�B
    Call HC_StartupPerf.StartupPerfMark("before_wait_and_init")
    Call Main.WaitAndInit
    Call HC_StartupPerf.StartupPerfMark("after_wait_and_init_returned")
    
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' �R�[���o�b�N��: Workbook_BeforeClose
' �쐬��: 2026-02-01
' ���Ŕԍ�����ї���:
'   1.0.7 (2026-06-13) [�I��] ScheduleDeferredShutdown �p�~�AShutdownExcelUiVbaOnly �ցB
'   1.0.6 (2026-06-13) [�I��] ScheduleDeferredShutdown ��ɑ����v���[�u��ǉ��B
'   1.0.5 (2026-06-13) [�I��] ���� ShutdownExcelUiCleanup ��p�~�� ScheduleDeferredShutdown �ցB
'   1.0.4 (2026-06-07) [�I��] �I�� RunPython �� ShutdownExcelUiCleanup 1 ��ɓ����i�d�� shutdown �폜�j�B
'   1.0.3 (2026-06-03) shutdown_all_with_force_kill �ɕύX�B
'   1.0.2 (2026-05-31) Main.ShutdownExcelUiCleanup 1.1.0�iEnableEvents�EPython �����j�ɒǐ��B�{�N���X�͌Ăяo���̂݁B
'   1.0.1 (2026-05-30) ShutdownExcelUiCleanup �Ăяo����ǉ��B
' �v���V�[�W���̓���T�v: �A�h�C���I���iExcel �I���j���O�ɔ��΂��A�S���\�[�X�����S�ɉ������B
' ����: Cancel (Boolean) - �I�����L�����Z������ꍇ�� True ���Z�b�g�B
' �߂�l: �Ȃ�
' ---------------------------------------------------------------------------------------------------------------------
Private Sub Workbook_BeforeClose(ByRef Cancel As Boolean)
    On Error Resume Next

    If Cancel Then
        Exit Sub
    End If

    ' 1. VBA �� UI �̂ݕЕt���iPython �I���� core.excel_lifecycle_monitor �� Excel �I�����m�Ŏ��{�j�B
    ' # �y�ړI�z�I���m�F�L�����Z������ Python �풓���ێ����A�J�[�\���EWaitForm ���𕜌����邽�߁B
    Call HC_Log.Info("ThisWorkbook", "--- SHUTDOWN: VBA UI cleanup only (Python lifecycle monitor) ---")

    Call Main.ShutdownExcelUiVbaOnly

    On Error GoTo 0
End Sub

' ==============================================================================
' �����Ǘ��E��̓��\�b�h
' ==============================================================================

' ---------------------------------------------------------------------------------------------------------------------
' �v���V�[�W����: LogEnvironmentInfo
' �쐬��: 2026-02-01
' ���Ŕԍ�����ї���: 1.0.0 (2026-02-01) �V�K�쐬�B
' �v���V�[�W���̓���T�v: ���s���̏ڍׁi�o�[�W�����AOS�A�A�h�C�������p�X�j�����O�ɋL�^����B
' ---------------------------------------------------------------------------------------------------------------------
Public Sub LogEnvironmentInfo()
    Dim envMsg As String                                    ' ����񕶎���
    
    On Error Resume Next
    
    ' 1. ��{���̍\�z�B
    ' # �y�ړI�zExcel �̃o�[�W��������ѓ��� OS �̏��𕨗����肷�邽�߁B
    envMsg = "Excel Ver: " & Application.VERSION & " (" & Application.OperatingSystem & ")"
    
    ' 2. ���O�֏o�́B
    Call HC_Log.Info("ThisWorkbook", envMsg)
    
    ' 3. ���A�h�C���̕����p�X���L�^�B
    ' # �y�ړI�z�J�����Ɩ{�Ԋ��ňقȂ�p�X�̃A�h�C�������݂��Ă��Ȃ������m�F���邽�߁B
    Call HC_Log.Info("ThisWorkbook", "Add-in Path: " & ThisWorkbook.FullName)
    
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' �v���V�[�W����: LogInstalledAddins
' �쐬��: 2026-02-01
' ���Ŕԍ�����ї���: 1.0.1 (2026-04-06) �^�p���O�ihc_csv.log�j�փC���X�g�[���ς݌����� 1 �s�T�}����ǉ��B
' ���Ŕԍ�����ї���: 1.0.0 (2026-02-01) �V�K�쐬�B
' �v���V�[�W���̓���T�v: ���� Excel �ɗL��������Ă���A�h�C�������ׂĎ擾���A���O�ɋL�^����B
' ---------------------------------------------------------------------------------------------------------------------
Public Sub LogInstalledAddins()
    Dim addinObj As addin                                   ' �A�h�C������p�I�u�W�F�N�g
    Dim statusMsg As String                                 ' ���O�p������
    Dim installedCount As Long                              ' Installed=True �̌����i�^�p���O�T�}���p�j
    
    On Error Resume Next
    
    installedCount = 0
    
    ' # �y�ړI�z�N���v���Z�X�ɂ����āA�ǂ̃A�h�C�����L��������Ă��邩�𕨗��I�Ɍ��؂��邽�߁B
    Call HC_Log.Diag("ThisWorkbook", "Scanning active Excel Add-ins...")
    
    ' 1. �A�v���P�[�V�������x���̃A�h�C���R���N�V�����𑖍��B
    For Each addinObj In Application.AddIns
        ' # �y����z�A�h�C�������݃C���X�g�[���i�L���j��Ԃɂ��邩�B
        If addinObj.Installed = True Then
            installedCount = installedCount + 1
            ' # �y�ړI�z�L���ȃA�h�C���̖��̂ƕ����p�X���\�z�B
            statusMsg = "Active Add-in: [" & addinObj.Name & "] - Path: " & addinObj.Path
            
            ' # �y�ړI�z�\�z�����������O�֔��f�B
            Call HC_Log.Diag("ThisWorkbook", statusMsg)
        End If
    Next addinObj
    
    ' 2. �^�p���O�֌����T�}���i�ڍׂ� HC_LOG_DIAG ���̂� hc_csv_diag.log�j�B
    Call HC_Log.Info("ThisWorkbook", "Installed add-ins at startup: count=" & CStr(installedCount))
    
    ' 3. ���������̋L�^�i�f�f�j�B
    Call HC_Log.Diag("ThisWorkbook", "Add-in scan completed.")
    
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' �v���V�[�W����: InitEvents
' �쐬��: 2026-02-01
' �v���V�[�W���̓���T�v: �A�v���P�[�V�����Ď��Z���T�[�����̉����AExcel �̋����ߑ����J�n����B
' ---------------------------------------------------------------------------------------------------------------------
Public Sub InitEvents()
    On Error Resume Next
    
    Call HC_StartupPerf.StartupPerfMark("init_events_enter")
    
    ' 1. �����C���X�^���X�̕����j���B
    ' # �y�ړI�z��d�N���ɂ�铮��s�����h�~���邽�߁A��x Nothing �Ń��Z�b�g�����s�B
    Set mSensor = Nothing
    
    ' 2. �Z���T�[�I�u�W�F�N�g�̐V�K���̉��B
    Set mSensor = New AppEventClass
    
    ' # �y����z�Z���T�[�̐����ɐ����������B
    If Not mSensor Is Nothing Then
        ' 3. Excel �A�v���P�[�V�����{�̂Ƃ̕����ڑ��B
        ' # �y�ړI�zOS ���x���Ŕ��s����� Excel �C�x���g���Z���T�[�փo�C�p�X�����邽�߁B
        Set mSensor.App = Application
        
        ' # �y�ړI�z�����̎����𕨗����O�֋L�^�B
        Call HC_Log.Info("ThisWorkbook", "InitEvents: Event sensor successfully linked to Application.")
        Call HC_StartupPerf.StartupPerfMark("init_events_after_sensor_linked")
    Else
        ' # �y�ړI�z�v���I�Ȑ������s�����O�֋L�^�B
        Call HC_Log.Error("ThisWorkbook", "CRITICAL: Failed to instantiate AppEventClass.")
    End If
    
    ' 4. Python ���T�[�o�[�̎��ȏC���o�^�B
    ' # �y�ړI�zVBA ���̏����������ɓ������� Python �i�ߓ��Ƃ� COM �ʐM���m�����邽�߁B
    Call HC_StartupPerf.StartupPerfMark("init_events_before_init_python_server")
    If Main.IsWorkbookOpenFullPythonDone() Or Main.IsWorkbookOpenStartupFullStarted() Then
        Call HC_Log.Info("ThisWorkbook", "InitEvents: InitPythonServer skipped (startup_full done or in progress).")
        Call HC_StartupPerf.StartupPerfMark("init_events_skip_init_python_server")
    Else
        Call Main.InitPythonServer
        Call HC_StartupPerf.StartupPerfMark("init_events_after_init_python_server")
    End If
    
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' �v���V�[�W����: Manual_Init
' �쐬��: 2026-02-01
' �v���V�[�W���̓���T�v: ���[�U�[�̑���܂��̓G���[�����̂��߂ɁA�Ď��V�X�e�����ċN������B
' ---------------------------------------------------------------------------------------------------------------------
Public Sub Manual_Init()
    ' 1. �蓮����̎������L�^�B
    ' # �y�ړI�z��̉�͎��Ƀ��[�U�[����ɂ��ċN�������������Ƃ𔻕ʉ\�ɂ��邽�߁B
    Call HC_Log.Info("ThisWorkbook", "Manual_Init: Re-initialization triggered by user/system.")
    
    ' 1b. �x���������� RunPython�iexcel_startup_after_excel_idle�j��K���ʂ��B
    Call Main.ResetWorkbookOpenFullPythonDone
    
    ' 2. �����������\�b�h�ւ̕��������[�B
    Call ThisWorkbook.InitEvents
    
    ' 3. ���[�U�[�ւ̊����ʒm���s�B
    ' # �y�d�v�zComm �N���X�͎g�p�֎~�̂��߁A�W���� MsgBox ���g�p�B
    Call MsgBox("�A�h�C���̊Ď��Z���T�[�𐳏�ɍč\�z���܂����B", vbInformation, "CSV Tool ������")
End Sub
