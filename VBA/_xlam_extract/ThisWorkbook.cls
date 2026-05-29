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
' モジュール名: ThisWorkbook (クラスモジュール)
' 作成日: 2025-11-28
' 更新日: 2026-02-01
' 改版番号および履歴:
'   1.9.5 (2026-02-01) [最終仕様確定] 起動時の環境情報（Ver/Path）および全アドイン走査ログ機能を完全統合。
'                          VBA開発規定（非省略、1行1命令、ヘッダ強化、Comm禁止、HC_Log一本化）を全域適用。
'   1.8.0 (2026-02-01) 起動時に全アドインの読込状態を走査しログ出力する機能を実装。
'   1.7.0 (2026-02-01) 原点復帰に伴うアドインロード・終了シーケンスのログ記録強化。
' プロシージャの動作概要: アドインの物理的な生存期間（ライフサイクル）を監視・制御する。
'                          Excel 起動時の環境情報収集、監視センサーの起動、および終了時のクリーンアップを主導する。
' 注意事項: Comm クラスの使用は厳禁。メッセージ通知は標準 MsgBox、ログは HC_Log を使用すること。
'           マルチステートメント（:）を禁止し、全ての論理に「# 【目的】」コメントを付帯させる。
'           変数名は短縮形（mSensor 等）を維持すること。
' ---------------------------------------------------------------------------------------------------------------------

' --- モジュールレベル変数定義 ---
' # 【目的】Excel アプリケーション全体のイベントを捕捉し続けるため。
Private mSensor As AppEventClass                            ' アプリケーション監視センサーのインスタンス

' ==============================================================================
' ブックレベル・イベントハンドラ (アドイン・エントリポイント)
' ==============================================================================

' ---------------------------------------------------------------------------------------------------------------------
' コールバック名: Workbook_Open
' 作成日: 2026-02-01
' 改版番号および履歴: 1.3.0 (2026-02-01) 環境情報の収集とアドイン走査の呼び出しを追加。
' プロシージャの動作概要: アドイン読み込み時に自動執行され、環境情報の収集と初期化予約を行う。
' 引数: なし
' 戻り値: なし
' 呼出し例: Excel によるアドインロード時に自動実行。
' ヘルパープロシージャの親子関係: (子) HC_Log.InsertBlankLine, HC_Log.Info, ThisWorkbook.LogEnvironmentInfo,
'                                  ThisWorkbook.LogInstalledAddins, Main.WaitAndInit
' 注意事項: Python（svc/ui/bridge/register_book）は Workbook_Open 内の 1 回の RunPython で完了。
'           イベントセンサー接続のみ Main.WaitAndInit → RunInitEvents で遅延する。
' ---------------------------------------------------------------------------------------------------------------------
Private Sub Workbook_Open()
    On Error Resume Next
    
    ' 0. 起動計測（HC_LOG_PERF=1 のときのみ hc_csv_perf.log）。
    ' # 【目的】Workbook_Open から遅延初期化完了までの経過を同一基準時刻で追跡するため。
    Call HC_StartupPerf.StartupPerfReset
    
    ' 1. 物理ログの初期化。
    ' # 【目的】ログファイル上で新しい Excel セッションが開始されたことを視覚的に区別するため。
    Call HC_Log.InsertBlankLine
    Call HC_StartupPerf.StartupPerfMark("after_insert_blank_line")
    
    ' 2. 起動開始の記録。
    ' # 【目的】アドインがシステムに読み込まれた正確な時刻と事実を物理ログへ残すため。
    Call HC_Log.Info("ThisWorkbook", "--- STARTUP: HC_CSV Add-in loading initiated ---")
    
    ' 3. 実行環境情報の記録。
    ' # 【目的】問題発生時に Excel バージョンや OS 環境に依存する不具合を切り分ける証跡とするため。
    Call ThisWorkbook.LogEnvironmentInfo
    Call HC_StartupPerf.StartupPerfMark("after_log_environment_info")
    
    ' 4. インストール済みアドインの走査。
    ' # 【目的】起動時点での Excel 環境（有効な他アドインのリスト）を詳細を把握するため。
    Call ThisWorkbook.LogInstalledAddins
    Call HC_StartupPerf.StartupPerfMark("after_log_installed_addins")
    
    ' UI_SERVERの起動 → 二重起動の可能性があるため移行コメントアウト
    ' Dim UISERVER As String
    
    ' UISERVER = "from svc import hc_csv_mg; hc_csv_mg.ensure_qt_ui_server()"
    ' UISERVER = "from svc import hc_host; hc_host.ensure_ui_server()"
    
    ' Call HC_Log.Info("ThisWorkbook", "ui_server start:" & UISERVER)
    ' Call RunPython(UISERVER)
    
    Dim warmUp As String
    ' Python: svc / ui / bridge / register_book を 1 回の RunPython で完了（計測は HC_LOG_PERF）
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
    
    ' 5. 非同期初期化シーケンスの予約。
    ' # 【目的】Excel 本体が Ready 状態（アイドル）になるのを待ってから初期化を執行するため。
    
    ' 【重要】Python 常駐プロセスの起動は上記 startup_full のみ。遅延側はイベントセンサー＋必要時のみ InitPythonServer。
    Call HC_StartupPerf.StartupPerfMark("before_wait_and_init")
    Call Main.WaitAndInit
    Call HC_StartupPerf.StartupPerfMark("after_wait_and_init_returned")
    
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' コールバック名: Workbook_BeforeClose
' 作成日: 2026-02-01
' プロシージャの動作概要: アドイン終了（Excel 終了）直前に発火し、全リソースを安全に解放する。
' 引数: Cancel (Boolean) - 終了をキャンセルする場合に True をセット。
' 戻り値: なし
' ---------------------------------------------------------------------------------------------------------------------
Private Sub Workbook_BeforeClose(ByRef Cancel As Boolean)
    On Error Resume Next
    
    ' 1. 終了シーケンスの開始記録。
    ' # 【目的】クリーンアップが正常に開始されたことを物理ログへ記録するため。
    Call HC_Log.Info("ThisWorkbook", "--- SHUTDOWN: Add-in closing sequence started ---")
    
    ' 2. 監視センサーの物理解放。
    ' # 【目的】メモリリークを防止し、アプリケーションレベルのフックを安全に解除するため。
    Set mSensor = Nothing
    
    ' 3．Python/Qt側へ終了要求（共通基盤）
    Dim PYSHUT As String
    PYSHUT = "from svc import svc_host; svc_host.request_shutdown_all()"
    
    Call HC_Log.Info("ThisWorkbook", "ui_server shutdown:" & PYSHUT)
    Call RunPython(PYSHUT)
    
    ' 4. Python 側資源の物理抹消。
    ' # 【目的】Python プロセス側に残存するブック参照やキャッシュをクリアするため。
    Call Main.TerminatePython
    
    ' 5. 正常終了の記録。
    ' # 【目的】すべての終了工程が欠損なく完遂された事実を証明するため。
    Call HC_Log.Info("ThisWorkbook", "SHUTDOWN: Cleanup completed successfully.")
    
    On Error GoTo 0
End Sub

' ==============================================================================
' 物理管理・解析メソッド
' ==============================================================================

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: LogEnvironmentInfo
' 作成日: 2026-02-01
' 改版番号および履歴: 1.0.0 (2026-02-01) 新規作成。
' プロシージャの動作概要: 実行環境の詳細（バージョン、OS、アドイン物理パス）をログに記録する。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub LogEnvironmentInfo()
    Dim envMsg As String                                    ' 環境情報文字列
    
    On Error Resume Next
    
    ' 1. 基本情報の構築。
    ' # 【目的】Excel のバージョンおよび動作 OS の情報を物理特定するため。
    envMsg = "Excel Ver: " & Application.VERSION & " (" & Application.OperatingSystem & ")"
    
    ' 2. ログへ出力。
    Call HC_Log.Info("ThisWorkbook", envMsg)
    
    ' 3. 自アドインの物理パスを記録。
    ' # 【目的】開発環境と本番環境で異なるパスのアドインが混在していないかを確認するため。
    Call HC_Log.Info("ThisWorkbook", "Add-in Path: " & ThisWorkbook.FullName)
    
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: LogInstalledAddins
' 作成日: 2026-02-01
' 改版番号および履歴: 1.0.1 (2026-04-06) 運用ログ（hc_csv.log）へインストール済み件数の 1 行サマリを追加。
' 改版番号および履歴: 1.0.0 (2026-02-01) 新規作成。
' プロシージャの動作概要: 現在 Excel に有効化されているアドインをすべて取得し、ログに記録する。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub LogInstalledAddins()
    Dim addinObj As addin                                   ' アドイン操作用オブジェクト
    Dim statusMsg As String                                 ' ログ用文字列
    Dim installedCount As Long                              ' Installed=True の件数（運用ログサマリ用）
    
    On Error Resume Next
    
    installedCount = 0
    
    ' # 【目的】起動プロセスにおいて、どのアドインが有効化されているかを物理的に検証するため。
    Call HC_Log.Diag("ThisWorkbook", "Scanning active Excel Add-ins...")
    
    ' 1. アプリケーションレベルのアドインコレクションを走査。
    For Each addinObj In Application.AddIns
        ' # 【判定】アドインが現在インストール（有効）状態にあるか。
        If addinObj.Installed = True Then
            installedCount = installedCount + 1
            ' # 【目的】有効なアドインの名称と物理パスを構築。
            statusMsg = "Active Add-in: [" & addinObj.Name & "] - Path: " & addinObj.Path
            
            ' # 【目的】構築した情報をログへ反映。
            Call HC_Log.Diag("ThisWorkbook", statusMsg)
        End If
    Next addinObj
    
    ' 2. 運用ログへ件数サマリ（詳細は HC_LOG_DIAG 時のみ hc_csv_diag.log）。
    Call HC_Log.Info("ThisWorkbook", "Installed add-ins at startup: count=" & CStr(installedCount))
    
    ' 3. 走査完了の記録（診断）。
    Call HC_Log.Diag("ThisWorkbook", "Add-in scan completed.")
    
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: InitEvents
' 作成日: 2026-02-01
' プロシージャの動作概要: アプリケーション監視センサーを実体化し、Excel の挙動捕捉を開始する。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub InitEvents()
    On Error Resume Next
    
    Call HC_StartupPerf.StartupPerfMark("init_events_enter")
    
    ' 1. 既存インスタンスの物理破棄。
    ' # 【目的】二重起動による動作不安定を防止するため、一度 Nothing でリセットを執行。
    Set mSensor = Nothing
    
    ' 2. センサーオブジェクトの新規実体化。
    Set mSensor = New AppEventClass
    
    ' # 【判定】センサーの生成に成功したか。
    If Not mSensor Is Nothing Then
        ' 3. Excel アプリケーション本体との物理接続。
        ' # 【目的】OS レベルで発行される Excel イベントをセンサーへバイパスさせるため。
        Set mSensor.App = Application
        
        ' # 【目的】成功の事実を物理ログへ記録。
        Call HC_Log.Info("ThisWorkbook", "InitEvents: Event sensor successfully linked to Application.")
        Call HC_StartupPerf.StartupPerfMark("init_events_after_sensor_linked")
    Else
        ' # 【目的】致命的な生成失敗をログへ記録。
        Call HC_Log.Error("ThisWorkbook", "CRITICAL: Failed to instantiate AppEventClass.")
    End If
    
    ' 4. Python 側サーバーの自己修復登録。
    ' # 【目的】VBA 側の初期化完了に同期して Python 司令塔との COM 通信を確立するため。
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
' プロシージャ名: Manual_Init
' 作成日: 2026-02-01
' プロシージャの動作概要: ユーザーの操作またはエラー復旧のために、監視システムを再起動する。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub Manual_Init()
    ' 1. 手動介入の事実を記録。
    ' # 【目的】後の解析時にユーザー操作による再起動があったことを判別可能にするため。
    Call HC_Log.Info("ThisWorkbook", "Manual_Init: Re-initialization triggered by user/system.")
    
    ' 1b. 遅延初期化で RunPython（excel_startup_after_excel_idle）を必ず通す。
    Call Main.ResetWorkbookOpenFullPythonDone
    
    ' 2. 実初期化メソッドへの物理リレー。
    Call ThisWorkbook.InitEvents
    
    ' 3. ユーザーへの完了通知執行。
    ' # 【重要】Comm クラスは使用禁止のため、標準の MsgBox を使用。
    Call MsgBox("アドインの監視センサーを正常に再構築しました。", vbInformation, "CSV Tool 初期化")
End Sub
