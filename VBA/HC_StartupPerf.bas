Attribute VB_Name = "HC_StartupPerf"
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' モジュール名: HC_StartupPerf (標準モジュール)
' 作成日: 2026-04-06
' 文字コード: 本モジュールのソースは Shift-JIS（CP932）で保存すること。
' プロシージャの動作概要: Excel アドイン起動シーケンスの壁時計経過（ms）を HC_LOG_PERF=1 のときのみ hc_csv_perf.log へ記録する。
'                          基準時刻 T0 は Workbook_Open 先頭で取得し、OnTime 後の RunInitEvents まで同一基準を維持する。
' 注意事項: kernel32 GetTickCount64（VBA7）。計測無効時は API 呼び出しのみ（Reset）または即 return（Mark）。
' ---------------------------------------------------------------------------------------------------------------------

#If VBA7 Then
    Private Declare PtrSafe Function GetTickCount64 Lib "kernel32" () As LongLong
    Private m_t0 As LongLong
#End If

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: StartupPerfReset
' プロシージャの動作概要: 起動計測の基準時刻 T0 を記録し、Workbook_Open 開始行を perf ログへ出力する。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub StartupPerfReset()
    #If VBA7 Then
        m_t0 = GetTickCount64()
    #End If
    If Not HC_Log.PerfLogEnabled() Then
        Exit Sub
    End If
    Call HC_Log.Perf("ExcelStartup", "phase=workbook_open_enter elapsed_since_t0_ms=0")
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: StartupPerfMark
' プロシージャの動作概要: T0 からの経過 ms と識別子 phaseKey を perf ログへ出力する。
' 引数: phaseKey (String) - フェーズ識別子（例: after_runpython_warmup）
' ---------------------------------------------------------------------------------------------------------------------
Public Sub StartupPerfMark(ByVal phaseKey As String)
    If Not HC_Log.PerfLogEnabled() Then
        Exit Sub
    End If
    #If VBA7 Then
        Dim dt As LongLong
        dt = GetTickCount64() - m_t0
        Call HC_Log.Perf("ExcelStartup", "phase=" & phaseKey & " elapsed_since_t0_ms=" & CStr(dt))
    #End If
End Sub
