Attribute VB_Name = "HC_RibbonPerf"
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' モジュール名: HC_RibbonPerf (標準モジュール)
' 作成日: 2026-04-06
' 文字コード: Shift-JIS（CP932）で保存すること。
' プロシージャの動作概要: リボンクリックから xlwings RunPython 戻りまでを HC_LOG_PERF=1 のとき hc_csv_perf.log に記録する。
'                          基準時刻は RibbonCallback_hc_main 先頭。RibbonPerfEnd は複数回呼んでもよい（2 重終了は無視）。
' ---------------------------------------------------------------------------------------------------------------------

#If VBA7 Then
    Private Declare PtrSafe Function GetTickCount64 Lib "kernel32" () As LongLong
    Private m_clickT0 As LongLong
#End If

Private m_chainActive As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: RibbonPerfBegin
' プロシージャの動作概要: リボンコールバック先頭で呼ぶ。クリック基準時刻を記録し phase=click_enter を出力。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub RibbonPerfBegin()
    m_chainActive = False
    #If VBA7 Then
        m_clickT0 = GetTickCount64()
    #End If
    If Not HC_Log.PerfLogEnabled() Then
        Exit Sub
    End If
    m_chainActive = True
    Call HC_Log.Perf("RibbonInvoke", "phase=click_enter elapsed_since_click_ms=0")
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: RibbonPerfMark
' プロシージャの動作概要: m_chainActive のときのみ、クリックからの経過 ms で 1 行出力。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub RibbonPerfMark(ByVal phaseKey As String)
    If Not m_chainActive Then
        Exit Sub
    End If
    If Not HC_Log.PerfLogEnabled() Then
        Exit Sub
    End If
    #If VBA7 Then
        Dim dt As LongLong
        dt = GetTickCount64() - m_clickT0
        Call HC_Log.Perf("RibbonInvoke", "phase=" & phaseKey & " elapsed_since_click_ms=" & CStr(dt))
    #End If
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: RibbonPerfEnd
' プロシージャの動作概要: 1 クリック分の計測チェーンを終了。m_chainActive を False にする。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub RibbonPerfEnd()
    m_chainActive = False
End Sub
