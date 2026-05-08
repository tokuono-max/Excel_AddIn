Attribute VB_Name = "HC_Log"
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' モジュール名: HC_Log (標準モジュール)
' 作成日: 2026-01-30
' 更新日: 2026-04-06
' 文字コード: 本モジュールのソースは Shift-JIS（CP932）で保存すること。
' ログファイル: Python との共有のため UTF-8（バイナリ書込）を維持。hc_csv.log=運用、hc_csv_diag.log=診断、hc_csv_perf.log=計測。
' 改版番号および履歴:
'   1.5.0 (2026-04-06) Diag/Perf 追加。環境変数 HC_LOG_DIAG（別名 HC_DEBUG）/ HC_LOG_PERF。docs/environment_variables.md 参照。
'   1.4.1 (2026-01-31) VBA 開発規定（非省略、1行1命令、ヘッダ強化）を完全適用。
'   1.4.0 (2026-01-31) [視認性向上] 物理的なセッション区切り線（Separator）挿入機能を実装。
'   1.3.0 (2026-01-31) [設計変更] CodeName 依存を排除。識別子として Base64 GUID を受け入れるよう調整。
'   1.2.0 (2026-01-31) LIFO（最新行先頭挿入）および UTF-8 エンコード変換の実装。
' プロシージャの動作概要: 運用ログは Info/Error。詳細トレースは Diag、性能計測は Perf（いずれも環境変数 ON 時のみファイルへ）。
' 注意事項: Win32 API 使用。ファイル操作は排他ロック。Python 側 core_log と同一フォルダ %TEMP%\csv_tool\。
' ---------------------------------------------------------------------------------------------------------------------

' --- Win32 API 構造体定義 ---
' 【目的】ミリ秒精度のシステム時刻を取得・保持するため。
Private Type SYSTEMTIME
    wYear As Integer                                        ' 年
    wMonth As Integer                                       ' 月
    wDayOfWeek As Integer                                   ' 曜日
    wDay As Integer                                         ' 日
    wHour As Integer                                        ' 時
    wMinute As Integer                                      ' 分
    wSecond As Integer                                      ' 秒
    wMilliseconds As Integer                                ' ミリ秒
End Type

' --- Win32 API 外部関数宣言 ---
' 【目的】高精度な時刻取得および文字コード変換を OS レベルで執行するため。
#If VBA7 Then
    ' 64bit/32bit 共通（VBA7環境）
    Private Declare PtrSafe Sub GetLocalTime Lib "kernel32" (ByRef lpSystemTime As SYSTEMTIME)
    Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" ( _
        ByVal CodePage As Long, _
        ByVal dwFlags As Long, _
        ByVal lpWideCharStr As LongPtr, _
        ByVal cchWideChar As Long, _
        ByVal lpMultiByteStr As LongPtr, _
        ByVal cbMultiByte As Long, _
        ByVal lpDefaultChar As LongPtr, _
        ByVal lpUsedDefaultChar As LongPtr) As Long
#Else
    ' レガシー環境用
    Private Declare Sub GetLocalTime Lib "kernel32" (ByRef lpSystemTime As SYSTEMTIME)
    Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
        ByVal CodePage As Long, _
        ByVal dwFlags As Long, _
        ByVal lpWideCharStr As Long, _
        ByVal cchWideChar As Long, _
        ByVal lpMultiByteStr As Long, _
        ByVal cbMultiByte As Long, _
        ByVal lpDefaultChar As Long, _
        ByVal lpUsedDefaultChar As Long) As Long
#End If

' --- 共通定数定義 ---
' 【目的】ログ制御に関する物理パラメータを一貫して管理するため。
Private Const CP_UTF8 As Long = 65001                       ' UTF-8 コードページ識別子
Private Const LOG_FILE_NAME As String = "hc_csv.log"         ' 運用ログ
Private Const LOG_DIAG_NAME As String = "hc_csv_diag.log"   ' 診断ログ（HC_LOG_DIAG または HC_DEBUG）
Private Const LOG_PERF_NAME As String = "hc_csv_perf.log"   ' 計測ログ（HC_LOG_PERF）
Private Const MAX_LOG_SIZE As Long = 1048576                ' ログファイル最大保持サイズ (1MB)
Private Const RETRY_COUNT As Integer = 5                    ' ファイルロック競合時の最大再試行回数

' ==============================================================================
' 公開インターフェース・セクション
' ==============================================================================

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: InsertBlankLine
' 作成日: 2026-01-31
' 改版番号および履歴: 1.0.0 (2026-01-31) 新規作成。
' プロシージャの動作概要: ログファイルの冒頭に視覚的なセッション区切り線を物理挿入し、新旧セッションを分離する。
' 引数: なし
' 戻り値: なし
' 呼出し例: Call HC_Log.InsertBlankLine()
' ヘルパープロシージャの親子関係: (子) WriteRawData
' 注意事項: ThisWorkbook.Workbook_Open の最優先事項として呼び出すこと。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub InsertBlankLine()
    Dim separatorText As String                             ' 構築する区切り文字列
    
    ' 【目的】エディタ上でセッションの切り替わりを一目で判別可能にするため、物理的な境界線を構築。
    separatorText = "========================================================================" & vbCrLf & _
                    "Excel Session Start at " & Format(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf & _
                    "========================================================================" & vbCrLf & vbCrLf
    
    ' 命令分離: 詳細書き込みエンジンへリレー。
    Call HC_Log.WriteRawData(separatorText)
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: Info
' 作成日: 2026-01-30
' 改版番号および履歴: 1.1.0 (2026-01-31) 引数名称の非省略化。
' プロシージャの動作概要: INFO レベルのログメッセージを物理ログファイルへ出力する。
' 引数:
'   senderName (String) - 出力元のモジュール名またはクラス名
'   messageText (String) - ログメッセージ本文
' 戻り値: なし
' 呼出し例: Call HC_Log.Info("Main", "Application Started")
' ヘルパープロシージャの親子関係: (子) WriteReverseLog
' ---------------------------------------------------------------------------------------------------------------------
Public Sub Info(ByVal senderName As String, ByVal messageText As String)
    ' 命令分離: レベル「INFO」を指定して物理構築へリレー。
    Call HC_Log.WriteReverseLog("INFO", senderName, messageText)
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: Error
' 作成日: 2026-01-30
' 改版番号および履歴: 1.1.0 (2026-01-31) 引数名称の非省略化。
' プロシージャの動作概要: ERROR レベルのログメッセージを物理ログファイルへ出力する。
' 引数:
'   senderName (String) - 出力元のモジュール名
'   messageText (String) - エラー内容の詳細
' 戻り値: なし
' 呼出し例: Call HC_Log.Error("ExcelUtil", "GUID generation failed")
' ヘルパープロシージャの親子関係: (子) WriteReverseLog
' ---------------------------------------------------------------------------------------------------------------------
Public Sub Error(ByVal senderName As String, ByVal messageText As String)
    ' 命令分離: レベル「ERROR」を指定して物理構築へリレー。
    Call HC_Log.WriteReverseLog("ERROR", senderName, messageText)
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: Diag
' 概要: 診断ログ（hc_csv_diag.log）。HC_LOG_DIAG=1 または HC_DEBUG=1 のときのみ出力（Python core_env と整合）。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub Diag(ByVal senderName As String, ByVal messageText As String)
    If Not DiagLogEnabled() Then
        Exit Sub
    End If
    Call WriteReverseLogToFile("DIAG", senderName, messageText, LOG_DIAG_NAME)
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: Perf
' 概要: 計測ログ（hc_csv_perf.log）。HC_LOG_PERF=1 のときのみ出力。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub Perf(ByVal senderName As String, ByVal messageText As String)
    If Not PerfLogEnabled() Then
        Exit Sub
    End If
    Call WriteReverseLogToFile("PERF", senderName, messageText, LOG_PERF_NAME)
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' 公開: 診断ログが有効か（他モジュールからの早期判定用）
' ---------------------------------------------------------------------------------------------------------------------
Public Function DiagLogEnabled() As Boolean
    DiagLogEnabled = DiagEnvOn()
End Function

' ---------------------------------------------------------------------------------------------------------------------
' 公開: 計測ログが有効か
' ---------------------------------------------------------------------------------------------------------------------
Public Function PerfLogEnabled() As Boolean
    PerfLogEnabled = PerfEnvOn()
End Function

' ==============================================================================
' 物理書き込みエンジン・セクション
' ==============================================================================

Private Function DiagEnvOn() As Boolean
    If EnvTruthy(Environ("HC_LOG_DIAG")) Then
        DiagEnvOn = True
        Exit Function
    End If
    If EnvTruthy(Environ("HC_DEBUG")) Then
        DiagEnvOn = True
        Exit Function
    End If
End Function

Private Function PerfEnvOn() As Boolean
    PerfEnvOn = EnvTruthy(Environ("HC_LOG_PERF"))
End Function

Private Function EnvTruthy(ByVal rawVal As String) As Boolean
    Dim t As String
    t = LCase$(Trim$(rawVal))
    EnvTruthy = (t = "1" Or t = "true" Or t = "yes" Or t = "on" Or t = "y")
End Function

Private Function EnsureLogDir() As String
    Dim logDir As String
    logDir = Environ("TEMP") & "\csv_tool"
    If Len(Dir(logDir, vbDirectory)) = 0 Then
        MkDir logDir
    End If
    EnsureLogDir = logDir
End Function

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: WriteReverseLog
' 作成日: 2026-01-30
' 改版番号および履歴: 1.2.0 (2026-01-31) ISO 8601 準拠のタイムスタンプ構築ロジックを強化。
' プロシージャの動作概要: タイムスタンプ、レベル、送信元を含むログ行を構築し、物理書き込みへリレーする。
' 引数:
'   levelString (String) - ログレベル（INFO/ERROR等）
'   senderName (String) - 出力元名称
'   messageText (String) - ログ本文
' 戻り値: なし
' ヘルパープロシージャの親子関係: (親) Info, Error (子) WriteRawData
' ---------------------------------------------------------------------------------------------------------------------
Private Sub WriteReverseLog(ByVal levelString As String, ByVal senderName As String, ByVal messageText As String)
    Dim timeSystem As SYSTEMTIME                            ' Win32 高精度時刻構造体
    Dim timeString As String                                ' フォーマット済み時刻文字列
    Dim lineContent As String                               ' 最終的なログ行文字列
    
    On Error Resume Next
    
    ' 1. 高精度時刻の物理取得。
    Call GetLocalTime(timeSystem)
    
    ' 【目的】ミリ秒を含む標準的なログ形式を物理構築。
    timeString = Format(DateSerial(timeSystem.wYear, timeSystem.wMonth, timeSystem.wDay), "yyyy-mm-dd") & " " & _
                 Format(TimeSerial(timeSystem.wHour, timeSystem.wMinute, timeSystem.wSecond), "hh:nn:ss") & "." & _
                 Format(timeSystem.wMilliseconds, "000")
            
    ' 2. ログ行の完全構築。
    ' 判定コメント: 改行コードを付与して 1行を完結させる。
    lineContent = "[" & timeString & "] [" & levelString & "] [VBA." & senderName & "] " & messageText & vbCrLf
    
    ' 3. 詳細書き込み処理へのリレー。
    Call WriteRawDataFile(EnsureLogDir() & "\" & LOG_FILE_NAME, lineContent)
    
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: WriteReverseLogToFile
' 概要: 指定ファイル名（csv_tool 配下）へ LIFO 書込。Diag/Perf 用。
' ---------------------------------------------------------------------------------------------------------------------
Private Sub WriteReverseLogToFile(ByVal levelString As String, ByVal senderName As String, ByVal messageText As String, ByVal baseFileName As String)
    Dim timeSystem As SYSTEMTIME
    Dim timeString As String
    Dim lineContent As String
    
    On Error Resume Next
    
    Call GetLocalTime(timeSystem)
    timeString = Format(DateSerial(timeSystem.wYear, timeSystem.wMonth, timeSystem.wDay), "yyyy-mm-dd") & " " & _
                 Format(TimeSerial(timeSystem.wHour, timeSystem.wMinute, timeSystem.wSecond), "hh:nn:ss") & "." & _
                 Format(timeSystem.wMilliseconds, "000")
    lineContent = "[" & timeString & "] [" & levelString & "] [VBA." & senderName & "] " & messageText & vbCrLf
    Call WriteRawDataFile(EnsureLogDir() & "\" & baseFileName, lineContent)
    
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: WriteRawData
' 作成日: 2026-01-31
' 改版番号および履歴: 1.5.0 (2026-04-06) WriteRawDataFile へ委譲。
' ---------------------------------------------------------------------------------------------------------------------
Private Sub WriteRawData(ByVal writeContent As String)
    Call WriteRawDataFile(EnsureLogDir() & "\" & LOG_FILE_NAME, writeContent)
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: WriteRawDataFile
' 作成日: 2026-04-06
' プロシージャの動作概要: 絶対パス指定で UTF-8 LIFO 書込（運用・診断・計測で共用）。
' ---------------------------------------------------------------------------------------------------------------------
Private Sub WriteRawDataFile(ByVal logFilePath As String, ByVal writeContent As String)
    Dim fileHandle As Integer                               ' ファイル番号
    Dim newBinary() As Byte                                 ' 変換後の UTF-8 バイト配列
    Dim oldBinary() As Byte                                 ' 吸引した既存データのバイト配列
    Dim finalBinary() As Byte                               ' 結合後の最終データ配列
    Dim retryCounter As Integer                             ' 排他ロック待機用のカウンタ
    
    On Error Resume Next
    
    ' 1. 文字コード変換 (Unicode -> UTF-8)。Python core_log と同一ファイルを共有するため UTF-8 維持。
    newBinary = StringToUTF8(writeContent)
    
    ' 2. 排他制御を伴う書き込みシーケンス。
    For retryCounter = 1 To RETRY_COUNT
        ' 命令分離: 使用可能なファイル番号を取得。
        fileHandle = FreeFile
        
        ' 【重要】Binary モードかつ Lock Read Write で開き、Python 側との同時アクセスを物理ガード。
        Open logFilePath For Binary Access Read Write Lock Read Write As #fileHandle
        
        ' 判定コメント: ファイルオープン（ロック確保）に成功したか。
        If Err.Number = 0 Then
            ' 既存データの物理吸引。
            If LOF(fileHandle) > 0 Then
                ' 配列サイズを物理確保。
                ReDim oldBinary(LOF(fileHandle) - 1)
                ' ファイル全体をメモリへ一括吸引。
                Get #fileHandle, , oldBinary
            End If
            
            ' 判定コメント: LIFO (New Data + Old Data) 形式で結合。
            If (Not oldBinary) = -1 Then
                ' 命令分離: 新規ファイル。
                finalBinary = newBinary
            Else
                ' 命令分離: 配列結合の執行。
                finalBinary = CombineBytes(newBinary, oldBinary)
            End If
            
            ' 判定コメント: 最大ファイルサイズ (1MB) を超過していないか。
            If UBound(finalBinary) + 1 > MAX_LOG_SIZE Then
                ' 【目的】古いログ（配列の末尾側）を切り捨ててファイルサイズを物理維持。
                ReDim Preserve finalBinary(MAX_LOG_SIZE - 1)
            End If
            
            ' 【重要】ファイルの先頭 (位置: 1) から最終データを物理的に上書き保存。
            Put #fileHandle, 1, finalBinary
            
            ' 命令分離: ファイルをクローズしロックを解放。
            Close #fileHandle
            
            ' 命令分離: 成功したためリトライループを脱出。
            Exit For
        Else
            ' 命令分離: エラー情報をクリアし再試行を準備。
            Err.Clear
            ' 命令分離: OS への制御返却。
            DoEvents
            ' 動作補足: 50ms のインターバルを置いて再試行。
            Call Application.Wait(Now + TimeSerial(0, 0, 1) / 20)
        End If
    Next retryCounter
    
    On Error GoTo 0
End Sub

' ==============================================================================
' 物理変換・配列演算セクション（Private）
' ==============================================================================

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: StringToUTF8
' 作成日: 2026-01-31
' 概要: VBA内部の Unicode 文字列を物理的な UTF-8 バイト配列へ変換する。
' ---------------------------------------------------------------------------------------------------------------------
Private Function StringToUTF8(ByVal sourceText As String) As Byte()
    Dim resultBytes() As Byte                               ' 変換後のバイト配列
    Dim bufferSize As Long                                  ' 変換に必要なバッファサイズ
    
    ' 判定。
    If Len(sourceText) = 0 Then
        Exit Function
    End If
    
    ' 1. 必要なバッファサイズの物理取得。
    bufferSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sourceText), Len(sourceText), 0, 0, 0, 0)
    
    ' 判定。
    If bufferSize > 0 Then
        ' 命令分離: 出力先配列の確保。
        ReDim resultBytes(bufferSize - 1)
        ' 2. Win32 API による変換の執行。
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(sourceText), Len(sourceText), VarPtr(resultBytes(0)), bufferSize, 0, 0)
    End If
    
    ' 戻り値。
    StringToUTF8 = resultBytes
End Function

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: CombineBytes
' 作成日: 2026-01-31
' 概要: 2つのバイト配列を物理的に結合した新しい配列を生成する。
' ---------------------------------------------------------------------------------------------------------------------
Private Function CombineBytes(ByRef firstArray() As Byte, ByRef secondArray() As Byte) As Byte()
    Dim combinedBuffer() As Byte                            ' 結合後のバッファ
    Dim lengthFirst As Long                                 ' 第1配列の長さ
    Dim lengthSecond As Long                                ' 第2配列の長さ
    Dim loopIndex As Long                                   ' ループカウンタ
    
    ' 命令分離: 各配列の物理長を特定。
    lengthFirst = UBound(firstArray) + 1
    lengthSecond = UBound(secondArray) + 1
    
    ' 命令分離: 結合後のサイズを確保。
    ReDim combinedBuffer(lengthFirst + lengthSecond - 1)
    
    ' 【目的】LIFO を実現するため、第1配列（新規ログ）を先頭に配置。
    For loopIndex = 0 To lengthFirst - 1
        combinedBuffer(loopIndex) = firstArray(loopIndex)
    Next loopIndex
    
    ' 【目的】第2配列（既存ログ）を後方に配置。
    For loopIndex = 0 To lengthSecond - 1
        combinedBuffer(lengthFirst + loopIndex) = secondArray(loopIndex)
    Next loopIndex
    
    ' 戻り値。
    CombineBytes = combinedBuffer
End Function

