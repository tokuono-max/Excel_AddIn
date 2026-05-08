Attribute VB_Name = "ExcelUtil"
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' モジュール名: ExcelUtil (標準モジュール)
' 作成日: 2025-11-28
' 更新日: 2026-02-02
' 改版番号および履歴:
'   4.6.2 (2026-02-02) [仕様変更] ステータス保持用キーを KEY_INFO (HC_SHEET_INFO) から KEY_STAT (HC_STATUS_INFO) へ換装。
'                          ステータス表示と完了通知の物理分離アーキテクチャをサポート。VBA開発規定を全域適用。
'   4.6.1 (2026-02-02) 所属ブック照合による自己修復時、履歴情報を削除せず維持するよう変更。
'   4.6.0 (2026-02-02) 他ブックからのシート移動/コピーを検知し GUID を強制再生成する自己修復機能を搭載。
' プロシージャの動作概要: ワークシート個別の属性管理（識別子発行・所属チェック・実行情報保存）を主導する。
'                          ステータスバー表示用の情報を専用プロパティ (HC_STATUS_INFO) で物理管理する。
' 注意事項: メッセージ通知は標準 MsgBox、ログは HC_Log を使用すること（Comm使用禁止）。
'           マルチステートメント（:）を禁止し、全ての論理に「# 【目的】」コメントを付帯させる。
' ---------------------------------------------------------------------------------------------------------------------

' --- Windows API 外部関数宣言 ---
' # 【目的】OS 標準の不変識別子（GUID）を生成するため。
#If VBA7 Then
    Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (ByRef pGuid As GUID) As Long
#Else
    Private Declare Function CoCreateGuid Lib "ole32.dll" (ByRef pGuid As GUID) As Long
#End If

' --- 構造体定義 ---
' # 【目的】API 経由で取得した GUID バイナリを保持するため。
Private Type GUID
    Data1 As Long                                           ' GUID 第1ブロック
    Data2 As Integer                                        ' GUID 第2ブロック
    Data3 As Integer                                        ' GUID 第3ブロック
    Data4(7) As Byte                                        ' GUID 第4ブロック
End Type

' --- 共通定数定義 ---
' # 【目的】 Worksheet.CustomProperties で使用する物理キー名称を一元管理するため。
Private Const KEY_GUID As String = "HC_GUID_B64"            ' 不変識別子用キー
Private Const KEY_BOOK As String = "HC_BOOK_NAME"           ' 所属ブック名用キー（自己修復判定用）
' # 【重要】ステータスバー専用のキー名称を詳細定義。
Private Const KEY_STAT As String = "HC_STATUS_INFO"         ' ステータスバー表示情報用キー

' ==============================================================================
' 公開インターフェース（識別子管理・検索・自己修復）
' ==============================================================================

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: GetSheetIdSafe
' 作成日: 2026-02-02
' プロシージャの動作概要: シートから識別子（GUID）を吸引する。所属ブックが不一致の場合は強制的に再発行を行う。
' 引数: Sh (Worksheet) - 対象シートオブジェクト
' 戻り値: String - 22文字の Base64 GUID
' ---------------------------------------------------------------------------------------------------------------------
Public Function GetSheetIdSafe(ByVal Sh As Worksheet) As String
    Dim sId As String                                       ' 吸引した識別子保持用
    Dim sBk As String                                       ' 吸引した所属ブック名保持用
    Dim curBk As String                                     ' 現在の物理ブック名保持用
    
    On Error Resume Next
    
    ' # 【判定】引数が有効なオブジェクトであるか。
    If Sh Is Nothing Then
        Exit Function
    End If
    
    ' 1. 現在の環境情報の物理取得。
    ' # 【目的】現在そのシートが所属しているブックの名称を特定するため。
    curBk = Sh.Parent.Name
    
    ' 2. 既存の属性情報を吸引。
    ' # 【目的】シートに刻印されている ID と所属情報を確認するため。
    sId = ExcelUtil.GetSheetProp(Sh, KEY_GUID)
    sBk = ExcelUtil.GetSheetProp(Sh, KEY_BOOK)
    
    ' 3. 身元確認と再発行の執行判定。
    ' # 【判定】ID が不在、または記録された所属ブック名が現物と異なる（別名保存・移動等）か。
    If sId = "" Or sBk <> curBk Then
        
        ' # 【補足】環境変化（移動/コピー/別名保存）を検知した場合の処理。
        If sId <> "" And sBk <> curBk Then
            ' # 【目的】コンテキストの変化があった事実を物理ログへ記録するため。
            Call HC_Log.Diag("ExcelUtil", "Context change detected. Updating GUID for [" & Sh.Name & "]")
        End If
        
        ' # 【目的】このブック専用の新しい一意な識別子を生成するため。
        sId = ExcelUtil.GenerateBase64GUID()
        
        ' # 【目的】新しい識別子と、現在の所属ブック名を物理的に刻印（永続化）するため。
        Call ExcelUtil.WriteProp(Sh, KEY_GUID, sId)
        Call ExcelUtil.WriteProp(Sh, KEY_BOOK, curBk)
        
        ' # 【重要】実行履歴 (HC_STATUS_INFO) は削除せず維持。
        
        ' # 【目的】更新完了のログを記録。
        Call HC_Log.Diag("ExcelUtil", "GUID and Ownership updated. History preserved for [" & Sh.Name & "]")
    End If
    
    ' # 【戻り値】
    GetSheetIdSafe = sId
    
    On Error GoTo 0
End Function

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: FindSheetById
' 作成日: 2026-02-01
' プロシージャの動作概要: 指定された GUID を保持するシートを ActiveWorkbook 内から走査し特定する。
' ---------------------------------------------------------------------------------------------------------------------
Public Function FindSheetById(ByVal sId As String) As Worksheet
    Dim ws As Worksheet                                     ' 走査用変数
    Dim sFound As String                                    ' 吸引された ID 保持用
    
    On Error Resume Next
    
    ' # 【判定】
    If sId = "" Then
        Exit Function
    End If
    
    ' 1. 全シートの物理巡回。
    For Each ws In ActiveWorkbook.Worksheets
        ' # 【目的】各シートの GUID 保持プロパティを吸引。
        sFound = ExcelUtil.GetSheetProp(ws, KEY_GUID)
        
        ' # 【判定】吸引した ID がターゲットと完全に一致するか。
        If sFound = sId Then
            ' # 【目的】特定したシートオブジェクトを戻り値にセット。
            Set FindSheetById = ws
            Exit For
        End If
    Next ws
    
    On Error GoTo 0
End Function

' ==============================================================================
' 公開インターフェース（実行情報管理・プロパティ操作）
' ==============================================================================

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: GetSheetProp
' 概要: Worksheet.CustomProperties から値を安全に吸引する。
' ---------------------------------------------------------------------------------------------------------------------
Public Function GetSheetProp(ByVal Sh As Worksheet, ByVal Key As String) As String
    Dim propObj As CustomProperty                           ' 走査用オブジェクト
    
    On Error Resume Next
    
    ' # 【判定】
    If Sh Is Nothing Then
        Exit Function
    End If
    
    ' 1. 物理走査の執行。
    For Each propObj In Sh.CustomProperties
        ' # 【判定】
        If propObj.Name = Key Then
            GetSheetProp = CStr(propObj.Value)
            Exit Function
        End If
    Next propObj
    
    On Error GoTo 0
End Function

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: WriteProp
' 概要: Worksheet.CustomProperties へ値を書き込み保存する。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub WriteProp(ByVal Sh As Worksheet, ByVal Key As String, ByVal valText As String)
    On Error Resume Next
    
    ' # 【判定】
    If Sh Is Nothing Then
        Exit Sub
    End If
    
    ' 1. 既存キーの物理抹消。
    Call ExcelUtil.DeleteSheetProp(Sh, Key)
    
    ' 2. 新規追加の執行。
    Call Sh.CustomProperties.Add(Name:=Key, Value:=valText)
    
    ' 3. 失敗時のログ出力。
    If Err.Number <> 0 Then
        Call HC_Log.Error("ExcelUtil", "WriteProp FAILED [" & Key & "]: " & Err.Description)
        Err.Clear
    End If
    
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: DeleteSheetProp
' 概要: 指定されたシートのカスタムプロパティを物理的に削除する。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub DeleteSheetProp(ByVal Sh As Worksheet, ByVal Key As String)
    Dim propObj As CustomProperty                           ' 走査用オブジェクト
    
    On Error Resume Next
    
    ' # 【判定】
    If Sh Is Nothing Then
        Exit Sub
    End If
    
    ' 1. 物理走査と削除の執行。
    For Each propObj In Sh.CustomProperties
        If propObj.Name = Key Then
            Call propObj.Delete
        End If
    Next propObj
    
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: SetSheetInfo
' 作成日: 2026-02-02
' プロシージャの動作概要: 実行結果メッセージをステータス専用プロパティへ物理保存する。
' 引数: Sh (Worksheet) - 対象シート、valText (String) - 保存内容
' ---------------------------------------------------------------------------------------------------------------------
Public Sub SetSheetInfo(ByVal Sh As Worksheet, ByVal valText As String)
    ' # 【目的】ステータスバー専用のキー (HC_STATUS_INFO) へ情報を永続刻印するため。
    Call ExcelUtil.WriteProp(Sh, KEY_STAT, valText)
End Sub

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: GetSheetInfo
' 作成日: 2026-02-02
' プロシージャの動作概要: シートに保存されているステータス専用メッセージを吸引する。
' 戻り値: String - 保存されている実行結果メッセージ
' ---------------------------------------------------------------------------------------------------------------------
Public Function GetSheetInfo(ByVal Sh As Worksheet) As String
    ' # 【目的】ステータスバー復旧用の情報を専用プロパティ (HC_STATUS_INFO) から取得するため。
    GetSheetInfo = ExcelUtil.GetSheetProp(Sh, KEY_STAT)
End Function

' ==============================================================================
' 物理識別子（GUID）生成エンジン（内部用）
' ==============================================================================

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: GenerateBase64GUID
' 概要: OS 標準の GUID を生成し、短縮版の Base64 (22文字) へ変換して返却する。
' ---------------------------------------------------------------------------------------------------------------------
Public Function GenerateBase64GUID() As String
    Dim tGuid As GUID                                       ' 物理構造体
    Dim sGuid As String                                     ' 連結用文字列
    Dim i As Integer                                        ' ループカウンタ
    
    ' 1. OS レベルでの GUID 物理生成。
    If CoCreateGuid(tGuid) <> 0 Then
        GenerateBase64GUID = "ERR-" & Timer
        Exit Function
    End If
    
    ' 2. 各ブロックの HEX 文字列変換と連結。
    sGuid = Hex(tGuid.Data1)
    sGuid = sGuid & Hex(tGuid.Data2)
    sGuid = sGuid & Hex(tGuid.Data3)
    
    ' # 【目的】第4ブロックのバイト列を連結。
    For i = 0 To 7
        sGuid = sGuid & Hex(tGuid.Data4(i))
    Next i
    
    ' 3. 22文字の固定長整形。
    GenerateBase64GUID = Left(sGuid & String(22, "X"), 22)
End Function
