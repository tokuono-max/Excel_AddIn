Attribute VB_Name = "HC_Debug"
Option Explicit

' ---------------------------------------------------------------------------------------------------------------------
' モジュール名: HC_Debug (標準モジュール)
' 作成日: 2026-02-02
' 改版番号および履歴:
'   1.0.0 (2026-02-02) 新規作成。シートプロパティの物理状態を確認するためのデバッグ専用ツール。
' プロシージャの動作概要: アクティブシートに設定されている CustomProperties をすべて列挙し、イミディエイトへ出力する。
' 注意事項: 実行前にイミディエイトウィンドウ（Ctrl+G）を開いておくこと。
'           開発規定（1行1命令、意図コメント、詳細ヘッダ）を完全適用。
' ---------------------------------------------------------------------------------------------------------------------

' ---------------------------------------------------------------------------------------------------------------------
' プロシージャ名: ShowSheetProperties
' 作成日: 2026-02-02
' 改版番号および履歴: 1.0.0 (2026-02-02) 初版。
' プロシージャの動作概要: アクティブシートの全カスタムプロパティを名称と値のペアでイミディエイトウィンドウへ出力する。
' 引数: なし
' 戻り値: なし
' 呼出し例: イミディエイトウィンドウで「Call HC_Debug.ShowSheetProperties」と入力して実行。
' ヘルパープロシージャの親子関係: (子) HC_Log.Info
' 注意事項: プロパティが存在しない場合はその旨を表示する。
' ---------------------------------------------------------------------------------------------------------------------
Public Sub ShowSheetProperties()
    Dim p As CustomProperty                                 ' プロパティループ用
    Dim i As Long                                           ' カウント用
    Dim sN As String                                        ' シート名保持
    
    ' 1. 対象の生存確認。
    ' # 【判定】アクティブなシートが存在するか。
    If ActiveSheet Is Nothing Then
        ' # 【目的】無効参照によるエラーを防止するため。
        Debug.Print "Error: ActiveSheet is nothing."
        Exit Sub
    End If
    
    sN = ActiveSheet.Name
    Debug.Print "=== Property Inspection: [" & sN & "] ==="
    
    ' 2. プロパティの列挙と出力。
    i = 0
    ' # 【目的】シートに刻印されたすべての情報を物理的に走査するため。
    For Each p In ActiveSheet.CustomProperties
        i = i + 1
        ' # 【補足】名称と値をセットで出力し、キーの typo や値の欠落を視認可能にする。
        Debug.Print i & ": Key=[" & p.Name & "] / Value=[" & p.Value & "]"
    Next p
    
    ' 3. 結果のサマリー表示。
    ' # 【判定】プロパティが 1 つも見つからなかったか。
    If i = 0 Then
        ' # 【目的】分割処理時にプロパティが消失している可能性をユーザーに示唆するため。
        Debug.Print "Result: No CustomProperties found in this sheet."
    Else
        Debug.Print "Result: " & i & " properties found."
    End If
    
    Debug.Print "========================================"
    
    ' 4. 証跡の記録。
    Call HC_Log.Diag("HC_Debug", "Property inspection completed for [" & sN & "]. Count: " & i)
End Sub

Sub test_01()
    MsgBox Environ("XLWINGS_PYTHON")
End Sub
