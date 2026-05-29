Attribute VB_Name = "WaitForm"
Attribute VB_Base = "0{235E8C8E-B1B2-48E2-859F-8E0FE890C7F6}{DC467E66-4209-4B5C-8D36-0695CC71B077}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Option Explicit

Private Sub UserForm_Activate()
    ' # 【目的】モードレスでも Excel メインウィンドウの中央に表示する。
    On Error Resume Next
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (Application.Width - Me.Width) / 2
    Me.Top = Application.Top + (Application.Height - Me.Height) / 2
    On Error GoTo 0
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' # 【目的】× ボタンで閉じられないようにする（仕様）。
    Cancel = True
End Sub

