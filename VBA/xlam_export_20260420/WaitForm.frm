VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} WaitForm 
   Caption         =   "起動処理中"
   ClientHeight    =   1200
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   4560
   OleObjectBlob   =   "WaitForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "WaitForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

