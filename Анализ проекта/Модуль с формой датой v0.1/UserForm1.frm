VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   1500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3450
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    ' Проверка правильности ввода дат
    If IsDate(TextBox1.Value) And IsDate(TextBox2.Value) Then
        UserFormCancelled = False
        Me.Hide
    Else
        MsgBox "Пожалуйста, введите корректные даты.", vbExclamation
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Если форма закрыта через крестик
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        UserFormCancelled = True
        Me.Hide
    End If
End Sub

