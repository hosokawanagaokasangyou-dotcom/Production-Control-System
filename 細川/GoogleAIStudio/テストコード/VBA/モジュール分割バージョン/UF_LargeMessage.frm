VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_LargeMessage 
   Caption         =   "お知らせ"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "UF_LargeMessage.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UF_LargeMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public DialogResult As VbMsgBoxResult

Private Sub SetupStripeColor(ByVal buttons As VbMsgBoxStyle)
    Dim ic As Long
    ic = buttons And &HF0&
    Select Case ic
        Case vbCritical
            lblStripe.BackColor = RGB(200, 55, 45)
        Case vbExclamation
            lblStripe.BackColor = RGB(255, 185, 0)
        Case vbInformation, vbQuestion
            lblStripe.BackColor = RGB(0, 115, 200)
        Case Else
            lblStripe.BackColor = RGB(170, 170, 170)
    End Select
End Sub

Private Sub LayoutButtons(ByVal buttons As VbMsgBoxStyle)
    ' MSForms.UserForm に ClientWidth は無い。枠付き全体の幅は Me.Width（InsideWidth でも可）。
    Dim innerW As Single
    innerW = Me.Width
    Dim grp As Long
    grp = buttons And &H7&
    cmdOK.Visible = False
    cmdCancel.Visible = False
    cmdYes.Visible = False
    cmdNo.Visible = False
    cmdOK.Caption = "OK"
    cmdCancel.Caption = "キャンセル"
    cmdYes.Caption = "はい"
    cmdNo.Caption = "いいえ"
    Select Case grp
        Case 0
            cmdOK.Visible = True
            cmdOK.Left = (innerW - cmdOK.Width) / 2
        Case vbOKCancel
            cmdOK.Visible = True
            cmdCancel.Visible = True
            cmdOK.Left = innerW / 2 - cmdOK.Width - 120
            cmdCancel.Left = innerW / 2 + 120
        Case vbYesNo
            cmdYes.Visible = True
            cmdNo.Visible = True
            cmdYes.Left = innerW / 2 - cmdYes.Width - 120
            cmdNo.Left = innerW / 2 + 120
        Case Else
            cmdOK.Visible = True
            cmdOK.Left = (innerW - cmdOK.Width) / 2
    End Select
End Sub

Public Sub ApplySetup(ByVal prompt As String, ByVal buttons As VbMsgBoxStyle, ByVal title As String)
    txtBody.Text = prompt
    If Len(title) > 0 Then
        Me.Caption = title
    Else
        Me.Caption = "お知らせ"
    End If
    SetupStripeColor buttons
    LayoutButtons buttons
    DialogResult = vbOK
End Sub

Private Sub cmdOK_Click()
    DialogResult = vbOK
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    DialogResult = vbCancel
    Me.Hide
End Sub

Private Sub cmdYes_Click()
    DialogResult = vbYes
    Me.Hide
End Sub

Private Sub cmdNo_Click()
    DialogResult = vbNo
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode <> vbFormControlMenu Then Exit Sub
    If cmdCancel.Visible Then
        DialogResult = vbCancel
    ElseIf cmdNo.Visible Then
        DialogResult = vbNo
    Else
        DialogResult = vbOK
    End If
End Sub
