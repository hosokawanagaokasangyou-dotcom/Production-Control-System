VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_LargeMessage
   BackColor       =   &H8000000F&
   BorderStyle     =   1  'fmBorderStyleSingle
   Caption         =   "お知らせ"
   ClientHeight    =   4860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8280
   ShowModal       =   -1  'True
   StartUpPosition =   1  'CenterOwner
   Begin MSForms.Label lblStripe
      BackColor       =   &H00C0C0C0&
      Height          =   4860
      Left            =   0
      Top             =   0
      Width           =   120
      BorderStyle     =   0
      VariousPropertyBits=   8388627
      Size            =   "635;6350"
      SpecialEffect   =   0
      FontName        =   "Meiryo UI"
      FontHeight      =   200
      FontWeight      =   400
      FontCharSet     =   128
   End
   Begin MSForms.TextBox txtBody
      Height          =   3960
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7800
      VariousPropertyBits=   746604571
      Size            =   "5292;6350"
      SpecialEffect   =   0
      FontName        =   "Meiryo UI"
      FontHeight      =   280
      FontWeight      =   400
      FontCharSet     =   128
      BorderStyle     =   1
      ScrollBars      =   2
      PasswordChar    =   0
      MatchEntry      =   0
      ShowDropButtonWhen=   0
      DropButtonStyle =   0
      MultiLine       =   -1  'True
      AutoSize        =   0   'False
      WordWrap        =   -1  'True
      Locked          =   -1  'True
      EnterKeyBehavior=   -1  'True
   End
   Begin MSForms.CommandButton cmdOK
      Caption         =   "OK"
      Height          =   420
      Left            =   3360
      TabIndex        =   1
      Top             =   4320
      Width           =   1560
      Default         =   -1  'True
      VariousPropertyBits=   8388739
      Size            =   "2640;635"
      FontName        =   "Meiryo UI"
      FontHeight      =   240
      FontWeight      =   400
      FontCharSet     =   128
      Accelerator     =   79
   End
   Begin MSForms.CommandButton cmdCancel
      Caption         =   "キャンセル"
      Height          =   420
      Left            =   5160
      TabIndex        =   2
      Top             =   4320
      Width           =   1680
      Cancel          =   -1  'True
      VariousPropertyBits=   8388739
      Size            =   "2640;635"
      FontName        =   "Meiryo UI"
      FontHeight      =   240
      FontWeight      =   400
      FontCharSet     =   128
   End
   Begin MSForms.CommandButton cmdYes
      Caption         =   "はい"
      Height          =   420
      Left            =   2760
      TabIndex        =   3
      Top             =   4320
      Width           =   1560
      VariousPropertyBits=   8388739
      Size            =   "2640;635"
      FontName        =   "Meiryo UI"
      FontHeight      =   240
      FontWeight      =   400
      FontCharSet     =   128
   End
   Begin MSForms.CommandButton cmdNo
      Caption         =   "いいえ"
      Height          =   420
      Left            =   4560
      TabIndex        =   4
      Top             =   4320
      Width           =   1560
      VariousPropertyBits=   8388739
      Size            =   "2640;635"
      FontName        =   "Meiryo UI"
      FontHeight      =   240
      FontWeight      =   400
      FontCharSet     =   128
   End
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
            cmdOK.Left = (Me.ClientWidth - cmdOK.Width) / 2
        Case vbOKCancel
            cmdOK.Visible = True
            cmdCancel.Visible = True
            cmdOK.Left = Me.ClientWidth / 2 - cmdOK.Width - 120
            cmdCancel.Left = Me.ClientWidth / 2 + 120
        Case vbYesNo
            cmdYes.Visible = True
            cmdNo.Visible = True
            cmdYes.Left = Me.ClientWidth / 2 - cmdYes.Width - 120
            cmdNo.Left = Me.ClientWidth / 2 + 120
        Case Else
            cmdOK.Visible = True
            cmdOK.Left = (Me.ClientWidth - cmdOK.Width) / 2
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
