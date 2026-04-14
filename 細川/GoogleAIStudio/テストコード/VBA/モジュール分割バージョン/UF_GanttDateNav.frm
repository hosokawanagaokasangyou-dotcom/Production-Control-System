VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UF_GanttDateNav 
   Caption         =   "設備ガント：日付へ移動"
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4710
   OleObjectBlob   =   "UF_GanttDateNav.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
   Begin MSForms.ListBox lstDates
      Height          =   2580
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      BorderStyle     =   1
      IntegralHeight  =   0
   End
   Begin MSForms.CommandButton cmdClose
      Caption         =   "閉じる"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   2880
      Width           =   1215
   End
End
Attribute VB_Name = "UF_GanttDateNav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public TargetWs As Worksheet

Private Sub UserForm_Activate()
    On Error Resume Next
    If TargetWs Is Nothing Then Exit Sub
    GanttDateNav_FillListBox lstDates, TargetWs
    On Error GoTo 0
End Sub

Private Sub lstDates_Change()
    On Error GoTo Quiet
    If modGanttDateNav.mGanttDateNavFillBusy Then Exit Sub
    If lstDates.ListIndex < 0 Then Exit Sub
    If TargetWs Is Nothing Then Exit Sub
    Dim tr As Long
    tr = CLng(lstDates.List(lstDates.ListIndex, 1))
    If tr < 1 Then Exit Sub
    TargetWs.Activate
    Application.GoTo TargetWs.Cells(tr, 1), True
Quiet:
End Sub

Private Sub cmdClose_Click()
    Me.Hide
End Sub
