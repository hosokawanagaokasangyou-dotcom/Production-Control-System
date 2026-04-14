VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGanttDateNavCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' シート上の Forms 2.0 コンボボックス（OLE）用。WithEvents で選択変更を拾う。

Public WithEvents cbo As MSForms.ComboBox
Public HostSheet As Worksheet

Private Sub Class_Terminate()
    Set cbo = Nothing
    Set HostSheet = Nothing
End Sub

Private Sub cbo_Change()
    On Error GoTo Quiet
    If modGanttDateNav.mGanttDateNavFillBusy Then Exit Sub
    If cbo.ListIndex < 0 Then Exit Sub
    If HostSheet Is Nothing Then Exit Sub
    Dim tr As Long
    tr = CLng(cbo.List(cbo.ListIndex, 1))
    If tr < 1 Then Exit Sub
    HostSheet.Activate
    Application.GoTo HostSheet.Cells(tr, 1), True
Quiet:
End Sub
