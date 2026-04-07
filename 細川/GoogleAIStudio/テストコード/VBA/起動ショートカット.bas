Private Sub MoveToMainSheetA1()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetMainWorksheet()
    If ws Is Nothing Then Exit Sub
    ws.Activate
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

Public Sub ShortcutMainSheet_CtrlShift0()
    On Error Resume Next
    If Not ActiveWorkbook Is ThisWorkbook Then Exit Sub
    MoveToMainSheetA1
    On Error GoTo 0
End Sub

Public Sub ShortcutMainSheet_OnKeyRegister()
    On Error Resume Next
    Application.OnKey Key:=SHORTCUT_MAIN_SHEET_ONKEY, Procedure:="ShortcutMainSheet_CtrlShift0"
    On Error GoTo 0
End Sub

Public Sub ShortcutMainSheet_OnKeyUnregister()
    On Error Resume Next
    Application.OnKey Key:=SHORTCUT_MAIN_SHEET_ONKEY
    On Error GoTo 0
End Sub

Sub Auto_Open()
    ShortcutMainSheet_OnKeyRegister
End Sub
