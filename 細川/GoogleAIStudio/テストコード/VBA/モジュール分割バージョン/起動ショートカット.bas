<<<<<<< HEAD
=======
Attribute VB_Name = "起動ショートカット"
>>>>>>> hosokawa/main2
Option Explicit

Public Sub ShortcutMainSheet_CtrlShift0()
    On Error Resume Next
    If Not ActiveWorkbook Is ThisWorkbook Then Exit Sub
    メインシートA1を選択
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

<<<<<<< HEAD
Sub Auto_Open()
    ShortcutMainSheet_OnKeyRegister
End Sub



=======
>>>>>>> hosokawa/main2
