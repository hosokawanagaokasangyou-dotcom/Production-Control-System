<<<<<<< HEAD
Private Sub メインシートA1へ移動()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("メイン_")
    If ws Is Nothing Then Set ws = ThisWorkbook.Worksheets(1)
    If ws Is Nothing Then Exit Sub
    ws.Activate
    ws.Range("A1").Select
    On Error GoTo 0
End Sub
=======
Option Explicit
>>>>>>> main4

Public Sub ショートカット_メイン_CtrlShift0()
    On Error Resume Next
    If Not ActiveWorkbook Is ThisWorkbook Then Exit Sub
<<<<<<< HEAD
    メインシートA1へ移動
=======
    メインシートA1を選択
>>>>>>> main4
    On Error GoTo 0
End Sub

Public Sub ショートカット_メイン_OnKey登録()
    On Error Resume Next
    Application.OnKey Key:=SHORTCUT_MAIN_SHEET_ONKEY, Procedure:="ショートカット_メイン_CtrlShift0"
    On Error GoTo 0
End Sub

Public Sub ショートカット_メイン_OnKey解除()
    On Error Resume Next
    Application.OnKey Key:=SHORTCUT_MAIN_SHEET_ONKEY
    On Error GoTo 0
End Sub

Sub Auto_Open()
    ショートカット_メイン_OnKey登録
End Sub



