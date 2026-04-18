Option Explicit

Public Sub ShortcutMainSheet_CtrlShift0()
    On Error Resume Next
    If Not ActiveWorkbook Is ThisWorkbook Then Exit Sub
    メインシートA1を選択
    On Error GoTo 0
End Sub

' Ctrl+Shift+テンキー 1 … 図形「アニメ付き_タスク抽出を実行」と同じ（段階1）
Public Sub ShortcutStage1_OnKey()
    On Error Resume Next
    If Not ActiveWorkbook Is ThisWorkbook Then Exit Sub
    アニメ付き_タスク抽出を実行
    On Error GoTo 0
End Sub

' Ctrl+Shift+テンキー 2 … 図形「アニメ付き_計画生成を実行」と同じ（段階2）
Public Sub ShortcutStage2_OnKey()
    On Error Resume Next
    If Not ActiveWorkbook Is ThisWorkbook Then Exit Sub
    アニメ付き_計画生成を実行
    On Error GoTo 0
End Sub

' Ctrl+Shift+テンキー 0 … 図形「アニメ付き_段階1と段階2を連続実行」と同じ
Public Sub ShortcutStage1Then2_OnKey()
    On Error Resume Next
    If Not ActiveWorkbook Is ThisWorkbook Then Exit Sub
    アニメ付き_段階1と段階2を連続実行
    On Error GoTo 0
End Sub

Public Sub ShortcutMainSheet_OnKeyRegister()
    On Error Resume Next
    Application.OnKey Key:=SHORTCUT_MAIN_SHEET_ONKEY, Procedure:="ShortcutMainSheet_CtrlShift0"
    Application.OnKey Key:=SHORTCUT_STAGE1_THEN_STAGE2_ONKEY, Procedure:="ShortcutStage1Then2_OnKey"
    Application.OnKey Key:=SHORTCUT_STAGE1_ONKEY, Procedure:="ShortcutStage1_OnKey"
    Application.OnKey Key:=SHORTCUT_STAGE2_ONKEY, Procedure:="ShortcutStage2_OnKey"
    On Error GoTo 0
End Sub

Public Sub ShortcutMainSheet_OnKeyUnregister()
    On Error Resume Next
    Application.OnKey Key:=SHORTCUT_MAIN_SHEET_ONKEY
    Application.OnKey Key:=SHORTCUT_STAGE1_THEN_STAGE2_ONKEY
    Application.OnKey Key:=SHORTCUT_STAGE1_ONKEY
    Application.OnKey Key:=SHORTCUT_STAGE2_ONKEY
    On Error GoTo 0
End Sub

Sub Auto_Open()
    ShortcutMainSheet_OnKeyRegister
End Sub



