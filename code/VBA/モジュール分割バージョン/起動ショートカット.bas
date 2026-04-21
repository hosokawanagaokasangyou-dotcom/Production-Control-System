Option Explicit

' version.txt の先頭行（改行まで）をバージョン文字列として返す。UTF-8 想定（GeminiReadUtf8File 経由）
Private Function VersionTxtFirstLineNormalized(ByVal fileUtf8Content As String) As String
    Dim t As String
    Dim p As Long
    t = Trim$(fileUtf8Content)
    If Len(t) = 0 Then VersionTxtFirstLineNormalized = "": Exit Function
    p = InStr(1, t, vbCrLf, vbBinaryCompare)
    If p > 0 Then t = Trim$(Left$(t, p - 1))
    p = InStr(1, t, vbLf, vbBinaryCompare)
    If p > 0 Then t = Trim$(Left$(t, p - 1))
    p = InStr(1, t, vbCr, vbBinaryCompare)
    If p > 0 Then t = Trim$(Left$(t, p - 1))
    VersionTxtFirstLineNormalized = Trim$(t)
End Function

' マクロブック直下の version.txt と共有 UNC の version.txt の 1 行目が異なるとき警告（同一 Excel セッションで 1 回のみ）
Public Sub 配台AI_共有Version照合_起動時警告()
    Static done As Boolean
    Dim wbFolder As String
    Dim localPath As String
    Dim localVer As String
    Dim shareVer As String
    Dim msg As String
    
    If done Then Exit Sub
    
    wbFolder = Trim$(ThisWorkbook.Path)
    If Len(wbFolder) = 0 Then Exit Sub
    
    localPath = wbFolder & "\" & VERSION_TXT_FILE_NAME
    If Len(Dir(localPath)) = 0 Then Exit Sub
    
    localVer = VersionTxtFirstLineNormalized(GeminiReadUtf8File(localPath))
    If Len(localVer) = 0 Then
        done = True
        Exit Sub
    End If
    
    shareVer = VersionTxtFirstLineNormalized(GeminiReadUtf8File(VERSION_TXT_SHARED_REFERENCE_PATH))
    If Len(shareVer) = 0 Then
        done = True
        Exit Sub
    End If
    
    If StrComp(localVer, shareVer, vbBinaryCompare) = 0 Then
        done = True
        Exit Sub
    End If
    
    done = True
    msg = "配台AIシステムのバージョンが共有フォルダの正と一致しません。" & vbCrLf & vbCrLf & _
          "このブックのフォルダ内 version.txt: " & localVer & vbCrLf & _
          "共有の version.txt: " & shareVer & vbCrLf & vbCrLf & _
          "共有の一式を取得し、マクロブックと同じフォルダの version.txt を更新してください。"
    MsgBox msg, vbExclamation + vbOKOnly, "バージョン不一致"
End Sub

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
    On Error Resume Next
    配台AI_共有Version照合_起動時警告
    On Error GoTo 0
    ShortcutMainSheet_OnKeyRegister
End Sub



