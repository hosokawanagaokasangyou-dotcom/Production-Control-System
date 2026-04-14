Attribute VB_Name = "modGanttDateNav"
Option Explicit

' 結果_設備ガント／結果_設備ガント_実績明細 に、Forms 2.0 のコンボボックス（OLE）を載せ、
' 列 A の【yyyy/mm/dd】日付ブロック先頭行へジャンプする。

Public mGanttDateNavFillBusy As Boolean

Private Const GANTT_DATE_NAV_OLE_NAME As String = "GanttDateNavCombo"
Private Const GANTT_DATE_BANNER_L As String = "【"
Private Const GANTT_DATE_BANNER_R As String = "】"

Private mNavComboHost As clsGanttDateNavCombo

Public Function 結果_設備ガント日付ナビ_対象シートか(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    If StrComp(ws.Name, SHEET_RESULT_EQUIP_GANTT, vbBinaryCompare) = 0 Then
        結果_設備ガント日付ナビ_対象シートか = True
    ElseIf StrComp(ws.Name, SHEET_RESULT_EQUIP_GANTT_ACTUAL_DETAIL, vbBinaryCompare) = 0 Then
        結果_設備ガント日付ナビ_対象シートか = True
    End If
End Function

' MSForms の ComboBox / ListBox 向け（2 列目に先頭行番号を隠し保持）
Public Sub GanttDateNav_FillMsFormsList(ByVal lst As Object, ByVal ws As Worksheet)
    Dim lastR As Long
    Dim r As Long
    Dim c As Range
    Dim v As Variant
    Dim s As String
    Dim topR As Long
    
    mGanttDateNavFillBusy = True
    On Error GoTo CleanBusy
    
    On Error Resume Next
    lst.Clear
    lst.ColumnCount = 2
    lst.ColumnWidths = "110 pt;0 pt"
    On Error GoTo CleanBusy
    
    If ws Is Nothing Then GoTo CleanBusy
    
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 4 Then GoTo CleanBusy
    
    For r = 4 To lastR
        Set c = ws.Cells(r, 1)
        If c.MergeCells Then
            topR = c.MergeArea.Row
            If topR <> r Then GoTo NextR
        Else
            topR = r
        End If
        
        v = c.Value
        If IsError(v) Then GoTo NextR
        s = Trim$(CStr(v))
        If Len(s) < 3 Then GoTo NextR
        If Left$(s, 1) <> GANTT_DATE_BANNER_L Then GoTo NextR
        If Right$(s, 1) <> GANTT_DATE_BANNER_R Then GoTo NextR
        
        lst.AddItem
        lst.List(lst.ListCount - 1, 0) = Mid$(s, 2, Len(s) - 2)
        lst.List(lst.ListCount - 1, 1) = CStr(topR)
NextR:
    Next r
    
CleanBusy:
    mGanttDateNavFillBusy = False
End Sub

Private Sub GanttDateNav_WireHostIfNeeded(ByVal ws As Worksheet, ByVal cb As MSForms.ComboBox)
    If mNavComboHost Is Nothing Then
        Set mNavComboHost = New clsGanttDateNavCombo
    End If
    Set mNavComboHost.HostSheet = ws
    Set mNavComboHost.cbo = cb
End Sub

Private Function GanttDateNav_GetOrCreateOleCombo(ByVal ws As Worksheet) As OLEObject
    Dim ole As OLEObject
    Dim needNew As Boolean
    needNew = True
    
    On Error Resume Next
    Set ole = ws.OLEObjects(GANTT_DATE_NAV_OLE_NAME)
    If Err.Number = 0 And Not ole Is Nothing Then
        If TypeOf ole.Object Is MSForms.ComboBox Then
            needNew = False
        Else
            ole.Delete
            Set ole = Nothing
        End If
    End If
    Err.Clear
    On Error GoTo 0
    
    If needNew Then
        Dim a As Range
        Dim leftPt As Double, topPt As Double
        Set a = ws.Range("A2")
        leftPt = a.Left + a.Width * 6#
        topPt = a.Top + 1#
        Set ole = ws.OLEObjects.Add( _
            ClassType:="Forms.ComboBox.1", _
            Left:=leftPt, _
            Top:=topPt, _
            Width:=140, _
            Height:=22)
        ole.Name = GANTT_DATE_NAV_OLE_NAME
        ole.Placement = xlFreeFloating
        ole.PrintObject = False
        On Error Resume Next
        ole.Object.Font.Size = 10
        Err.Clear
        On Error GoTo 0
    End If
    
    Set GanttDateNav_GetOrCreateOleCombo = ole
End Function

' アクティブシートが設備ガント系のとき、コンボをシート上に用意し日付一覧を入れる。
Public Sub 結果_設備ガント系_日付ジャンプコンボを確保()
    Dim ws As Worksheet
    Dim ole As OLEObject
    Dim cb As MSForms.ComboBox
    
    Set ws = ActiveSheet
    If Not 結果_設備ガント日付ナビ_対象シートか(ws) Then
        MsgBox "「" & SHEET_RESULT_EQUIP_GANTT & "」または「" & SHEET_RESULT_EQUIP_GANTT_ACTUAL_DETAIL _
            & "」を表示してから実行してください。", vbExclamation, "日付へ移動"
        Exit Sub
    End If
    
    On Error GoTo FailOle
    Set ole = GanttDateNav_GetOrCreateOleCombo(ws)
    Set cb = ole.Object
    GanttDateNav_WireHostIfNeeded ws, cb
    GanttDateNav_FillMsFormsList cb, ws
    Exit Sub
    
FailOle:
    MsgBox "シート上へのコンボボックス配置に失敗しました: " & Err.Description & vbCrLf _
        & "ブックの保護・参照設定（Microsoft Forms 2.0 Object Library）を確認してください。", vbCritical, "日付へ移動"
End Sub

' 既にコンボがある場合のみリストを再構築（段階2取込直後などから呼ぶ場合用）
Public Sub 結果_設備ガント系_日付コンボを再充填()
    Dim ws As Worksheet
    Dim ole As OLEObject
    
    Set ws = ActiveSheet
    If Not 結果_設備ガント日付ナビ_対象シートか(ws) Then Exit Sub
    
    On Error Resume Next
    Set ole = ws.OLEObjects(GANTT_DATE_NAV_OLE_NAME)
    On Error GoTo 0
    If ole Is Nothing Then Exit Sub
    
    GanttDateNav_WireHostIfNeeded ws, ole.Object
    GanttDateNav_FillMsFormsList ole.Object, ws
End Sub
