Option Explicit

' 結果_設備ガント／結果_設備ガント_実績明細 に、Forms 2.0 のコンボボックス（OLE）を載せ、
' 列 A の【yyyy/mm/dd】日付ブロック先頭行へジャンプする。

Public mGanttDateNavFillBusy As Boolean

Private Const GANTT_DATE_NAV_OLE_NAME As String = "GanttDateNavCombo"
Private Const GANTT_DATE_NAV_UPDATE_BTN As String = "GanttDateNavUpdateBtn"
Private Const GANTT_DATE_NAV_FORM_BTN_CAPTION As String = "更新"
Private Const GANTT_DATE_BANNER_L As String = "【"
Private Const GANTT_DATE_BANNER_R As String = "】"
' ガント日付コンボのフォント（小さすぎると視認性が悪い。既存 OLE にも再適用する）
Private Const GANTT_COMBO_FONT_SIZE As Single = 12!

Private mNavComboHostPlan As clsGanttDateNavCombo
Private mNavComboHostActual As clsGanttDateNavCombo

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
    lst.ColumnWidths = "130 pt;0 pt"
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
    ' シートが2枚あるため WithEvents はシートごとに別インスタンスが必要。
    ' シート差し替え後は OLE が新規になるため、都度ホストを作り直してイベントを確実に再接続する。
    Dim h As clsGanttDateNavCombo
    If StrComp(ws.Name, SHEET_RESULT_EQUIP_GANTT_ACTUAL_DETAIL, vbBinaryCompare) = 0 Then
        Set mNavComboHostActual = Nothing
        Set mNavComboHostActual = New clsGanttDateNavCombo
        Set h = mNavComboHostActual
    ElseIf StrComp(ws.Name, SHEET_RESULT_EQUIP_GANTT, vbBinaryCompare) = 0 Then
        Set mNavComboHostPlan = Nothing
        Set mNavComboHostPlan = New clsGanttDateNavCombo
        Set h = mNavComboHostPlan
    Else
        Exit Sub
    End If
    Set h.HostSheet = ws
    Set h.cbo = cb
End Sub

Private Sub GanttDateNav_ApplyComboFont(ByVal ole As OLEObject)
    On Error Resume Next
    If ole Is Nothing Then GoTo X
    If TypeOf ole.Object Is MSForms.ComboBox Then
        Dim cb As MSForms.ComboBox
        Set cb = ole.Object
        cb.Font.Size = GANTT_COMBO_FONT_SIZE
    End If
X:
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub GanttDateNav_PositionOleCombo(ByVal ole As OLEObject, ByVal ws As Worksheet)
    Dim a1 As Range
    Dim rh As Double
    On Error Resume Next
    Set a1 = ws.Range("A1")
    ole.Left = a1.Left
    ole.Top = a1.Top
    ole.Width = ws.Range("A1:B1").Width
    rh = ws.Rows(1).RowHeight
    If rh < 12 Then rh = 15
    ' フォント 12pt 程度でも切れにくいよう最小高さを確保
    ole.Height = Application.Max(rh - 1#, 22#)
    ole.Placement = xlFreeFloating
    ole.PrintObject = False
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub GanttDateNav_DeleteUpdateButtonIfPresent(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Shapes(GANTT_DATE_NAV_UPDATE_BTN).Delete
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub GanttDateNav_PositionUpdateButton(ByVal shp As Shape, ByVal ole As OLEObject)
    On Error Resume Next
    shp.Left = ole.Left + ole.Width + 4#
    shp.Top = ole.Top
    If ole.Height > 8# Then
        shp.Height = ole.Height
    Else
        shp.Height = 20#
    End If
    If shp.Height < 16# Then shp.Height = 16#
    shp.Width = 52#
    shp.Placement = xlFreeFloating
    shp.PrintObject = False
    Err.Clear
    On Error GoTo 0
End Sub

Private Function GanttDateNav_GetOrCreateFormUpdateButton(ByVal ws As Worksheet, ByVal ole As OLEObject) As Shape
    Dim shp As Shape
    Dim needNew As Boolean
    needNew = True
    
    On Error Resume Next
    Set shp = ws.Shapes(GANTT_DATE_NAV_UPDATE_BTN)
    If Err.Number = 0 And Not shp Is Nothing Then
        needNew = False
    End If
    Err.Clear
    On Error GoTo 0
    
    If needNew Then
        GanttDateNav_DeleteUpdateButtonIfPresent ws
        On Error GoTo GanttDateNav_CreateUpdateBtnFail
        Set shp = ws.Shapes.AddFormControl(xlButtonControl, ole.Left + ole.Width + 4#, ole.Top, 52#, ole.Height)
        shp.Name = GANTT_DATE_NAV_UPDATE_BTN
        shp.OnAction = "modGanttDateNav.GanttDateNav_RunRefreshActualDetail_Click"
        On Error Resume Next
        shp.TextFrame.Characters.Text = GANTT_DATE_NAV_FORM_BTN_CAPTION
        Err.Clear
    End If
    On Error GoTo 0
    
    If Not shp Is Nothing Then
        GanttDateNav_PositionUpdateButton shp, ole
    End If
    Set GanttDateNav_GetOrCreateFormUpdateButton = shp
    Exit Function
    
GanttDateNav_CreateUpdateBtnFail:
    Set GanttDateNav_GetOrCreateFormUpdateButton = Nothing
End Function

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
        Set ole = ws.OLEObjects.Add( _
            ClassType:="Forms.ComboBox.1", _
            Left:=ws.Range("A1").Left, _
            Top:=ws.Range("A1").Top, _
            Width:=ws.Range("A1:B1").Width, _
            Height:=20)
        ole.Name = GANTT_DATE_NAV_OLE_NAME
    End If
    
    If Not ole Is Nothing Then
        GanttDateNav_PositionOleCombo ole, ws
        GanttDateNav_ApplyComboFont ole
    End If
    Set GanttDateNav_GetOrCreateOleCombo = ole
End Function

Private Sub GanttDateNav_EnsureComboOnSheetQuiet(ByVal ws As Worksheet)
    On Error Resume Next
    If Not 結果_設備ガント日付ナビ_対象シートか(ws) Then Exit Sub
    Dim ole As OLEObject
    Dim shp As Shape
    Set ole = GanttDateNav_GetOrCreateOleCombo(ws)
    If ole Is Nothing Then GoTo X
    GanttDateNav_WireHostIfNeeded ws, ole.Object
    GanttDateNav_FillMsFormsList ole.Object, ws
    
    If StrComp(ws.Name, SHEET_RESULT_EQUIP_GANTT_ACTUAL_DETAIL, vbBinaryCompare) = 0 Then
        Set shp = GanttDateNav_GetOrCreateFormUpdateButton(ws, ole)
    Else
        GanttDateNav_DeleteUpdateButtonIfPresent ws
    End If
X:
    Err.Clear
    On Error GoTo 0
End Sub

' 段階2完了時など: 設備ガント／実績明細の両方にコンボを置く（メッセージなし）
Public Sub 結果_設備ガント系_日付ジャンプコンボを両シートで確保(Optional ByVal wb As Workbook)
    Dim targetWb As Workbook
    Dim nm As Variant
    Dim wsc As Worksheet
    
    If wb Is Nothing Then
        Set targetWb = ThisWorkbook
    Else
        Set targetWb = wb
    End If
    
    For Each nm In Array(SHEET_RESULT_EQUIP_GANTT, SHEET_RESULT_EQUIP_GANTT_ACTUAL_DETAIL)
        On Error Resume Next
        Set wsc = Nothing
        Set wsc = targetWb.Worksheets(CStr(nm))
        On Error GoTo 0
        If Not wsc Is Nothing Then
            GanttDateNav_EnsureComboOnSheetQuiet wsc
        End If
    Next nm
End Sub

' アクティブシートが設備ガント系のとき、コンボをシート上に用意し日付一覧を入れる。
Public Sub 結果_設備ガント系_日付ジャンプコンボを確保()
    Dim ws As Worksheet
    
    Set ws = ActiveSheet
    If Not 結果_設備ガント日付ナビ_対象シートか(ws) Then
        MsgBox "「" & SHEET_RESULT_EQUIP_GANTT & "」または「" & SHEET_RESULT_EQUIP_GANTT_ACTUAL_DETAIL _
            & "」を表示してから実行してください。", vbExclamation, "日付へ移動"
        Exit Sub
    End If
    
    On Error GoTo FailOle
    GanttDateNav_EnsureComboOnSheetQuiet ws
    Exit Sub
    
FailOle:
    MsgBox "シート上へのコンボボックス配置に失敗しました: " & Err.Description & vbCrLf _
        & "ブックの保護・参照設定（Microsoft Forms 2.0 Object Library）を確認してください。", vbCritical, "日付へ移動"
End Sub

' 既にコンボがある場合のみリストを再構築（段階2取込直後などから呼ぶ場合用）
' 図形ボタン OnAction 用（標準モジュール名 modGanttDateNav を前提）
Public Sub GanttDateNav_RunRefreshActualDetail_Click()
    On Error Resume Next
    実績設備ガント_のみ更新_実行
    On Error GoTo 0
End Sub

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
