Option Explicit

' =========================================================
' 配台計画_タスク入力: 列設定（並び順・表示/非表示）
' - 設定シート: SHEET_COL_CONFIG_PLAN_INPUT_TASK
'   A: 並び順（小数可。反映時に整数へ詰め直し）
'   B: 列名（配台計画_タスク入力 1行目の見出し）
'   C: 表示フラグ（True/False）
' =========================================================

Private Const PLAN_INPUT_HEADER_ROW As Long = 1

Public Sub 配台計画_タスク入力_列設定_設定シート作成更新()
    Call RunWithShapeSpinner(SHAPE_PLAN_INPUT_COL_CONFIG_REFRESH, "列設定 取得中...", "配台計画_タスク入力_列設定__Core_RefreshColConfigSheet")
End Sub

Public Sub 配台計画_タスク入力_列設定_反映()
    Call RunWithShapeSpinner(SHAPE_PLAN_INPUT_COL_CONFIG_APPLY, "列設定 反映中...", "配台計画_タスク入力_列設定__Core_ApplyColConfigToTaskSheet")
End Sub

' ----------------------------
' Core: refresh config sheet
' ----------------------------
Public Sub 配台計画_タスク入力_列設定__Core_RefreshColConfigSheet()
    Dim wsTask As Worksheet
    Dim wsCfg As Worksheet
    Dim lastCol As Long
    Dim c As Long
    Dim header As String
    Dim existing As Object
    Dim writeRow As Long
    Dim nextOrder As Double
    Dim keys As Variant
    Dim oldOrd As Double
    Dim oldVis As Boolean

    On Error Resume Next
    Set wsTask = ThisWorkbook.Worksheets(SHEET_PLAN_INPUT_TASK)
    On Error GoTo 0
    If wsTask Is Nothing Then Exit Sub

    Set wsCfg = EnsurePlanInputColConfigSheet()
    If wsCfg Is Nothing Then Exit Sub

    Call EnsurePlanInputColConfigSheetHeader(wsCfg)

    ' 既存設定を記憶（同名列は order/visible を引き継ぐ）
    Set existing = LoadExistingConfig(wsCfg)

    lastCol = LastUsedColumnInRow(wsTask, PLAN_INPUT_HEADER_ROW)
    If lastCol <= 0 Then Exit Sub

    wsCfg.Range("A2:C" & wsCfg.Rows.Count).ClearContents

    writeRow = 2
    nextOrder = 1

    For c = 1 To lastCol
        header = Trim$(CStr(wsTask.Cells(PLAN_INPUT_HEADER_ROW, c).Value))
        If Len(header) = 0 Then GoTo NextHeader

        If existing.Exists(NormalizeKey(header)) Then
            keys = existing(NormalizeKey(header))
            oldOrd = CDbl(keys(0))
            oldVis = CBool(keys(1))
            wsCfg.Cells(writeRow, 1).Value = oldOrd
            wsCfg.Cells(writeRow, 2).Value = header
            wsCfg.Cells(writeRow, 3).Value = CBool(oldVis)
        Else
            wsCfg.Cells(writeRow, 1).Value = nextOrder
            wsCfg.Cells(writeRow, 2).Value = header
            wsCfg.Cells(writeRow, 3).Value = True
        End If

        writeRow = writeRow + 1
        nextOrder = nextOrder + 1
NextHeader:
    Next c

    ' 並び順でソート → 整数へ詰め直し
    Call SortAndRenumberConfig(wsCfg)

    ' 設定シートを見える化
    On Error Resume Next
    wsCfg.Visible = xlSheetVisible
    On Error GoTo 0
End Sub

' ----------------------------
' Core: apply to task sheet
' ----------------------------
Public Sub 配台計画_タスク入力_列設定__Core_ApplyColConfigToTaskSheet()
    Dim wsTask As Worksheet
    Dim wsCfg As Worksheet
    Dim entries As Collection
    Dim i As Long
    Dim targetPos As Long
    Dim colName As String
    Dim vis As Boolean
    Dim colIndex As Long
    Dim lastColAfter As Long

    On Error Resume Next
    Set wsTask = ThisWorkbook.Worksheets(SHEET_PLAN_INPUT_TASK)
    Set wsCfg = ThisWorkbook.Worksheets(SHEET_COL_CONFIG_PLAN_INPUT_TASK)
    On Error GoTo 0
    If wsTask Is Nothing Then Exit Sub
    If wsCfg Is Nothing Then
        Set wsCfg = EnsurePlanInputColConfigSheet()
        If wsCfg Is Nothing Then Exit Sub
        Call EnsurePlanInputColConfigSheetHeader(wsCfg)
        Exit Sub
    End If

    Call SortAndRenumberConfig(wsCfg)

    Set entries = ReadConfigEntries(wsCfg)
    If entries Is Nothing Or entries.Count = 0 Then Exit Sub

    On Error GoTo CleanExit
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.CutCopyMode = False

    targetPos = 1
    For i = 1 To entries.Count
        colName = CStr(entries(i)(0))
        vis = CBool(entries(i)(1))

        colIndex = FindHeaderColumn(wsTask, PLAN_INPUT_HEADER_ROW, colName)
        If colIndex <= 0 Then GoTo NextEntry

        If colIndex <> targetPos Then
            wsTask.Columns(colIndex).Cut
            wsTask.Columns(targetPos).Insert Shift:=xlToRight
            Application.CutCopyMode = False
        End If

        On Error Resume Next
        wsTask.Columns(targetPos).Hidden = Not vis
        On Error GoTo 0

        targetPos = targetPos + 1
NextEntry:
    Next i

    ' 設定対象外の列は、そのまま末尾側に残す（Hidden は現状維持）
    ' Select はアクティブシートでないと 1004 になり得るため、表示位置調整は Goto で安全に行う
    lastColAfter = LastUsedColumnInRow(wsTask, PLAN_INPUT_HEADER_ROW)
    If lastColAfter > 0 Then
        On Error Resume Next
        wsTask.Activate
        Application.Goto Reference:=wsTask.Cells(PLAN_INPUT_HEADER_ROW, 1), Scroll:=True
        On Error GoTo 0
    End If

CleanExit:
    On Error Resume Next
    Application.CutCopyMode = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub

' =========================================================
' Helpers: config sheet
' =========================================================

Private Function EnsurePlanInputColConfigSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_COL_CONFIG_PLAN_INPUT_TASK)
    On Error GoTo 0
    If ws Is Nothing Then
        On Error GoTo CreateFail
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_COL_CONFIG_PLAN_INPUT_TASK
        ws.Visible = xlSheetVisible
        Call EnsurePlanInputColConfigSheetHeader(ws)
    End If
    Set EnsurePlanInputColConfigSheet = ws
    Exit Function
CreateFail:
    Set EnsurePlanInputColConfigSheet = Nothing
End Function

Private Sub EnsurePlanInputColConfigSheetHeader(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    ws.Cells(1, 1).Value = "並び順"
    ws.Cells(1, 2).Value = "列名"
    ws.Cells(1, 3).Value = "表示"
    ws.Rows(1).Font.Bold = True
    ws.Columns(1).NumberFormatLocal = "0.########"
    ws.Columns(1).ColumnWidth = 10
    ws.Columns(2).ColumnWidth = 30
    ws.Columns(3).ColumnWidth = 10
End Sub

Private Function LoadExistingConfig(ByVal wsCfg As Worksheet) As Object
    Dim d As Object
    Dim lastRow As Long
    Dim r As Long
    Dim ordVal As Double
    Dim nm As String
    Dim vis As Boolean
    Dim key As String
    Dim arr(0 To 1) As Variant

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1 ' vbTextCompare
    If wsCfg Is Nothing Then
        Set LoadExistingConfig = d
        Exit Function
    End If

    lastRow = wsCfg.Cells(wsCfg.Rows.Count, 2).End(xlUp).Row
    If lastRow < 2 Then
        Set LoadExistingConfig = d
        Exit Function
    End If

    For r = 2 To lastRow
        nm = Trim$(CStr(wsCfg.Cells(r, 2).Value))
        If Len(nm) = 0 Then GoTo NextR
        key = NormalizeKey(nm)

        ordVal = ParseOrderValue(wsCfg.Cells(r, 1).Value, CDbl(r - 1))
        vis = ParseVisibleFlag(wsCfg.Cells(r, 3).Value, True)

        arr(0) = ordVal
        arr(1) = vis
        d(key) = arr
NextR:
    Next r

    Set LoadExistingConfig = d
End Function

Private Sub SortAndRenumberConfig(ByVal wsCfg As Worksheet)
    Dim lastRow As Long
    Dim r As Long
    Dim lastData As Long
    Dim rng As Range
    Dim ord As Double

    If wsCfg Is Nothing Then Exit Sub

    lastRow = wsCfg.Cells(wsCfg.Rows.Count, 2).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ' 途中空行があっても列名がある最終行までを対象
    lastData = lastRow

    Set rng = wsCfg.Range("A1:C" & lastData)
    On Error Resume Next
    rng.Sort Key1:=wsCfg.Range("A2"), Order1:=xlAscending, Header:=xlYes
    On Error GoTo 0

    For r = 2 To lastData
        If Len(Trim$(CStr(wsCfg.Cells(r, 2).Value))) = 0 Then Exit For
        ord = CDbl(r - 1)
        wsCfg.Cells(r, 1).Value = ord
        wsCfg.Cells(r, 3).Value = ParseVisibleFlag(wsCfg.Cells(r, 3).Value, True)
    Next r
End Sub

Private Function ReadConfigEntries(ByVal wsCfg As Worksheet) As Collection
    Dim col As New Collection
    Dim lastRow As Long
    Dim r As Long
    Dim nm As String
    Dim vis As Boolean
    Dim item(0 To 1) As Variant

    If wsCfg Is Nothing Then
        Set ReadConfigEntries = col
        Exit Function
    End If

    lastRow = wsCfg.Cells(wsCfg.Rows.Count, 2).End(xlUp).Row
    If lastRow < 2 Then
        Set ReadConfigEntries = col
        Exit Function
    End If

    For r = 2 To lastRow
        nm = Trim$(CStr(wsCfg.Cells(r, 2).Value))
        If Len(nm) = 0 Then Exit For
        vis = ParseVisibleFlag(wsCfg.Cells(r, 3).Value, True)
        item(0) = nm
        item(1) = vis
        col.Add item
    Next r

    Set ReadConfigEntries = col
End Function

Private Function ParseOrderValue(ByVal v As Variant, ByVal fallback As Double) As Double
    On Error GoTo Fail
    If IsNumeric(v) Then
        ParseOrderValue = CDbl(v)
        Exit Function
    End If
Fail:
    ParseOrderValue = fallback
End Function

Private Function ParseVisibleFlag(ByVal v As Variant, ByVal defaultVal As Boolean) As Boolean
    Dim t As String
    On Error GoTo Fail
    If VarType(v) = vbBoolean Then
        ParseVisibleFlag = CBool(v)
        Exit Function
    End If
    t = LCase$(Trim$(CStr(v)))
    If Len(t) = 0 Then ParseVisibleFlag = defaultVal: Exit Function
    If t = "true" Or t = "1" Or t = "yes" Or t = "on" Or t = "y" Or t = "はい" Then
        ParseVisibleFlag = True
        Exit Function
    End If
    If t = "false" Or t = "0" Or t = "no" Or t = "off" Or t = "n" Or t = "いいえ" Then
        ParseVisibleFlag = False
        Exit Function
    End If
Fail:
    ParseVisibleFlag = defaultVal
End Function

Private Function NormalizeKey(ByVal s As String) As String
    NormalizeKey = LCase$(Replace(Replace(Trim$(s), vbTab, ""), "　", " "))
End Function

' =========================================================
' Helpers: task sheet column operations
' =========================================================

Private Function LastUsedColumnInRow(ByVal ws As Worksheet, ByVal rowNum As Long) As Long
    On Error GoTo Fail
    If ws Is Nothing Then LastUsedColumnInRow = 0: Exit Function
    LastUsedColumnInRow = ws.Cells(rowNum, ws.Columns.Count).End(xlToLeft).Column
    Exit Function
Fail:
    LastUsedColumnInRow = 0
End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal headerName As String) As Long
    Dim lastCol As Long
    Dim c As Long
    Dim t As String
    If ws Is Nothing Then FindHeaderColumn = 0: Exit Function
    lastCol = LastUsedColumnInRow(ws, headerRow)
    If lastCol <= 0 Then FindHeaderColumn = 0: Exit Function
    For c = 1 To lastCol
        t = Trim$(CStr(ws.Cells(headerRow, c).Value))
        If StrComp(t, headerName, vbTextCompare) = 0 Then
            FindHeaderColumn = c
            Exit Function
        End If
    Next c
    FindHeaderColumn = 0
End Function

' =========================================================
' Helpers: 図形ラベル更新しながら実行（図形が無い場合は処理のみ実行）
' =========================================================

' ラッパ（スピナー更新しながら実行）
Private Sub RunWithShapeSpinner(ByVal shapeName As String, ByVal baseCaption As String, ByVal procName As String)
    Dim ws As Worksheet
    Dim shp As Shape
    Dim origText As String
    Dim sp(0 To 3) As String
    Dim i As Long

    sp(0) = "|": sp(1) = "/": sp(2) = "-": sp(3) = "\"

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_PLAN_INPUT_TASK)
    If Not ws Is Nothing Then Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If Not shp Is Nothing Then
        origText = shp.TextFrame2.TextRange.Text
        shp.TextFrame2.TextRange.Text = baseCaption & " " & sp(0)
        DoEvents
    End If

    On Error GoTo CleanFail
    For i = 0 To 1
        If Not shp Is Nothing Then
            shp.TextFrame2.TextRange.Text = baseCaption & " " & sp(i Mod 4)
            DoEvents
            Sleep 80
        End If
    Next i

    ' 実処理
    Application.Run procName

    If Not shp Is Nothing Then
        shp.TextFrame2.TextRange.Text = origText
    End If
    Exit Sub

CleanFail:
    On Error Resume Next
    If Not shp Is Nothing Then shp.TextFrame2.TextRange.Text = origText
    On Error GoTo 0
End Sub

