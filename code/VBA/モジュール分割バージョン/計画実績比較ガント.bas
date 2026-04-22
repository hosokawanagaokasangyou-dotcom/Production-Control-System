Option Explicit

' Python planning_core.ENV_COMPARE_GANTT_SNAPSHOT_DIR と同じキー
Private Const ENV_COMPARE_GANTT_SNAPSHOT_DIR As String = "COMPARE_GANTT_SNAPSHOT_DIR"
Private Const SHEET_PLAN_ACTUAL_COMPARE As String = "結果_設備ガント_計画実績比較"
Private Const OUT_COMPARE_XLSX As String = "plan_actual_compare_gantt.xlsx"

' スナップショット出力（スナップショット出力.bas）と同じ相対フォルダ名
Private Const PDF_SNAPSHOT_REL_FOLDER As String = "pdf"

' 選択 UI 用シート・コントロール名（同一ブック内で一意）
Private Const SHEET_COMPARE_PICK As String = "選択_計画実績比較"
Private Const OLE_SNAP_LIST As String = "CompareGanttSnapListBox"
' 実行ボタンは OLE の OnAction が 1004 になる環境があるため、フォームコントロール（Shape）を使用
Private Const SHAPE_COMPARE_RUN_BTN As String = "CompareGanttRunBtnForm"

' --- 公開入口 ---

' 選択用シートを作成／更新し、一覧を再取得してアクティブにする。
Public Sub 計画実績比較ガント_選択シートを表示()
    Dim targetDir As String
    On Error GoTo EH
    targetDir = ThisWorkbook.path
    If Len(targetDir) = 0 Then
        AppMsgBox "先にこのブックを保存してください。", vbExclamation, "計画実績比較ガント"
        Exit Sub
    End If
    EnsureCompareGanttPickSheet ThisWorkbook, targetDir
    ThisWorkbook.Worksheets(SHEET_COMPARE_PICK).Activate
    Exit Sub
EH:
    AppMsgBox "エラー: " & Err.Number & " / " & Err.Description, vbCritical, "計画実績比較ガント"
End Sub

' 互換: 従来名は選択シート表示へ誘導する。
Public Sub 計画実績比較ガント_スナップショット選択実行()
    計画実績比較ガント_選択シートを表示
End Sub

' リストボックスで選んだスナップショットで比較ガント生成→取り込み（実行ボタンの OnAction）。
Public Sub 計画実績比較ガント_リストから生成実行()
    Dim targetDir As String
    Dim ws As Worksheet
    Dim lo As OLEObject
    Dim lb As Object
    Dim snap As String
    
    On Error GoTo EH
    targetDir = ThisWorkbook.path
    If Len(targetDir) = 0 Then
        AppMsgBox "先にこのブックを保存してください。", vbExclamation, "計画実績比較ガント"
        Exit Sub
    End If
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_COMPARE_PICK)
    On Error GoTo EH
    If ws Is Nothing Then
        AppMsgBox "先に「計画実績比較ガント_選択シートを表示」を実行してください。", vbExclamation, "計画実績比較ガント"
        Exit Sub
    End If
    
    Set lo = FindOleOnSheet(ws, OLE_SNAP_LIST)
    If lo Is Nothing Then
        AppMsgBox "リストボックスが見つかりません。選択シートを再表示してください。", vbCritical, "計画実績比較ガント"
        Exit Sub
    End If
    Set lb = lo.Object
    
    If lb.ListCount <= 0 Then
        AppMsgBox "スナップショットがありません。pdf 配下に履歴フォルダがあるか確認してください。", vbExclamation, "計画実績比較ガント"
        Exit Sub
    End If
    If lb.ListIndex < 0 Then
        AppMsgBox "リストからスナップショットを選択してください。", vbExclamation, "計画実績比較ガント"
        Exit Sub
    End If
    
    ' 2 列目にフルパス（1 列目は表示名）
    snap = Trim$(CStr(lb.List(lb.ListIndex, 1)))
    If Len(snap) = 0 Then
        AppMsgBox "選択のパスが空です。", vbCritical, "計画実績比較ガント"
        Exit Sub
    End If
    
    RunCompareGanttPythonAndImport targetDir, snap
    Exit Sub
EH:
    AppMsgBox "エラー: " & Err.Number & " / " & Err.Description, vbCritical, "計画実績比較ガント"
End Sub

' --- 内部 ---

Private Function FindOleOnSheet(ByVal ws As Worksheet, ByVal wantName As String) As OLEObject
    Dim o As OLEObject
    Set FindOleOnSheet = Nothing
    For Each o In ws.OLEObjects
        If StrComp(o.Name, wantName, vbTextCompare) = 0 Then
            Set FindOleOnSheet = o
            Exit Function
        End If
    Next o
End Function

Private Sub DeleteOleIfExists(ByVal ws As Worksheet, ByVal nm As String)
    Dim o As OLEObject
    On Error Resume Next
    Set o = FindOleOnSheet(ws, nm)
    If Not o Is Nothing Then o.Delete
    On Error GoTo 0
End Sub

Private Sub DeleteShapeIfExists(ByVal ws As Worksheet, ByVal nm As String)
    Dim shp As Shape
    For Each shp In ws.Shapes
        If StrComp(shp.Name, nm, vbTextCompare) = 0 Then
            shp.Delete
            Exit Sub
        End If
    Next shp
End Sub

' pdf\<stamp>\ に 結果_タスク一覧.csv があるフォルダだけを降順で列挙しリストへ反映
Private Sub RefreshCompareGanttSnapshotList(ByVal pdfRoot As String, ByVal lb As Object)
    Dim stamp As String
    Dim p As String
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim tmp As String
    Dim stamps() As String
    Dim paths() As String
    Dim attr As Long
    Dim isDir As Boolean
    
    lb.Clear
    lb.ColumnCount = 2
    lb.ColumnWidths = "160 pt;0 pt"
    lb.ListStyle = 1
    
    If Len(Dir(pdfRoot, vbDirectory)) = 0 Then Exit Sub
    
    n = 0
    stamp = Dir(pdfRoot & "\*", vbDirectory)
    Do While Len(stamp) > 0
        If stamp <> "." And stamp <> ".." Then
            p = pdfRoot & "\" & stamp
            isDir = False
            On Error Resume Next
            attr = GetAttr(p)
            If Err.Number = 0 Then
                isDir = ((attr And vbDirectory) = vbDirectory)
            End If
            Err.Clear
            On Error GoTo 0
            If isDir Then
                If Len(Dir(p & "\結果_タスク一覧.csv")) > 0 Then
                    n = n + 1
                    ReDim Preserve stamps(1 To n)
                    ReDim Preserve paths(1 To n)
                    stamps(n) = stamp
                    paths(n) = p
                End If
            End If
        End If
        stamp = Dir
    Loop
    
    If n <= 0 Then Exit Sub
    
    ' スタンプ名の降順（新しい履歴が上）
    For i = 1 To n - 1
        For j = i + 1 To n
            If StrComp(stamps(i), stamps(j), vbBinaryCompare) < 0 Then
                tmp = stamps(i): stamps(i) = stamps(j): stamps(j) = tmp
                tmp = paths(i): paths(i) = paths(j): paths(j) = tmp
            End If
        Next j
    Next i
    
    For i = 1 To n
        lb.AddItem stamps(i)
        ' 2 列目にフルパス（0 行目起点の List）
        On Error Resume Next
        lb.List(lb.ListCount - 1, 1) = paths(i)
        On Error GoTo 0
    Next i
End Sub

Private Sub EnsureCompareGanttPickSheet(ByVal wb As Workbook, ByVal targetDir As String)
    Dim ws As Worksheet
    Dim pdfRoot As String
    Dim lo As OLEObject
    Dim lb As Object
    Dim shpRun As Shape
    
    pdfRoot = targetDir & "\" & PDF_SNAPSHOT_REL_FOLDER
    
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_COMPARE_PICK)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        On Error GoTo SheetNameFail
        ws.Name = SHEET_COMPARE_PICK
        On Error GoTo 0
    End If
    
    ws.Cells.Clear
    ws.Range("A1").Value = "過去配台スナップショット（pdf\日時フォルダ）から比較ガントを生成します。"
    ws.Range("A2").Value = "① 下の一覧でフォルダを選択 ②「比較ガントを生成」をクリック。"
    ws.Range("A3").Value = "※ シート保護でクリックできないときは、一時的に保護を解除してください。"
    ws.Columns("A").ColumnWidth = 90
    
    DeleteOleIfExists ws, OLE_SNAP_LIST
    ' 旧版 OLE ボタン名が残っている場合の掃除
    On Error Resume Next
    DeleteOleIfExists ws, "CompareGanttRunButton"
    On Error GoTo 0
    DeleteShapeIfExists ws, SHAPE_COMPARE_RUN_BTN
    
    Set lo = ws.OLEObjects.Add(ClassType:="Forms.ListBox.1", Left:=18, Top:=72, Width:=520, Height:=260)
    lo.Name = OLE_SNAP_LIST
    lo.Placement = 1  ' xlMoveAndSize
    Set lb = lo.Object
    lb.IntegralHeight = False
    
    RefreshCompareGanttSnapshotList pdfRoot, lb
    
    ' フォームコントロールのボタン（OLE の CommandButton では OnAction が 1004 になることがある）
    Set shpRun = ws.Shapes.AddFormControl(xlButtonControl, 18, 345, 180, 30)
    shpRun.Name = SHAPE_COMPARE_RUN_BTN
    shpRun.OnAction = "'" & wb.Name & "'!計画実績比較ガント_リストから生成実行"
    shpRun.TextFrame.Characters.Text = "比較ガントを生成"
    shpRun.Placement = 1  ' xlMoveAndSize
    
    Exit Sub
SheetNameFail:
    Err.Raise vbObjectError + 91001, , "シート名「" & SHEET_COMPARE_PICK & "」を設定できませんでした。"
End Sub

Private Sub RunCompareGanttPythonAndImport(ByVal targetDir As String, ByVal snap As String)
    Dim wsh As Object
    Dim runBat As String
    Dim exitCode As Long
    Dim refreshPath As String
    Dim targetWb As Workbook
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim ws As Worksheet
    Dim sheetName As String
    Dim hideCmd As Boolean
    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim stUnlock As Boolean
    
    On Error GoTo RunEH
    
    If Len(Dir(snap & "\結果_タスク一覧.csv")) = 0 Then
        AppMsgBox "選択フォルダに「結果_タスク一覧.csv」がありません。" & vbCrLf & snap, vbCritical, "計画実績比較ガント"
        Exit Sub
    End If
    
    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ThisWorkbook.Save
    
    stUnlock = False
    配台マクロ_全シート保護を試行解除
    stUnlock = True
    
    Set wsh = CreateObject("WScript.Shell")
    wsh.Environment("Process")("TASK_INPUT_WORKBOOK") = ThisWorkbook.FullName
    wsh.Environment("Process")(ENV_COMPARE_GANTT_SNAPSHOT_DIR) = snap
    
    On Error Resume Next
    Kill targetDir & "\log\compare_gantt_exitcode.txt"
    Err.Clear
    On Error GoTo RunEH
    
    hideCmd = Stage12CmdHideWindowEffective()
    wsh.CurrentDirectory = Environ("TEMP")
    runBat = "@echo off" & vbCrLf & "setlocal EnableDelayedExpansion" & vbCrLf & "pushd """ & targetDir & """" & vbCrLf & _
             "if not exist log mkdir log" & vbCrLf & _
             "chcp 65001>nul" & vbCrLf & _
             "echo [compare_gantt] Running plan_compare_gantt_from_snapshot.py ..." & vbCrLf & _
             "py -3 -u python\plan_compare_gantt_from_snapshot.py" & vbCrLf & _
             "set PM_CMP_EXIT=!ERRORLEVEL!" & vbCrLf & _
             "(echo !PM_CMP_EXIT!)>log\compare_gantt_exitcode.txt" & vbCrLf & _
             "echo [compare_gantt] Finished. ERRORLEVEL=!PM_CMP_EXIT!" & vbCrLf
    If Not hideCmd Then
        runBat = runBat & "if not !PM_CMP_EXIT! equ 0 (" & vbCrLf & _
                 "echo." & vbCrLf & _
                 "echo [compare_gantt] Python error. Press any key to close..." & vbCrLf & _
                 "pause" & vbCrLf & ")" & vbCrLf
    End If
    runBat = runBat & "exit /b !PM_CMP_EXIT!"
    exitCode = RunTempCmdWithConsoleLayout(wsh, runBat, Not hideCmd, hideCmd)
    
    If exitCode <> 0 Then
        AppMsgBox "Python の終了コードが " & CStr(exitCode) & " です。" & vbCrLf & _
                   "log\execution_log.txt を確認してください。", vbExclamation, "計画実績比較ガント"
        GoTo DoneProtect
    End If
    
    refreshPath = targetDir & "\output\" & OUT_COMPARE_XLSX
    If Len(Dir(refreshPath)) = 0 Then
        AppMsgBox "出力ファイルが見つかりません: " & refreshPath, vbExclamation, "計画実績比較ガント"
        GoTo DoneProtect
    End If
    
    Set targetWb = ThisWorkbook
    マクロブックから計画取込シート同源名シートを削除 targetWb, SHEET_PLAN_ACTUAL_COMPARE
    
    Set sourceWb = Workbooks.Open(refreshPath)
    sourceWb.Windows(1).Visible = False
    Set sourceWs = Nothing
    On Error Resume Next
    Set sourceWs = sourceWb.Worksheets(SHEET_PLAN_ACTUAL_COMPARE)
    On Error GoTo RunEH
    If sourceWs Is Nothing Then
        sourceWb.Close SaveChanges:=False
        AppMsgBox "取込元ブックに「" & SHEET_PLAN_ACTUAL_COMPARE & "」シートがありません。", vbCritical, "計画実績比較ガント"
        GoTo DoneProtect
    End If
    
    sheetName = Trim$(sourceWs.Name)
    sourceWs.Copy After:=targetWb.Sheets(targetWb.Sheets.Count)
    sourceWb.Close SaveChanges:=False
    Set sourceWb = Nothing
    
    Set ws = 取込ブック内のコピー先シートを取得(targetWb, sheetName)
    If ws Is Nothing Then
        AppMsgBox "シートの取り込みに失敗しました。", vbCritical, "計画実績比較ガント"
        GoTo DoneProtect
    End If
    
    AppMsgBox "「" & SHEET_PLAN_ACTUAL_COMPARE & "」を取り込みました。", vbInformation, "計画実績比較ガント"
    GoTo DoneProtect
    
RunEH:
    AppMsgBox "エラー: " & Err.Number & " / " & Err.Description, vbCritical, "計画実績比較ガント"
    Resume DoneProtect
DoneProtect:
    On Error Resume Next
    If stUnlock Then
        配台マクロ_対象シートを条件どおりに保護 targetDir
    End If
    On Error GoTo 0
    Application.DisplayAlerts = prevDisplayAlerts
    Application.ScreenUpdating = prevScreenUpdating
    Exit Sub
End Sub
