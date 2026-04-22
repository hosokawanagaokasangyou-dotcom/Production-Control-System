Option Explicit

' Python planning_core.ENV_COMPARE_GANTT_SNAPSHOT_DIR と同じキー
Private Const ENV_COMPARE_GANTT_SNAPSHOT_DIR As String = "COMPARE_GANTT_SNAPSHOT_DIR"
' スナップショットに同一機械の時間重なりがあっても比較ガントを生成する（Python 側で警告のみ）
Private Const ENV_COMPARE_GANTT_ALLOW_PLAN_OVERLAP As String = "COMPARE_GANTT_ALLOW_PLAN_OVERLAP"
Private Const SHEET_PLAN_ACTUAL_COMPARE As String = "結果_設備ガント_計画実績比較"
Private Const OUT_COMPARE_XLSX As String = "plan_actual_compare_gantt.xlsx"

' スナップショット出力（スナップショット出力.bas）と同じ相対フォルダ名
Private Const PDF_SNAPSHOT_REL_FOLDER As String = "pdf"

' 選択 UI 用シート・コントロール名（同一ブック内で一意）
Private Const SHEET_COMPARE_PICK As String = "選択_計画実績比較"
' 旧版: ActiveX の Forms.ListBox（OLE）。再表示時に削除する。
Private Const OLE_SNAP_LIST As String = "CompareGanttSnapListBox"
' 一覧はフォームのリストボックス（ActiveX よりデザインモード問題が出にくい）
Private Const SHAPE_COMPARE_SNAP_LIST As String = "CompareGanttSnapListForm"
' 実行ボタンは OLE の OnAction が 1004 になる環境があるため、フォームコントロール（Shape）を使用
Private Const SHAPE_COMPARE_RUN_BTN As String = "CompareGanttRunBtnForm"
Private Const COMPARE_GANTT_DAY_ROW_MAP_START As Long = 500
Private Const COMPARE_GANTT_DAY_ROW_MAP_DATE_COL As Long = 52   ' AZ
Private Const COMPARE_GANTT_DAY_ROW_MAP_FIRSTROW_COL As Long = 53 ' BA
' フォームリストと同じ並びで Z 列にフルパスを格納（1 行目＝リスト先頭）
Private Const PICK_LIST_DATA_START_ROW As Long = 5
Private Const PICK_SNAP_ROWS_MAX As Long = 500
Private Const PICK_PATH_COL As Long = 26  ' 列 Z

' --- 公開入口 ---

' 選択用シートを作成／更新し、一覧を再取得してアクティブにする。
' Power Query / データ接続は業務ロジックと同様 TryRefreshWorkbookQueries で一括更新してから表示する。
' 試行順パターン系と同様、接続更新中は MacroSplash で進捗を表示する。
Public Sub 計画実績比較ガント_選択シートを表示()
    Dim targetDir As String
    Dim errNum As Long
    Dim errDesc As String
    On Error GoTo EH
    targetDir = ThisWorkbook.path
    If Len(targetDir) = 0 Then
        AppMsgBox "先にこのブックを保存してください。", vbExclamation, "計画実績比較ガント"
        Exit Sub
    End If
    MacroSplash_Show "計画実績比較ガント: 選択シートを表示しています…", False
    MacroSplash_SetStep "データ接続（Power Query 等）を更新しています…"
    If Not TryRefreshWorkbookQueries() Then
        MacroSplash_Hide
        AppMsgBox "データ接続の更新に失敗したため中断しました。" & vbCrLf & m_lastRefreshQueriesErrMsg, vbExclamation, "計画実績比較ガント"
        Exit Sub
    End If
    MacroSplash_SetStep "選択用シートを作成・更新しています…"
    EnsureCompareGanttPickSheet ThisWorkbook, targetDir
    ThisWorkbook.Worksheets(SHEET_COMPARE_PICK).Activate
    MacroSplash_Hide
    Exit Sub
EH:
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    If m_macroSplashShown Then MacroSplash_Hide
    On Error GoTo 0
    AppMsgBox "エラー: " & CStr(errNum) & " / " & errDesc, vbCritical, "計画実績比較ガント"
End Sub

' 互換: 従来名は選択シート表示へ誘導する。
Public Sub 計画実績比較ガント_スナップショット選択実行()
    計画実績比較ガント_選択シートを表示
End Sub

' B1 の表示日に対応する日ブロック先頭行へスクロール（Python が AZ/BA に書いたマップを参照）。
Public Sub 計画実績比較ガント_表示日へジャンプ()
    Dim ws As Worksheet
    Dim want As String
    Dim i As Long
    Dim cellD As Variant
    Dim jumpR As Variant
    On Error GoTo EH
    Set ws = ThisWorkbook.Worksheets(SHEET_PLAN_ACTUAL_COMPARE)
    want = CompareGanttB1ValueAsIsoDate(ws.Range("B1").Value)
    If Len(want) = 0 Then Exit Sub
    For i = 0 To 399
        cellD = ws.Cells(COMPARE_GANTT_DAY_ROW_MAP_START + i, COMPARE_GANTT_DAY_ROW_MAP_DATE_COL).Value
        If IsEmpty(cellD) Then Exit Sub
        If Trim$(CStr(cellD)) = want Then
            jumpR = ws.Cells(COMPARE_GANTT_DAY_ROW_MAP_START + i, COMPARE_GANTT_DAY_ROW_MAP_FIRSTROW_COL).Value
            If Not IsEmpty(jumpR) And IsNumeric(jumpR) Then
                If CLng(jumpR) > 0 Then
                    ws.Activate
                    Application.ActiveWindow.ScrollRow = CLng(jumpR)
                    ws.Cells(CLng(jumpR), 2).Select
                End If
            End If
            Exit Sub
        End If
    Next i
    Exit Sub
EH:
End Sub

' ThisWorkbook の Workbook_SheetChange から呼ぶ（B1 変更で自動ジャンプ）。
Public Sub 計画実績比較ガント_WorkbookSheetChange入口(ByVal Sh As Object, ByVal Target As Range)
    Dim prevEv As Boolean
    On Error GoTo EH
    If TypeName(Sh) <> "Worksheet" Then Exit Sub
    If StrComp(Sh.Name, SHEET_PLAN_ACTUAL_COMPARE, vbTextCompare) <> 0 Then Exit Sub
    If Target Is Nothing Then Exit Sub
    If Intersect(Target, Sh.Range("B1")) Is Nothing Then Exit Sub
    prevEv = Application.EnableEvents
    Application.EnableEvents = False
    計画実績比較ガント_表示日へジャンプ
    Application.EnableEvents = prevEv
    Exit Sub
EH:
    On Error Resume Next
    Application.EnableEvents = True
    On Error GoTo 0
End Sub

' リストボックスで選んだスナップショットで比較ガント生成→取り込み（実行ボタンの OnAction）。
Public Sub 計画実績比較ガント_リストから生成実行()
    Dim targetDir As String
    Dim ws As Worksheet
    Dim lo As OLEObject
    Dim lb As Object
    Dim shp As Shape
    Dim snap As String
    Dim li As Long
    
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
    
    Set shp = FindCompareSnapListShape(ws)
    If Not shp Is Nothing Then
        With shp.ControlFormat
            If .ListCount <= 0 Then
                AppMsgBox "スナップショットがありません。pdf 配下に履歴フォルダがあるか確認してください。", vbExclamation, "計画実績比較ガント"
                Exit Sub
            End If
            li = .ListIndex
            If li < 1 Then
                AppMsgBox "リストからスナップショットを選択してください。", vbExclamation, "計画実績比較ガント"
                Exit Sub
            End If
        End With
        snap = Trim$(CStr(ws.Cells(PICK_LIST_DATA_START_ROW + li - 1, PICK_PATH_COL).Value))
        If Len(snap) = 0 Then
            AppMsgBox "選択のパスが空です。選択シートを再表示してください。", vbCritical, "計画実績比較ガント"
            Exit Sub
        End If
        RunCompareGanttPythonAndImport targetDir, snap
        Exit Sub
    End If
    
    ' 互換: 旧 ActiveX リストのみが残っているブック
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

Private Function CompareGanttB1ValueAsIsoDate(ByVal v As Variant) As String
    On Error Resume Next
    CompareGanttB1ValueAsIsoDate = vbNullString
    If IsEmpty(v) Then Exit Function
    If IsDate(v) Then
        CompareGanttB1ValueAsIsoDate = Format$(CDate(v), "yyyy-mm-dd")
    Else
        CompareGanttB1ValueAsIsoDate = Trim$(Replace(Replace(CStr(v), "/", "-"), ".", "-"))
    End If
    On Error GoTo 0
End Function

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

Private Function FindCompareSnapListShape(ByVal ws As Worksheet) As Shape
    Dim sh As Shape
    Set FindCompareSnapListShape = Nothing
    For Each sh In ws.Shapes
        If StrComp(sh.Name, SHAPE_COMPARE_SNAP_LIST, vbTextCompare) = 0 Then
            Set FindCompareSnapListShape = sh
            Exit Function
        End If
    Next sh
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

' 配台側の全シート保護が DrawingObjects:=True のとき、OLE リスト／フォームボタンがクリック不能になることがある。
Private Sub TryUnprotectSheetAnyPassword(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Unprotect Password:=""
    ws.Unprotect Password:=SHEET_FONT_UNPROTECT_PASSWORD
    On Error GoTo 0
End Sub

' 図形は操作可（DrawingObjects:=False）、セルは保護（説明セルは編集可のため全セルロック解除）
Private Sub ProtectComparePickSheetForUi(ByVal ws As Worksheet)
    Dim pwd As String
    pwd = SHEET_FONT_UNPROTECT_PASSWORD
    On Error Resume Next
    TryUnprotectSheetAnyPassword ws
    Err.Clear
    ws.Cells.Locked = False
    ws.Protect Password:=pwd, DrawingObjects:=False, Contents:=True, UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

' 全シート再保護のあと、選択シートだけリスト／ボタンが使える状態に戻す
Private Sub RestoreComparePickSheetAfterWorkbookProtect(ByVal wb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_COMPARE_PICK)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    TryUnprotectSheetAnyPassword ws
    ProtectComparePickSheetForUi ws
End Sub

' 結果_設備ガント_計画実績比較: 表示日（B1）のデータ検証を操作するため A1:C3 をロック解除し、図形操作可で再保護
Private Sub ProtectPlanActualCompareSheetForUi(ByVal ws As Worksheet)
    Dim pwd As String
    pwd = SHEET_FONT_UNPROTECT_PASSWORD
    On Error Resume Next
    TryUnprotectSheetAnyPassword ws
    Err.Clear
    ws.Cells.Locked = True
    ws.Range("A1:C3").Locked = False
    ws.Protect Password:=pwd, DrawingObjects:=False, Contents:=True, UserInterfaceOnly:=True
    On Error GoTo 0
    ' 旧版で置いていた「該当日へジャンプ」フォームボタン（B1 変更は Workbook_SheetChange で処理）
    On Error Resume Next
    DeleteShapeIfExists ws, "CompareGanttJumpToDateBtnForm"
    On Error GoTo 0
End Sub

Private Sub RestorePlanActualCompareSheetAfterWorkbookProtect(ByVal wb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_PLAN_ACTUAL_COMPARE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    TryUnprotectSheetAnyPassword ws
    ProtectPlanActualCompareSheetForUi ws
End Sub

' pdf\<stamp>\ に 結果_タスク一覧.csv があるフォルダだけを降順で列挙しリストへ反映
' フォームのリストボックスは表示名のみ。フルパスは ws の Z 列に同順で格納する。
Private Sub RefreshCompareGanttSnapshotList(ByVal ws As Worksheet, ByVal listShp As Shape, ByVal pdfRoot As String)
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
    Dim fso As Object
    Dim hasCsv As Boolean
    Dim cf As ControlFormat
    
    ws.Range(ws.Cells(PICK_LIST_DATA_START_ROW, PICK_PATH_COL), _
             ws.Cells(PICK_LIST_DATA_START_ROW + PICK_SNAP_ROWS_MAX - 1, PICK_PATH_COL)).ClearContents
    On Error Resume Next
    listShp.ControlFormat.RemoveAllItems
    On Error GoTo 0
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
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
                ' ネストした Dir は列挙状態を壊すため、CSV 存在は FileSystemObject で確認する
                hasCsv = fso.FileExists(p & "\結果_タスク一覧.csv")
                If hasCsv Then
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
    
    Set cf = listShp.ControlFormat
    For i = 1 To n
        cf.AddItem stamps(i)
        ws.Cells(PICK_LIST_DATA_START_ROW + i - 1, PICK_PATH_COL).Value = paths(i)
    Next i
End Sub

Private Sub EnsureCompareGanttPickSheet(ByVal wb As Workbook, ByVal targetDir As String)
    Dim ws As Worksheet
    Dim pdfRoot As String
    Dim shpList As Shape
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
    
    TryUnprotectSheetAnyPassword ws
    
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
    DeleteShapeIfExists ws, SHAPE_COMPARE_SNAP_LIST
    DeleteShapeIfExists ws, SHAPE_COMPARE_RUN_BTN
    
    ' ActiveX の OLE ではなくフォームのリストボックス（デザインモード問題を避ける）
    Set shpList = ws.Shapes.AddFormControl(xlListBox, 18, 72, 520, 260)
    shpList.Name = SHAPE_COMPARE_SNAP_LIST
    shpList.Placement = 1  ' xlMoveAndSize
    shpList.Locked = False
    
    RefreshCompareGanttSnapshotList ws, shpList, pdfRoot
    ws.Columns(PICK_PATH_COL).Hidden = True
    
    On Error Resume Next
    ws.Range("A1:A3").Font.Name = BIZ_UDP_GOTHIC_FONT_NAME
    On Error GoTo 0
    
    ' フォームコントロールのボタン（OLE の CommandButton では OnAction が 1004 になることがある）
    'Set shpRun = ws.Shapes.AddFormControl(xlButtonControl, 18, 345, 180, 30)
    'shpRun.Name = SHAPE_COMPARE_RUN_BTN
    'shpRun.Locked = False
    'shpRun.OnAction = "'" & wb.Name & "'!計画実績比較ガント_リストから生成実行"
    'shpRun.TextFrame.Characters.Text = "比較ガントを生成"
    'shpRun.Placement = 1  ' xlMoveAndSize
    
    ProtectComparePickSheetForUi ws
    
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
    wsh.Environment("Process")(ENV_COMPARE_GANTT_ALLOW_PLAN_OVERLAP) = "1"
    
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
    
    ProtectPlanActualCompareSheetForUi ws
    
    AppMsgBox "「" & SHEET_PLAN_ACTUAL_COMPARE & "」を取り込みました。", vbInformation, "計画実績比較ガント"
    GoTo DoneProtect
    
RunEH:
    AppMsgBox "エラー: " & Err.Number & " / " & Err.Description, vbCritical, "計画実績比較ガント"
    Resume DoneProtect
DoneProtect:
    On Error Resume Next
    If stUnlock Then
        配台マクロ_対象シートを条件どおりに保護 targetDir
        RestoreComparePickSheetAfterWorkbookProtect ThisWorkbook
        RestorePlanActualCompareSheetAfterWorkbookProtect ThisWorkbook
    End If
    On Error GoTo 0
    Application.DisplayAlerts = prevDisplayAlerts
    Application.ScreenUpdating = prevScreenUpdating
    Exit Sub
End Sub
