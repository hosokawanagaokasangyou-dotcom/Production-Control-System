Attribute VB_Name = "modCompareGantt"
Option Explicit

' Office.MsoFileDialogType（参照設定なしでも動かすため数値固定）
Private Const msoFileDialogFolderPicker As Long = 4

' Python planning_core.ENV_COMPARE_GANTT_SNAPSHOT_DIR と同じキー
Private Const ENV_COMPARE_GANTT_SNAPSHOT_DIR As String = "COMPARE_GANTT_SNAPSHOT_DIR"
Private Const SHEET_PLAN_ACTUAL_COMPARE As String = "結果_設備ガント_計画実績比較"
Private Const OUT_COMPARE_XLSX As String = "plan_actual_compare_gantt.xlsx"

' 過去配台スナップショット（pdf\yyyymmdd_hhnnss 等）をフォルダ選択し、比較ガントを生成して取り込む。
Public Sub 計画実績比較ガント_スナップショット選択実行()
    Dim targetDir As String
    Dim wsh As Object
    Dim fd As Object
    Dim snap As String
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
    
    On Error GoTo EH
    targetDir = ThisWorkbook.path
    If Len(targetDir) = 0 Then
        AppMsgBox "先にこのブックを保存してください。", vbExclamation, "計画実績比較ガント"
        Exit Sub
    End If
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "過去配台スナップショットのフォルダを選択（例: pdf 配下の yyyymmdd_hhnnss）"
        .InitialFileName = targetDir & "\pdf\"
        If .Show <> -1 Then Exit Sub
        snap = .SelectedItems(1)
    End With
    
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
    On Error GoTo EH
    
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
    On Error GoTo EH
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
    
EH:
    AppMsgBox "エラー: " & Err.Number & " / " & Err.Description, vbCritical, "計画実績比較ガント"
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
