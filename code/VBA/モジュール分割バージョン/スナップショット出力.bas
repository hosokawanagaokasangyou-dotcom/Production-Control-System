Option Explicit

' pdf\ 直下 … 常に参照する最新版（上書き）
' pdf\yyyymmdd_hhnnss\ … 履歴（削除しない）。中身は直下と同じファイル名

Private Const FMT_CSV_UTF8 As Long = 62 ' xlCSVUTF8（Excel 2019 以降）
' 共有フォルダ（社内サーバー）の到達性確認: 0.5 秒 ping の宛先（IP推奨）。空なら UNC のホスト名/アドレスを使用。
Private Const SNAPSHOT_SHARE_PING_HOST_OVERRIDE As String = "192.168.0.101"

Private Sub EnsureFolder(ByVal folderPath As String)
    If Len(Dir(folderPath, vbDirectory)) > 0 Then Exit Sub
    On Error Resume Next
    MkDir folderPath
    On Error GoTo 0
End Sub

Private Function SnapshotStampDirName() As String
    SnapshotStampDirName = Format$(Year(Now), "0000") & Format$(Month(Now), "00") & Format$(Day(Now), "00") & "_" & _
        Format$(Hour(Now), "00") & Format$(Minute(Now), "00") & Format$(Second(Now), "00")
End Function

Private Sub TryDeleteFile(ByVal fp As String)
    On Error Resume Next
    If Len(Dir(fp)) > 0 Then Kill fp
    On Error GoTo 0
End Sub

Private Sub CopySnapshotToRoot(ByVal snapPath As String, ByVal rootPath As String)
    If Len(Dir(snapPath)) = 0 Then Exit Sub
    On Error Resume Next
    TryDeleteFile rootPath
    FileCopy snapPath, rootPath
    On Error GoTo 0
End Sub

Private Function TryExtractUncHost(ByVal p As String) As String
    Dim s As String
    Dim k As Long
    s = Trim$(p)
    TryExtractUncHost = ""
    If Len(s) < 3 Then Exit Function
    If Left$(s, 2) <> "\\" Then Exit Function
    s = Mid$(s, 3)
    k = InStr(1, s, "\", vbBinaryCompare)
    If k <= 1 Then Exit Function
    TryExtractUncHost = Left$(s, k - 1)
End Function

Private Function PingHostFast500msCached(ByVal host As String) As Boolean
    Static cachedHost As String
    Static cachedOk As Boolean
    Static cachedAt As Single
    Dim nowT As Single
    Dim rc As Long
    Dim cmd As String
    
    PingHostFast500msCached = True
    host = Trim$(host)
    If Len(host) = 0 Then Exit Function
    
    nowT = Timer
    If StrComp(cachedHost, host, vbTextCompare) = 0 Then
        ' 同一実行内での多重コピーを想定し、直近 30 秒は再 ping しない
        If nowT >= cachedAt And (nowT - cachedAt) < 30# Then
            PingHostFast500msCached = cachedOk
            Exit Function
        End If
    End If
    
    ' ping -n 1: 1回だけ / -w 500: タイムアウト 500ms
    cmd = "cmd /c ping -n 1 -w 500 " & host & " >nul"
    On Error Resume Next
    rc = CreateObject("WScript.Shell").Run(cmd, 0, True)
    On Error GoTo 0
    
    cachedHost = host
    cachedAt = nowT
    cachedOk = (rc = 0)
    PingHostFast500msCached = cachedOk
End Function

Private Function SharedPdfRootFromSetting(ByVal shareRoot As String, ByVal relPdfFolder As String) As String
    Dim s As String
    Dim seg As String
    Dim k As Long
    
    s = Trim$(shareRoot)
    SharedPdfRootFromSetting = ""
    If Len(s) = 0 Then Exit Function
    
    ' 末尾 \ を除去（\\server\share\ のようなUNCでも最後の\だけ落とす）
    Do While Right$(s, 1) = "\" Or Right$(s, 1) = "/"
        s = Left$(s, Len(s) - 1)
        If Len(s) = 0 Then Exit Function
    Loop
    
    ' E10 が既に ...\pdf を指している場合は、その直下をコピー先ルートとする（pdf\pdf を作らない）
    k = InStrRev(s, "\")
    If k > 0 And k < Len(s) Then
        seg = LCase$(Mid$(s, k + 1))
        If seg = LCase$(relPdfFolder) Then
            SharedPdfRootFromSetting = s
            Exit Function
        End If
    End If
    
    ' それ以外は親フォルダとみなし、pdf を付与
    SharedPdfRootFromSetting = s & "\" & relPdfFolder
End Function

Private Sub CopySnapshotToSharedIfConfigured(ByVal wb As Workbook, ByVal relPdfFolder As String, ByVal stamp As String, ByVal fileName As String, ByVal snapPath As String)
    Dim wsSet As Worksheet
    Dim shareRoot As String
    Dim sharePdfRoot As String
    Dim shareSnapDir As String
    Dim shareLatestPath As String
    Dim shareStampPath As String
    Dim host As String
    
    If wb Is Nothing Then Exit Sub
    If Len(fileName) = 0 Then Exit Sub
    If Len(snapPath) = 0 Then Exit Sub
    If Len(Dir(snapPath)) = 0 Then Exit Sub
    
    On Error Resume Next
    Set wsSet = wb.Worksheets(SHEET_SETTINGS)
    On Error GoTo 0
    If wsSet Is Nothing Then Exit Sub
    
    shareRoot = Trim$(CStr(wsSet.Range("E10").Value))
    If Len(shareRoot) = 0 Then Exit Sub

    sharePdfRoot = SharedPdfRootFromSetting(shareRoot, relPdfFolder)
    If Len(sharePdfRoot) = 0 Then Exit Sub
    
    host = Trim$(SNAPSHOT_SHARE_PING_HOST_OVERRIDE)
    If Len(host) = 0 Then host = TryExtractUncHost(sharePdfRoot)
    If Len(host) > 0 Then If Not PingHostFast500msCached(host) Then Exit Sub
    
    EnsureFolder sharePdfRoot
    shareSnapDir = sharePdfRoot & "\" & stamp
    EnsureFolder shareSnapDir
    
    shareLatestPath = sharePdfRoot & "\" & fileName
    shareStampPath = shareSnapDir & "\" & fileName
    
    On Error Resume Next
    TryDeleteFile shareLatestPath
    FileCopy snapPath, shareStampPath
    FileCopy snapPath, shareLatestPath
    On Error GoTo 0
End Sub

Private Function WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    Set ws = Nothing
    Err.Clear
    On Error GoTo 0
End Function

Private Sub ExportSheetPdfIfExists(ByVal wb As Workbook, ByVal sheetName As String, ByVal pdfPath As String)
    Dim ws As Worksheet
    If Not WorksheetExists(wb, sheetName) Then Exit Sub
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
    Set ws = Nothing
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub ExportSheetCsvIfExists(ByVal wb As Workbook, ByVal sheetName As String, ByVal csvPath As String)
    Dim ws As Worksheet
    Dim tmpWb As Workbook
    Dim prevAlerts As Boolean
    Set tmpWb = Nothing
    If Not WorksheetExists(wb, sheetName) Then Exit Sub
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error GoTo CsvCleanup
    Set ws = wb.Worksheets(sheetName)
    ws.Copy
    Set tmpWb = ActiveWorkbook
    If tmpWb Is Nothing Then GoTo CsvCleanup
    On Error Resume Next
    tmpWb.SaveAs Filename:=csvPath, FileFormat:=FMT_CSV_UTF8, Local:=True
    If Err.Number <> 0 Then
        Err.Clear
        tmpWb.SaveAs Filename:=csvPath, FileFormat:=xlCSV, Local:=True
    End If
CsvCleanup:
    On Error Resume Next
    If Not tmpWb Is Nothing Then
        tmpWb.Close SaveChanges:=False
        Set tmpWb = Nothing
    End If
    Set ws = Nothing
    Application.DisplayAlerts = prevAlerts
    Err.Clear
    On Error GoTo 0
End Sub

' targetDir … マクロ実行ブック直下（ThisWorkbook.path）
' wb … 通常 ThisWorkbook（加工計画DATA・配台計画・結果シートを同一ブックで参照）
Public Sub スナップショット_pdfとcsvを出力(ByVal targetDir As String, ByVal wb As Workbook)
    Dim pdfRoot As String
    Dim stamp As String
    Dim snapDir As String
    Dim f As String
    
    If Len(targetDir) = 0 Then Exit Sub
    If wb Is Nothing Then Exit Sub
    
    pdfRoot = targetDir & "\" & PDF_SNAPSHOT_REL_FOLDER
    EnsureFolder pdfRoot
    
    stamp = SnapshotStampDirName()
    snapDir = pdfRoot & "\" & stamp
    EnsureFolder snapDir
    
    ' --- PDF（計画+実績の設備ガント、実績明細ガント）---
    f = SHEET_RESULT_EQUIP_GANTT & ".pdf"
    ExportSheetPdfIfExists wb, SHEET_RESULT_EQUIP_GANTT, snapDir & "\" & f
    CopySnapshotToRoot snapDir & "\" & f, pdfRoot & "\" & f
    CopySnapshotToSharedIfConfigured wb, PDF_SNAPSHOT_REL_FOLDER, stamp, f, snapDir & "\" & f
    
    f = SHEET_RESULT_EQUIP_GANTT_ACTUAL_DETAIL & ".pdf"
    ExportSheetPdfIfExists wb, SHEET_RESULT_EQUIP_GANTT_ACTUAL_DETAIL, snapDir & "\" & f
    CopySnapshotToRoot snapDir & "\" & f, pdfRoot & "\" & f
    CopySnapshotToSharedIfConfigured wb, PDF_SNAPSHOT_REL_FOLDER, stamp, f, snapDir & "\" & f
    
    ' --- CSV ---
    f = SHEET_RESULT_TASK_LIST & ".csv"
    ExportSheetCsvIfExists wb, SHEET_RESULT_TASK_LIST, snapDir & "\" & f
    CopySnapshotToRoot snapDir & "\" & f, pdfRoot & "\" & f
    CopySnapshotToSharedIfConfigured wb, PDF_SNAPSHOT_REL_FOLDER, stamp, f, snapDir & "\" & f
    
    f = SHEET_PLAN_INPUT_TASK & ".csv"
    ExportSheetCsvIfExists wb, SHEET_PLAN_INPUT_TASK, snapDir & "\" & f
    CopySnapshotToRoot snapDir & "\" & f, pdfRoot & "\" & f
    CopySnapshotToSharedIfConfigured wb, PDF_SNAPSHOT_REL_FOLDER, stamp, f, snapDir & "\" & f
    
    f = SHEET_TASKS_RAW_PLAN_DATA & ".csv"
    ExportSheetCsvIfExists wb, SHEET_TASKS_RAW_PLAN_DATA, snapDir & "\" & f
    CopySnapshotToRoot snapDir & "\" & f, pdfRoot & "\" & f
    CopySnapshotToSharedIfConfigured wb, PDF_SNAPSHOT_REL_FOLDER, stamp, f, snapDir & "\" & f
    
    f = SHEET_ACTUAL_DETAIL_DATA & ".csv"
    ExportSheetCsvIfExists wb, SHEET_ACTUAL_DETAIL_DATA, snapDir & "\" & f
    CopySnapshotToRoot snapDir & "\" & f, pdfRoot & "\" & f
    CopySnapshotToSharedIfConfigured wb, PDF_SNAPSHOT_REL_FOLDER, stamp, f, snapDir & "\" & f
End Sub

' 手動用の無引数エントリは 業務ロジック の「スナップショット_手動でpdfとcsv出力」（図形 OnAction・他モジュールから常に解決できるよう同じブックの標準モジュール側に置く）。
