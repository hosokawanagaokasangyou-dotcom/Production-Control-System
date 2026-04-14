Attribute VB_Name = "modGanttDateNav"
Option Explicit

' 結果_設備ガント／結果_設備ガント_実績明細 の列 A 先頭行（【yyyy/mm/dd】）を走査し、
' UF_GanttDateNav のリストに「表示ラベル」と「先頭行番号（隠し列）」を詰める。

Public mGanttDateNavFillBusy As Boolean

Private Const GANTT_DATE_BANNER_L As String = "【"
Private Const GANTT_DATE_BANNER_R As String = "】"

Public Function 結果_設備ガント日付ナビ_対象シートか(ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then Exit Function
    If StrComp(ws.Name, SHEET_RESULT_EQUIP_GANTT, vbBinaryCompare) = 0 Then
        結果_設備ガント日付ナビ_対象シートか = True
    ElseIf StrComp(ws.Name, SHEET_RESULT_EQUIP_GANTT_ACTUAL_DETAIL, vbBinaryCompare) = 0 Then
        結果_設備ガント日付ナビ_対象シートか = True
    End If
End Function

Public Sub GanttDateNav_FillListBox(ByVal lst As Object, ByVal ws As Worksheet)
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

' リストで日付を選ぶと、その日ブロック先頭行の列 A をアクティブにする（UF_GanttDateNav.lstDates_Change）。
Public Sub 結果_設備ガント系_日付ジャンプフォームを表示()
    Dim ws As Worksheet
    
    Set ws = ActiveSheet
    If Not 結果_設備ガント日付ナビ_対象シートか(ws) Then
        MsgBox "「" & SHEET_RESULT_EQUIP_GANTT & "」または「" & SHEET_RESULT_EQUIP_GANTT_ACTUAL_DETAIL _
            & "」を表示してから実行してください。", vbExclamation, "日付へ移動"
        Exit Sub
    End If
    
    On Error GoTo FailOpen
    Load UF_GanttDateNav
    Set UF_GanttDateNav.TargetWs = ws
    UF_GanttDateNav.Show vbModeless
    Exit Sub
    
FailOpen:
    MsgBox "ユーザーフォーム「UF_GanttDateNav」がプロジェクトに無いか、.frm/.frx のインポートに失敗しています。" & vbCrLf _
        & "モジュール分割バージョンの「UF_GanttDateNav_インポート手順.txt」を参照してください。", vbCritical, "日付へ移動"
End Sub
