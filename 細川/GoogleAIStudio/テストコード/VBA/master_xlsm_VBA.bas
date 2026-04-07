#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

' ?Z???J?????_?[: ???t?E?Éü?????L????? A?i?????o?[???j?` AG?i???t??? B?`AG?j
Private Const CALENDAR_LAST_DATA_COL As Long = 33  ' AG??

' ?????o?[??o?É£?iWriteMemberAttendanceSheet?j: 1 ?s????o???E?f?[?^?? A ??u???t?v?` K ??
' ?i????????? Clear ?O?? A2?i??????? A1?j??t?H???g???E?T?C?Y???????A?r???E???o?????`????????j
Private Const MEMBER_ATT_HEADER_A1 As String = "???t"
Private Const MEMBER_ATT_FIRST_DATA_ROW As Long = 2
Private Const MEMBER_ATT_LAST_INPUT_COL As Long = 11   ' K?i???l?E?c??I????j
Private mAttRowHL_SheetName As String
Private mAttRowHL_DataRow As Long

' ?S?V?[?g?t?H???g????}?N???iPC ??t?H???g???????????K?v?????????j
Private Const BIZ_UDP_GOTHIC_FONT_NAME As String = "BIZ UDP?S?V?b?N"
' ?t?H???g?I???_?C?A???O?p????V?[?g?iVeryHidden?j
Private Const SCRATCH_SHEET_FONT As String = "_FontPick"
' ?S?V?[?g?t?H???g???????V?[?g????????????????p?X???[?h?i??p?X???[?h???????B?p?X???[?h?t???V?[?g?????????j
Private Const SHEET_FONT_UNPROTECT_PASSWORD As String = ""
' ?z??u?b?N??????V?[?g????????A??I?[?g?t?B?b?g??????O???b?h????????????O
Private Const SHEET_RESULT_EQUIP_GANTT As String = "????_????K???g"
' =========================================================
' ?g??????\?iskills ?? OP/AS ?? ?~ need ??K?v?l??????AOP1?????????g???????j
' ?EA??u?g?????sID?v?c ?f?[?^?s????? 1 ??????????i?z????p???????_?^?X?N?????????g?????\#?????]?L?j
' ?E?V?[?g???u?g??????\?v?c skills ?V?[?g?????iPython ?? TEAM_ASSIGN_USE_MASTER_COMBO_SHEET ?????j
' ??????K?????W???[??????u???iSub ???? Private Const ??u????R???p?C???G???[?j
' =========================================================
Private Const SHEET_SKILL_COMBINATIONS As String = "?g??????\"
Private Const SHEET_SKILLS_DATA As String = "skills"
Private Const SHEET_NEED_DATA As String = "need"
' =========================================================
' ?@?B?J?????_?[?i?V?[?g 1 ???u?@?B?J?????_?[?v?j
' ?E1 ?s????H?????A2 ?s????@?B???iskills ?????????????ÑÉE?l??]?L?j?B???o????F?t??
' ?EA ????t?{?????i1 ????P??E???C?? A12/B12 ?????`????????A???? 6:00?`22:00?j?^B ??j???i????j
' ?EC ???~??skills ????^??H?????E?@?B???B?f?[?^?s?????????????n??X?g???C?v?i????j
' ?E?????: ??????Z????????X???b?g?????s??i?R?????g?j?B???????????~?H???~?@?B?L?[?????
' =========================================================
Private Const SHEET_MACHINE_CALENDAR As String = "?@?B?J?????_?["
' ?i?K2 Python?iplanning_core._core?j?????C??V?[?g?B?V?[?g????R?[?h?????v?????Á∑??B
Private Const SHEET_MACHINE_CHANGEOVER As String = "???_??????O?áq??"
Private Const SHEET_MACHINE_DAILY_STARTUP As String = "???_?@?B_?????n?????"
Private Const MACHINE_CAL_HORIZON_DAYS As Long = 365
Private Const MACHINE_CAL_SLOT_HOUR_FIRST As Long = 6
Private Const MACHINE_CAL_SLOT_HOUR_LAST As Long = 22
' 6?`22 ??????????i17 ?s/???j?c A12/B12 ????äŒ?????B??äŒ?? (?I?????|?J?n??)+1 ?s/??
Private Const MACHINE_CAL_SLOTS_PER_DAY As Long = 17
' ???C???V?[?g D ??: skills A ??i3 ?s??`?j??????o?[??????A?????V?[?g??Éü?? A1 ??? HYPERLINK ??
' ?imaster_???C????Éü?????N???X?g???X?V?A????? master_?J?????_?[???????o?[?o?É£??? ????????X?V?j
Private Const MAIN_ATTENDANCE_LINK_COL As Long = 4
Private Const MAIN_ATTENDANCE_LINK_FIRST_DATA_ROW As Long = 2
Private Const MAIN_ATTENDANCE_LINK_CLEAR_MAX_ROWS As Long = 400
' Ctrl+Shift+?e???L?[ - ?? ???C????iApplication.OnKey?j?B{109}=vbKeySubtract?i?e???L?[ -?j?B{SUBTRACT} ???????? OnKey ?? 1004 ????s????????l?R?[?h???g?p
Private Const SHORTCUT_MAIN_SHEET_ONKEY As String = "^+{109}"

' ???C???V?[?g?i?u???C???v?u???C??_?v?uMain?v?A???????O??u???C???v?????j???èÔ
' ???u?Z?????C???J?????_?[?v???????p????????A???O??u?J?????_?[?v?????V?[?g??t?H?[???o?b?N???O?B
' ?????????????????V?[?g??????Z??????D??i????v??u?b?N?????????????O???????j?B
Private Function MasterGetMainWorksheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    Dim best As Worksheet
    Dim bestLen As Long
    Dim L As Long
    On Error Resume Next
    Set ws = wb.Worksheets("???C??")
    On Error GoTo 0
    If Not ws Is Nothing Then
        Set MasterGetMainWorksheet = ws
        Exit Function
    End If
    On Error Resume Next
    Set ws = wb.Worksheets("???C??_")
    On Error GoTo 0
    If Not ws Is Nothing Then
        Set MasterGetMainWorksheet = ws
        Exit Function
    End If
    On Error Resume Next
    Set ws = wb.Worksheets("Main")
    On Error GoTo 0
    If Not ws Is Nothing Then
        Set MasterGetMainWorksheet = ws
        Exit Function
    End If
    Set best = Nothing
    bestLen = 10000
    For Each ws In wb.Worksheets
        If InStr(ws.Name, "???C??") > 0 Then
            If InStr(ws.Name, "?J?????_?[") > 0 Then GoTo NextMainCand
            L = Len(ws.Name)
            If L < bestLen Then
                bestLen = L
                Set best = ws
            End If
        End If
NextMainCand:
    Next ws
    Set MasterGetMainWorksheet = best
End Function

Private Sub MasterSelectMainSheetA1()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = MasterGetMainWorksheet(ThisWorkbook)
    If ws Is Nothing Then Exit Sub
    ws.Activate
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

Public Sub MasterShortcutMainSheet_CtrlShift0()
    On Error Resume Next
    If Not ActiveWorkbook Is ThisWorkbook Then Exit Sub
    MasterSelectMainSheetA1
    On Error GoTo 0
End Sub

Public Sub MasterShortcutMainSheet_OnKeyRegister()
    On Error Resume Next
    Application.OnKey Key:=SHORTCUT_MAIN_SHEET_ONKEY, Procedure:="MasterShortcutMainSheet_CtrlShift0"
    On Error GoTo 0
End Sub

Public Sub MasterShortcutMainSheet_OnKeyUnregister()
    On Error Resume Next
    Application.OnKey Key:=SHORTCUT_MAIN_SHEET_ONKEY
    On Error GoTo 0
End Sub

' skills ?V?[?g A ?? 3 ?s???~???????o?[???????W???A?d???????E?????\?[?g?????z??i1..n?j????? n ?????Bskills ???????^0 ???? n=0?B
Private Sub MasterCollectSortedSkillsMembers(ByVal wb As Workbook, ByRef outNames() As String, ByRef n As Long)
    Dim ws As Worksheet
    Dim dict As Object
    Dim r As Long, lr As Long
    Dim v As String
    Dim k As Variant
    Dim i As Long, j As Long
    Dim tmp As String
    
    n = 0
    Erase outNames
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_SKILLS_DATA)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    Set dict = CreateObject("Scripting.Dictionary")
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr < 3 Then Exit Sub
    
    For r = 3 To lr
        v = Trim$(CStr(ws.Cells(r, 1).Value))
        If Len(v) = 0 Then GoTo NextSkillsRow
        If Not dict.Exists(v) Then dict.Add v, True
NextSkillsRow:
    Next r
    
    If dict.Count = 0 Then Exit Sub
    
    ReDim outNames(1 To dict.Count)
    i = 1
    For Each k In dict.Keys
        outNames(i) = CStr(k)
        i = i + 1
    Next k
    
    For i = 1 To UBound(outNames) - 1
        For j = i + 1 To UBound(outNames)
            If StrComp(outNames(i), outNames(j), vbTextCompare) > 0 Then
                tmp = outNames(i)
                outNames(i) = outNames(j)
                outNames(j) = tmp
            End If
        Next j
    Next i
    
    n = UBound(outNames)
End Sub

' ?Éü??V?[?g?i?????o?[???j?????[?J?? HYPERLINK ?????B?V?[?g????P???p?????d???A?\???????? " ??G?X?P?[?v?B
Private Function MasterAttendanceLinkFormula(ByVal sheetName As String) As String
    Dim esc As String
    Dim dispEsc As String
    esc = Replace(sheetName, "'", "''")
    dispEsc = Replace(sheetName, """", """""")
    MasterAttendanceLinkFormula = "=HYPERLINK(""#'" & esc & "'!A1"",""" & dispEsc & """)"
End Function

' ???C?? MAIN_ATTENDANCE_LINK_COL ??Éü?????N???X?g???????????B???s???????Ñëo???????????p???B
' D ?? Clear ??t?H???g???????????AClear ?O???\?Z?????·u?E?T?C?Y???????A?????????????????B
Private Sub MasterRefreshMainAttendanceLinkList(ByVal wb As Workbook)
    Dim wsMain As Worksheet
    Dim names() As String
    Dim n As Long
    Dim i As Long
    Dim r As Long
    Dim shTest As Worksheet
    Dim clearRng As Range
    Dim preserveFontName As String
    Dim preserveFontSize As Variant
    Dim samp As Range
    Dim lrD As Long
    Dim lrA As Long
    Dim lastD As Long
    Dim fontCol As Range
    
    On Error GoTo Fail
    
    Set wsMain = MasterGetMainWorksheet(wb)
    If wsMain Is Nothing Then GoTo Fail
    
    preserveFontName = ""
    preserveFontSize = Empty
    On Error Resume Next
    lrD = wsMain.Cells(wsMain.Rows.Count, MAIN_ATTENDANCE_LINK_COL).End(xlUp).Row
    If lrD >= 2 Then
        Set samp = wsMain.Cells(2, MAIN_ATTENDANCE_LINK_COL)
    ElseIf lrD >= 1 And Len(Trim$(CStr(wsMain.Cells(1, MAIN_ATTENDANCE_LINK_COL).Value))) > 0 Then
        Set samp = wsMain.Cells(1, MAIN_ATTENDANCE_LINK_COL)
    Else
        Set samp = Nothing
    End If
    If samp Is Nothing Then
        lrA = wsMain.Cells(wsMain.Rows.Count, 1).End(xlUp).Row
        If lrA >= 2 Then
            Set samp = wsMain.Cells(2, 1)
        Else
            Set samp = wsMain.Cells(1, 1)
        End If
    End If
    If Not samp Is Nothing Then
        preserveFontName = Trim$(CStr(samp.Font.Name))
        If Len(preserveFontName) = 0 Then preserveFontName = Trim$(CStr(samp.Font.NameFarEast))
        preserveFontSize = samp.Font.Size
    End If
    Err.Clear
    On Error GoTo Fail
    
    Set clearRng = wsMain.Range(wsMain.Cells(1, MAIN_ATTENDANCE_LINK_COL), wsMain.Cells(MAIN_ATTENDANCE_LINK_CLEAR_MAX_ROWS, MAIN_ATTENDANCE_LINK_COL))
    clearRng.Clear
    
    wsMain.Cells(1, MAIN_ATTENDANCE_LINK_COL).Value = "?Éü?????N"
    wsMain.Cells(1, MAIN_ATTENDANCE_LINK_COL).Font.Bold = True
    
    Call MasterCollectSortedSkillsMembers(wb, names, n)
    If n = 0 Then
        wsMain.Cells(MAIN_ATTENDANCE_LINK_FIRST_DATA_ROW, MAIN_ATTENDANCE_LINK_COL).Value = "?iskills ?? A ??i3 ?s??`?j??????o?[??????????j"
        GoTo FitCol
    End If
    
    r = MAIN_ATTENDANCE_LINK_FIRST_DATA_ROW
    For i = 1 To n
        Set shTest = Nothing
        On Error Resume Next
        Set shTest = wb.Worksheets(names(i))
        On Error GoTo Fail
        If Not shTest Is Nothing Then
            wsMain.Cells(r, MAIN_ATTENDANCE_LINK_COL).Formula = MasterAttendanceLinkFormula(names(i))
        Else
            wsMain.Cells(r, MAIN_ATTENDANCE_LINK_COL).Value = names(i) & "?i?Éü?V?[?g?????j"
        End If
        r = r + 1
    Next i

FitCol:
    On Error Resume Next
    lastD = wsMain.Cells(wsMain.Rows.Count, MAIN_ATTENDANCE_LINK_COL).End(xlUp).Row
    If lastD < 1 Then lastD = 1
    If Len(preserveFontName) > 0 Then
        Set fontCol = wsMain.Range(wsMain.Cells(1, MAIN_ATTENDANCE_LINK_COL), wsMain.Cells(lastD, MAIN_ATTENDANCE_LINK_COL))
        If AssignFontNameToRange(fontCol, preserveFontName) Then
            If Not IsEmpty(preserveFontSize) Then
                If IsNumeric(preserveFontSize) Then
                    If CDbl(preserveFontSize) > 0# Then fontCol.Font.Size = CDbl(preserveFontSize)
                End If
            End If
        End If
        wsMain.Cells(1, MAIN_ATTENDANCE_LINK_COL).Font.Bold = True
    End If
    wsMain.Columns(MAIN_ATTENDANCE_LINK_COL).AutoFit
    On Error GoTo 0
    Exit Sub
Fail:
    On Error GoTo 0
End Sub

' ???C???V?[?g??Éü??i?????o?[??V?[?g?j??? HYPERLINK ???? skills ??????X?V
Sub master_???C????Éü?????N???X?g???X?V()
    On Error GoTo EH
    MasterRefreshMainAttendanceLinkList ThisWorkbook
    MsgBox "???C???V?[?g?iD ??j??Éü?????N???X?V????????B", vbInformation
    Exit Sub
EH:
    MsgBox "?Éü?????N?X?V?G???[: " & Err.Description, vbCritical
End Sub

' -----------------------------------------------------------------------
' ?????o?[?Éü?i?o?É£?j?V?[?g: ?A?N?e?B?u?Z???s??????\?? A?`K ??F??n?C???C?g
' ?? ThisWorkbook ???W???[???????1????????????????i?\??t????A?C?x???g???L?????????j?B
'
' Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
'     On Error Resume Next
'     MasterAttendanceRowHighlight_OnSelection Sh, Target
'     On Error GoTo 0
' End Sub
' -----------------------------------------------------------------------

Private Function MasterIsMemberAttendanceSheet(ByVal ws As Worksheet) As Boolean
    Dim h As String
    MasterIsMemberAttendanceSheet = False
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    h = Trim$(CStr(ws.Cells(1, 1).Value))
    On Error GoTo 0
    If StrComp(h, MEMBER_ATT_HEADER_A1, vbTextCompare) <> 0 Then Exit Function
    MasterIsMemberAttendanceSheet = True
End Function

Private Sub MasterAttendanceRowHighlight_Clear(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim rng As Range
    If Len(mAttRowHL_SheetName) = 0 Then Exit Sub
    If mAttRowHL_DataRow < MEMBER_ATT_FIRST_DATA_ROW Then Exit Sub
    On Error Resume Next
    Set ws = wb.Worksheets(mAttRowHL_SheetName)
    On Error GoTo 0
    If ws Is Nothing Then GoTo ResetState
    If Not MasterIsMemberAttendanceSheet(ws) Then GoTo ResetState
    On Error Resume Next
    Set rng = ws.Range(ws.Cells(mAttRowHL_DataRow, 1), ws.Cells(mAttRowHL_DataRow, MEMBER_ATT_LAST_INPUT_COL))
    rng.Interior.ColorIndex = xlNone
    On Error GoTo 0
ResetState:
    mAttRowHL_SheetName = vbNullString
    mAttRowHL_DataRow = 0
End Sub

' ThisWorkbook.SheetSelectionChange ?????Ñëo???i?W?????W???[???? Public ?K?{?j
Public Sub MasterAttendanceRowHighlight_OnSelection(ByVal Sh As Object, ByVal Target As Range)
    Dim ws As Worksheet
    Dim r As Long
    Dim rng As Range
    Dim wb As Workbook
    
    On Error GoTo QuietExit
    If Sh Is Nothing Then Exit Sub
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    Set ws = Sh
    Set wb = ws.Parent
    
    MasterAttendanceRowHighlight_Clear wb
    
    If Not MasterIsMemberAttendanceSheet(ws) Then Exit Sub
    r = Target.Row
    If r < MEMBER_ATT_FIRST_DATA_ROW Then Exit Sub
    
    Set rng = ws.Range(ws.Cells(r, 1), ws.Cells(r, MEMBER_ATT_LAST_INPUT_COL))
    On Error Resume Next
    rng.Interior.Color = RGB(255, 247, 112)
    On Error GoTo 0
    mAttRowHL_SheetName = ws.Name
    mAttRowHL_DataRow = r
QuietExit:
End Sub

' ???C?? A15/B15?EA12/B12 ??: ?????Z?????????Z??????l???èÔ?i?????E????????????h???j
Private Function MasterMainMergedTopLeftValue(ByVal ws As Worksheet, ByVal cellAddr As String) As Variant
    Dim rng As Range
    On Error GoTo Fail
    Set rng = ws.Range(cellAddr)
    MasterMainMergedTopLeftValue = rng.MergeArea.Cells(1, 1).Value
    Exit Function
Fail:
    MasterMainMergedTopLeftValue = Empty
End Function

' Excel ?Z????????iDouble ????t?V???A???EDate ?^?E??????j????u???v0?`23 ???èÔ?B????s?\?? -1?B
' ???u???????v??Z???? Value ?????l???? IsDate(?l) ?? False ????????AIsNumeric ?o?H???K?{?B
Private Function MasterMainTimeCellToHour(ByVal v As Variant) As Long
    Dim d As Date
    On Error GoTo Fail
    MasterMainTimeCellToHour = -1
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbError Then Exit Function
    
    Select Case VarType(v)
    Case vbDate
        MasterMainTimeCellToHour = Hour(CDate(v))
        Exit Function
    Case vbDouble, vbSingle, vbCurrency, vbLong, vbInteger
        d = CDate(CDbl(v))
        MasterMainTimeCellToHour = Hour(d)
        Exit Function
    Case vbString
        If Len(Trim$(v)) = 0 Then Exit Function
        d = CDate(Trim$(v))
        MasterMainTimeCellToHour = Hour(d)
        Exit Function
    Case Else
        If IsDate(v) Then
            MasterMainTimeCellToHour = Hour(CDate(v))
            Exit Function
        End If
    End Select
Fail:
    MasterMainTimeCellToHour = -1
End Function

' ???C????????Z???? Date?i?????j????????B????s?\?? False?i??????? Double ????j?B
Private Function MasterMainCellToTimeDate(ByVal v As Variant, ByRef outT As Date) As Boolean
    On Error GoTo Fail
    If IsEmpty(v) Or VarType(v) = vbError Then GoTo Fail
    
    Select Case VarType(v)
    Case vbDate
        outT = CDate(v)
        MasterMainCellToTimeDate = True
        Exit Function
    Case vbDouble, vbSingle, vbCurrency, vbLong, vbInteger
        outT = CDate(CDbl(v))
        MasterMainCellToTimeDate = True
        Exit Function
    Case vbString
        If Len(Trim$(v)) = 0 Then GoTo Fail
        outT = CDate(Trim$(v))
        MasterMainCellToTimeDate = True
        Exit Function
    Case Else
        If IsDate(v) Then
            outT = CDate(v)
            MasterMainCellToTimeDate = True
            Exit Function
        End If
    End Select
Fail:
    MasterMainCellToTimeDate = False
End Function

' ???C?? A15?EB15 = ???J?n?E?I???i?J?????_?[???o?É£??u???v?o??/???E?@?B?J?????_?[?????O?X???b?g?? A ??????F?j?B??E?s?????? False?B
Private Function MasterMainReadRegularShiftTimes(ByVal wb As Workbook, ByRef tRegStart As Date, ByRef tRegEnd As Date) As Boolean
    Dim ws As Worksheet
    Dim vS As Variant, vE As Variant
    
    MasterMainReadRegularShiftTimes = False
    Set ws = MasterGetMainWorksheet(wb)
    If ws Is Nothing Then Exit Function
    vS = MasterMainMergedTopLeftValue(ws, "A15")
    vE = MasterMainMergedTopLeftValue(ws, "B15")
    If Not MasterMainCellToTimeDate(vS, tRegStart) Then Exit Function
    If Not MasterMainCellToTimeDate(vE, tRegEnd) Then Exit Function
    If TimeValue(tRegStart) >= TimeValue(tRegEnd) Then Exit Function
    MasterMainReadRegularShiftTimes = True
End Function

' ???C?? A12?EB12 ????X???b?g??J?n?E?I???u???v?i???????j???èÔ?B??E?s?????? 6?`22?B
Private Sub MasterMachineCalReadSlotHours(ByVal wb As Workbook, ByRef hFirst As Long, ByRef hLast As Long, ByRef slotsPer As Long)
    Dim ws As Worksheet
    Dim vS As Variant, vE As Variant
    Dim hhF As Long, hhL As Long
    
    hFirst = MACHINE_CAL_SLOT_HOUR_FIRST
    hLast = MACHINE_CAL_SLOT_HOUR_LAST
    slotsPer = MACHINE_CAL_SLOTS_PER_DAY
    
    Set ws = MasterGetMainWorksheet(wb)
    If ws Is Nothing Then Exit Sub
    
    vS = MasterMainMergedTopLeftValue(ws, "A12")
    vE = MasterMainMergedTopLeftValue(ws, "B12")
    
    hhF = MasterMainTimeCellToHour(vS)
    hhL = MasterMainTimeCellToHour(vE)
    
    If hhF < 0 Or hhF > 23 Or hhL < 0 Or hhL > 23 Then Exit Sub
    If hhF > hhL Then Exit Sub
    
    hFirst = hhF
    hLast = hhL
    slotsPer = hLast - hFirst + 1
End Sub
' =========================================================
' master.xlsm ?p?i?e?L?X?g?o?b?N?A?b?v?j?B?z??EPython ?N???? ???Y???_AI?z??e?X?g.xlsm ???? ???Y???_AI?z??e?X?g_xlsm_VBA.txt ???Q??B
' ???J?}?N????????? master_ ??t?^???A?z??u?b?N???????}?N??????u?b?N????s???????h???iThisWorkbook ??O?? Alt+F8 ???j?B
' ?S?V?[?g?t?H???g?????{?t?@?C???????i?z??u?b?N??? API?Bmaster ?????C???V?[?g?p AutoFit ??????j?B
' ?u?Z???J?????_?[?v?V?[?g?Q???????o?[??o?É£???????
' - 1?s??: ???N???iA1?j
' - 4?s??: ???t?iB??`AG????L???BAH??~??????j
' - A??(5?s???~): ?????o?[??
' - ??_?Z??: ??=??? / *=?N?x / ??=?????_?É§? / ?O?x=??O?N?x?E?x?e????1_?I???`???I??(???x?e14:45?`15:00?B?x?e1????s????l12:00?`12:50) / ??x=???J?n?`?x?e????1_?J?n?E???N?x?i??????C?? A15/B15?A???? 8:45?`17:00?j
' - ?Z???F???? or ???F???u?H???????v??????i???????L?????????????Z????L????D??j
' - ?????t???????? Interior ???????E?????????A???O??\???F?????????????‡}?[????
' =========================================================
Sub master_?J?????_?[???????o?[?o?É£???()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim recMap As Object            ' key: member, value: Collection(?z??)
    Dim calendarCount As Long
    Dim memberCount As Long
    Dim totalRec As Long
    Dim key As Variant
    Dim regSt As Date, regEn As Date
    Dim regFromMain As Boolean
    
    On Error GoTo EH
    Set wb = ThisWorkbook
    Set recMap = CreateObject("Scripting.Dictionary")
    Application.ScreenUpdating = False
    
    regFromMain = MasterMainReadRegularShiftTimes(wb, regSt, regEn)
    If Not regFromMain Then
        MsgBox "??????i???C???V?[?g A15/B15?j??????????B?????o?[??o?É£???????????B", vbCritical
        GoTo CleanExit
    End If
    
    For Each ws In wb.Worksheets
        If IsMonthlyCalendarSheet(ws) Then
            calendarCount = calendarCount + 1
            CollectAttendanceFromCalendar ws, recMap, totalRec, regSt, regEn
        End If
    Next ws
    
    If calendarCount = 0 Then
        MsgBox "?u?Z???J?????_?[?v?V?[?g??????????????????B", vbExclamation
        GoTo CleanExit
    End If
    
    For Each key In recMap.Keys
        WriteMemberAttendanceSheet CStr(key), recMap(key)
        memberCount = memberCount + 1
    Next key
    
    Call MasterRefreshMainAttendanceLinkList(wb)
    
    Call AutoFitAllWorksheetColumns
    
    Call TryAutoSaveMasterWorkbook
    
    MsgBox "?????o?[??o?É£???????????B" & vbCrLf & _
           "?J?????_?[?V?[?g: " & CStr(calendarCount) & " / ?????o?[: " & CStr(memberCount) & " / ???R?[?h: " & CStr(totalRec) & vbCrLf & _
           "???i???E???E??x??o??E?O?x??I???j: " & Format$(TimeValue(regSt), "hh:nn") & " ?` " & Format$(TimeValue(regEn), "hh:nn") & _
           IIf(regFromMain, "?i???C?? A15/B15?j", "?i???C??????????? 8:45?`17:00?j"), vbInformation
    
CleanExit:
    Application.ScreenUpdating = True
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "?o?É£????G???[: " & Err.Description, vbCritical
End Sub

Private Function IsMonthlyCalendarSheet(ByVal ws As Worksheet) As Boolean
    Dim nm As String
    nm = Trim$(ws.Name)
    If InStr(nm, "?J?????_?[") = 0 Then Exit Function
    If Left$(nm, 3) = "????_" Then Exit Function
    If LCase$(nm) = "skills" Or LCase$(nm) = "need" Or LCase$(nm) = "tasks" Then Exit Function
    If StrComp(nm, SHEET_MACHINE_CALENDAR, vbTextCompare) = 0 Then Exit Function
    IsMonthlyCalendarSheet = True
End Function

' ?????t????????u?????????v?h????Z??????f???A?????V?[?g???????t??????????????????
Private Sub FlattenConditionalFormattingToInterior(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim rng As Range
    Dim c As Range
    Dim topRow As Long
    Dim pat As Long
    
    topRow = 4
    If lastRow < topRow Then Exit Sub
    
    Set rng = ws.Range(ws.Cells(topRow, 2), ws.Cells(lastRow, lastCol))
    
    On Error Resume Next
    For Each c In rng.Cells
        With c.Interior
            pat = c.DisplayFormat.Interior.Pattern
            .Pattern = pat
            If pat = xlNone Or pat = xlPatternNone Then
                .ColorIndex = xlNone
            Else
                .Color = c.DisplayFormat.Interior.Color
                .TintAndShade = c.DisplayFormat.Interior.TintAndShade
                .PatternColor = c.DisplayFormat.Interior.PatternColor
                .PatternTintAndShade = c.DisplayFormat.Interior.PatternTintAndShade
            End If
        End With
        Err.Clear
    Next c
    On Error GoTo 0
    
    On Error Resume Next
    ws.Cells.FormatConditions.Delete
    On Error GoTo 0
End Sub

Private Sub CollectAttendanceFromCalendar(ByVal wsCal As Worksheet, ByRef recMap As Object, ByRef totalRec As Long, ByVal regSt As Date, ByVal regEn As Date)
    Dim y As Long, m As Long
    Dim lastCol As Long, lastRow As Long
    Dim r As Long, c As Long
    Dim mem As String
    Dim d As Date
    Dim marker As String
    Dim cell As Range
    Dim recs As Collection
    Dim rec As Variant
    Dim isWorkingDay As Boolean
    Dim st As Variant, en As Variant
    Dim b1s As Variant, b1e As Variant, b2s As Variant, b2e As Variant
    Dim note As String
    ' ???s??????x?e????1?i?????o?[?o?É£??x?e????1_?J?n/?I?????v?B??X???? Case "" / ?? / Else ??O?x?E??x?????j
    Dim defB1s As Date, defB1e As Date
    defB1s = TimeSerial(12, 0, 0)
    defB1e = TimeSerial(12, 50, 0)
    
    If Not ResolveCalendarYearMonth(wsCal, y, m) Then Exit Sub
    
    lastCol = wsCal.Cells(4, wsCal.Columns.Count).End(xlToLeft).Column
    If lastCol > CALENDAR_LAST_DATA_COL Then lastCol = CALENDAR_LAST_DATA_COL
    lastRow = wsCal.Cells(wsCal.Rows.Count, 1).End(xlUp).Row
    If lastCol < 2 Or lastRow < 5 Then Exit Sub
    
    FlattenConditionalFormattingToInterior wsCal, lastRow, lastCol
    
    For r = 5 To lastRow
        mem = Trim$(CStr(wsCal.Cells(r, 1).Value))
        If Len(mem) = 0 Then GoTo NextMemberRow
        
        If Not recMap.Exists(mem) Then
            Set recMap(mem) = New Collection
        End If
        Set recs = recMap(mem)
        
        For c = 2 To lastCol
            If Not TryResolveCalendarDate(wsCal.Cells(4, c).Value, y, m, d) Then GoTo NextDayCell
            
            Set cell = wsCal.Cells(r, c)
            marker = Trim$(CStr(cell.Value))
            isWorkingDay = IsFactoryWorkingCell(cell)
            
            st = Empty: en = Empty
            b1s = Empty: b1e = Empty: b2s = Empty: b2e = Empty
            note = ""
            
            ' ?H??x???F????A?O?x/??x/??/?? ????L?????????????????L????D??i????u?x?v?j
            If Not isWorkingDay And Len(marker) = 0 Then
                note = "?x"
            Else
                Select Case marker
                    Case "", "???"
                        st = regSt: en = regEn
                        b1s = defB1s: b1e = defB1e
                        b2s = TimeSerial(14, 45, 0): b2e = TimeSerial(15, 0, 0)
                        note = "???"
                    Case "*", "??"
                        note = "?N?x"
                    Case "??"
                        st = regSt: en = regEn
                        b1s = defB1s: b1e = defB1e
                        b2s = TimeSerial(14, 45, 0): b2e = TimeSerial(15, 0, 0)
                        note = "?????_?É§?"
                    Case "?O?x"
                        ' ??O??N?x?????B?x?e????1_?I???`???I???B???x?e???????? 14:45?`15:00?i???H?ÑÑ?o???j
                        st = defB1e: en = regEn
                        b2s = TimeSerial(14, 45, 0): b2e = TimeSerial(15, 0, 0)
                        note = "?O?x"
                    Case "??x"
                        ' ???J?n?`?x?e????1_?J?n ?É§??B????N?x????
                        st = regSt: en = defB1s
                        note = "??x"
                    Case Else
                        st = regSt: en = regEn
                        b1s = defB1s: b1e = defB1e
                        b2s = TimeSerial(14, 45, 0): b2e = TimeSerial(15, 0, 0)
                        note = marker
                End Select
            End If
            
            ' rec(8) = ?x????Arec(9) = ???l?i?J?????_?[???????????B?????????????V?[?g J?`K ????t?L?[??????EK ?????/???l/?????????????j
            rec = Array(d, st, en, b1s, b1e, b2s, b2e, 1#, note, "")
            recs.Add rec
            totalRec = totalRec + 1
            
NextDayCell:
        Next c
        
NextMemberRow:
    Next r
End Sub

' ?????o?[?o?É£????p: ?Z??????[?U?[?????????l???????i??E?G???[?l?? False?B???l??c?????E?????????? True?j
Private Function MemberAttendanceCellHasUserValue(ByVal v As Variant) As Boolean
    On Error GoTo Quiet
    MemberAttendanceCellHasUserValue = False
    If IsEmpty(v) Then Exit Function
    If VarType(v) = vbError Then Exit Function
    If VarType(v) = vbString Then
        MemberAttendanceCellHasUserValue = (Len(Trim$(v)) > 0)
        Exit Function
    End If
    MemberAttendanceCellHasUserValue = True
    Exit Function
Quiet:
End Function

' ??????O??A?????o?[?o?É£??????? J?`K?i???l?E?c??I??j????t?L?[????????i???????????j
Private Sub PreserveMemberAttendanceHandColumns(ByVal ws As Worksheet, ByVal notesMap As Object, ByVal otEndMap As Object)
    Dim lr As Long
    Dim rr As Long
    Dim d As Variant
    Dim note As String
    Dim hdr As String
    Dim ke As Variant
    Dim dk As String
    
    On Error Resume Next
    hdr = Trim$(CStr(ws.Cells(1, 1).Value))
    On Error GoTo 0
    If StrComp(hdr, "???t", vbTextCompare) <> 0 Then Exit Sub
    
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lr < 2 Then Exit Sub
    
    For rr = 2 To lr
        d = ws.Cells(rr, 1).Value
        If Not IsDate(d) Then GoTo NextPreserveRow
        dk = Format$(CDate(d), "yyyy-mm-dd")
        note = Trim$(CStr(ws.Cells(rr, 10).Value))
        If Len(note) > 0 Then notesMap(dk) = ws.Cells(rr, 10).Value
        ke = ws.Cells(rr, 11).Value
        If Not MemberAttendanceCellHasUserValue(ke) Then ke = ws.Cells(rr, 12).Value
        If MemberAttendanceCellHasUserValue(ke) Then otEndMap(dk) = ke
NextPreserveRow:
    Next rr
End Sub

' ?????o?[?o?É£?: ?r???E???o???F?E???s??E?B???h?E?g????i?X?N???[??????1?s???\???j
Private Sub FormatMemberAttendanceSheetLayout(ByVal ws As Worksheet, ByVal lastRow As Long)
    Dim dataRng As Range
    Dim prevWs As Worksheet
    If lastRow < 1 Then Exit Sub
    Set dataRng = ws.Range("A1:K" & CStr(lastRow))
    On Error Resume Next
    With dataRng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(160, 160, 160)
    End With
    With ws.Range("A1:K1")
        .Interior.Color = RGB(180, 220, 210)
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
    End With
    On Error GoTo 0
    
    Set prevWs = Nothing
    On Error Resume Next
    Set prevWs = ActiveSheet
    On Error GoTo 0
    ' ??Ñëo?????i?Éü?????????j?? ScreenUpdating=False ?????? True ???????
    On Error Resume Next
    If Not ActiveWorkbook Is ThisWorkbook Then ThisWorkbook.Activate
    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .Split = False
    End With
    DoEvents
    ' ?@????????? ?AA1 ??\????u?????Z?b?g?i??????????????1?s?????????????????Á∑???????j
    ws.Range("A1").Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Application.GoTo ws.Range("A1"), True
    DoEvents
    ' ?B???o??1?s?????iA2?j???A?N?e?B?u??????????
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
    If Not prevWs Is Nothing Then
        On Error Resume Next
        prevWs.Activate
        On Error GoTo 0
    End If
End Sub

' ???{??????? NumberFormatLocal = "General" ?? 1004 ???Á∑?????????A?W???\???? NumberFormat ???g??
Private Sub MemberAttApplyNumberFormat(ByVal tgt As Range, ByVal fmtLocal As String)
    If StrComp(Trim$(fmtLocal), "General", vbTextCompare) = 0 Then
        tgt.NumberFormat = "General"
    Else
        tgt.NumberFormatLocal = fmtLocal
    End If
End Sub

Private Sub WriteMemberAttendanceSheet(ByVal memberName As String, ByVal recs As Collection)
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim i As Long
    Dim r As Long
    Dim rec As Variant
    Dim sortRange As Range
    Dim preservedNotes As Object
    Dim preservedOtEnd As Object
    Dim dk As String
    Dim noteOut As Variant
    Dim fmtR As Long
    Dim kv As Variant
    ' ws.Cells.Clear ???????????????A??????O??t?H???g?i???E?T?C?Y?j????????
    Dim preserveFontName As String
    Dim preserveFontSize As Variant
    Dim samp As Range
    Dim lr0 As Long
    Dim fontApply As Range
    
    preserveFontName = ""
    preserveFontSize = Empty
    
    Set wb = ThisWorkbook
    Set preservedNotes = CreateObject("Scripting.Dictionary")
    Set preservedOtEnd = CreateObject("Scripting.Dictionary")
    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets(memberName)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        Call PreserveMemberAttendanceHandColumns(ws, preservedNotes, preservedOtEnd)
        On Error Resume Next
        lr0 = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lr0 >= 2 Then
            Set samp = ws.Cells(2, 1)
        ElseIf lr0 >= 1 Then
            Set samp = ws.Cells(1, 1)
        Else
            Set samp = Nothing
        End If
        If Not samp Is Nothing Then
            preserveFontName = Trim$(CStr(samp.Font.Name))
            If Len(preserveFontName) = 0 Then preserveFontName = Trim$(CStr(samp.Font.NameFarEast))
            preserveFontSize = samp.Font.Size
        End If
        Err.Clear
        On Error GoTo 0
    End If
    
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = memberName
    End If
    
    ws.Cells.Clear
    
    ws.Cells(1, 1).Value = "???t"
    ws.Cells(1, 2).Value = "?o?????"
    ws.Cells(1, 3).Value = "??????"
    ws.Cells(1, 4).Value = "?x?e????1_?J?n"
    ws.Cells(1, 5).Value = "?x?e????1_?I??"
    ws.Cells(1, 6).Value = "?x?e????2_?J?n"
    ws.Cells(1, 7).Value = "?x?e????2_?I??"
    ws.Cells(1, 8).Value = "??????"
    ws.Cells(1, 9).Value = "?x???"
    ws.Cells(1, 10).Value = "???l"
    ws.Cells(1, 11).Value = "?c??I??"
    
    r = 2
    For i = 1 To recs.Count
        rec = recs(i)
        ws.Cells(r, 1).Value = rec(0)
        If Not IsEmpty(rec(1)) Then ws.Cells(r, 2).Value = rec(1)
        If Not IsEmpty(rec(2)) Then ws.Cells(r, 3).Value = rec(2)
        If Not IsEmpty(rec(3)) Then ws.Cells(r, 4).Value = rec(3)
        If Not IsEmpty(rec(4)) Then ws.Cells(r, 5).Value = rec(4)
        If Not IsEmpty(rec(5)) Then ws.Cells(r, 6).Value = rec(5)
        If Not IsEmpty(rec(6)) Then ws.Cells(r, 7).Value = rec(6)
        ws.Cells(r, 8).Value = rec(7)
        ws.Cells(r, 9).Value = rec(8)  ' ?x????i?N?x/?O?x/??x/?????_?É§? ???j
        dk = Format$(CDate(rec(0)), "yyyy-mm-dd")
        If preservedNotes.Exists(dk) Then
            noteOut = preservedNotes(dk)
        Else
            noteOut = rec(9)
        End If
        ws.Cells(r, 10).Value = noteOut ' ???l?i?????E???L?B??????O?????????????e???????t??????j
        If preservedOtEnd.Exists(dk) Then ws.Cells(r, 11).Value = preservedOtEnd(dk)
        r = r + 1
    Next i
    
    If r > 2 Then
        Set sortRange = ws.Range("A1:K" & CStr(r - 1))
        sortRange.Sort Key1:=ws.Range("A2"), Order1:=xlAscending, Header:=xlYes
    End If
    
    Call MemberAttApplyNumberFormat(ws.Columns("A"), "yyyy/mm/dd")
    Call MemberAttApplyNumberFormat(ws.Columns("B:G"), "hh:mm")
    ' K ???u?c??I??i?????j?v??u?c?????i???l?j?v???????????????A????^???????\???`????t????
    Call MemberAttApplyNumberFormat(ws.Columns("K"), "General")
    If r > 2 Then
        For fmtR = 2 To r - 1
            If Not IsEmpty(ws.Cells(fmtR, 11).Value) Then
                kv = ws.Cells(fmtR, 11).Value
                If VarType(kv) = vbString Then
                    Call MemberAttApplyNumberFormat(ws.Cells(fmtR, 11), "General")
                ElseIf VarType(kv) = vbDate Then
                    Call MemberAttApplyNumberFormat(ws.Cells(fmtR, 11), "hh:mm")
                ElseIf IsNumeric(kv) Then
                    If CDbl(kv) >= 0# And CDbl(kv) < 1# Then
                        Call MemberAttApplyNumberFormat(ws.Cells(fmtR, 11), "hh:mm")
                    Else
                        Call MemberAttApplyNumberFormat(ws.Cells(fmtR, 11), "General")
                    End If
                ElseIf IsDate(kv) Then
                    Call MemberAttApplyNumberFormat(ws.Cells(fmtR, 11), "hh:mm")
                End If
            End If
        Next fmtR
    End If
    ws.Columns("A:K").AutoFit
    
    Call FormatMemberAttendanceSheetLayout(ws, r - 1)
    
    ' Clear ??????t?H???g??????A??????O?????????t?H???g?? A1:K ??\??????
    If Len(preserveFontName) > 0 And (r - 1) >= 1 Then
        On Error Resume Next
        Set fontApply = ws.Range("A1:K" & CStr(r - 1))
        If AssignFontNameToRange(fontApply, preserveFontName) Then
            If Not IsEmpty(preserveFontSize) Then
                If IsNumeric(preserveFontSize) Then
                    If CDbl(preserveFontSize) > 0# Then fontApply.Font.Size = CDbl(preserveFontSize)
                End If
            End If
        End If
        With ws.Range("A1:K1")
            .Font.Bold = True
            .Font.Color = RGB(0, 0, 0)
        End With
        Err.Clear
        On Error GoTo 0
    End If
End Sub

' =========================================================
' ?@?B?J?????_?[: 1 ?V?[?g?Askills ????^?? 1?`2 ?s???o???AA ?????t?????i1 ????P???X???b?g?j
' =========================================================
Public Sub master_?@?B?J?????_?[????()
    Dim wb As Workbook
    Dim wsSkill As Worksheet
    Dim wsOut As Worksheet
    Dim preserved As Object
    Dim startDay As Date
    Dim lastC As Long
    Dim useTwoHeader As Boolean
    Dim c As Long
    Dim r As Long
    Dim dayOffset As Long
    Dim procH As String
    Dim machH As String
    Dim colKey As String
    Dim mapKey As String
    Dim lastRow As Long
    Dim slotHour As Long
    Dim slotTime As Date
    Dim outCol As Long
    Dim lastOutCol As Long
    Dim hFirst As Long, hLast As Long, slotsPer As Long
    
    On Error GoTo EH
    Set wb = ThisWorkbook
    On Error Resume Next
    Set wsSkill = wb.Worksheets(SHEET_SKILLS_DATA)
    On Error GoTo EH
    If wsSkill Is Nothing Then
        MsgBox "skills ?V?[?g??????????B?H?????E?@?B??????o?????R?s?[????????B", vbExclamation, "?@?B?J?????_?["
        Exit Sub
    End If
    
    Call MachineCalGetSkillsMode(wsSkill, lastC, useTwoHeader)
    If lastC < 2 Then
        MsgBox "skills ?V?[?g??L????????????B", vbExclamation, "?@?B?J?????_?["
        Exit Sub
    End If
    
    Set wsOut = EnsureMachineCalendarSheet(wb, wsSkill)
    If wsOut Is Nothing Then
        MsgBox "?u" & SHEET_MACHINE_CALENDAR & "?v?V?[?g??p?????????B", vbExclamation, "?@?B?J?????_?["
        Exit Sub
    End If
    
    Set preserved = CreateObject("Scripting.Dictionary")
    Call PreserveMachineCalendarGrid(wsOut, preserved)
    
    Call MasterMachineCalReadSlotHours(wb, hFirst, hLast, slotsPer)
    
    Application.ScreenUpdating = False
    
    wsOut.Cells.UnMerge
    wsOut.Cells.Clear
    
    With wsOut.Range("A1:A2")
        .Merge
        .Value = "???t????"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With wsOut.Range("B1:B2")
        .Merge
        .Value = "?j??"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    lastOutCol = lastC + 1
    For c = 2 To lastC
        If Not MachineCalSkillsColumnUsable(wsSkill, c, useTwoHeader) Then GoTo NextHdr
        outCol = c + 1
        wsOut.Cells(1, outCol).Value = wsSkill.Cells(1, c).Value
        If useTwoHeader Then
            wsOut.Cells(2, outCol).Value = wsSkill.Cells(2, c).Value
        Else
            wsOut.Cells(2, outCol).ClearContents
        End If
NextHdr:
    Next c
    
    startDay = DateSerial(Year(Date), Month(Date), Day(Date))
    lastRow = 2 + MACHINE_CAL_HORIZON_DAYS * slotsPer
    
    r = 3
    For dayOffset = 0 To MACHINE_CAL_HORIZON_DAYS - 1
        For slotHour = hFirst To hLast
            slotTime = (startDay + dayOffset) + TimeSerial(slotHour, 0, 0)
            wsOut.Cells(r, 1).Value = slotTime
            wsOut.Cells(r, 2).Value = MachineCalWeekdayShort(slotTime)
            For c = 2 To lastC
                If Not MachineCalSkillsColumnUsable(wsSkill, c, useTwoHeader) Then GoTo NextCell
                outCol = c + 1
                procH = Trim$(CStr(wsOut.Cells(1, outCol).Value))
                machH = Trim$(CStr(wsOut.Cells(2, outCol).Value))
                colKey = procH & vbTab & machH
                mapKey = Format$(slotTime, "yyyy-mm-dd hh:nn") & "|" & colKey
                If preserved.Exists(mapKey) Then
                    wsOut.Cells(r, outCol).Value = preserved(mapKey)
                Else
                    wsOut.Cells(r, outCol).ClearContents
                End If
NextCell:
            Next c
            r = r + 1
        Next slotHour
    Next dayOffset
    
    Call FormatMachineCalendarGridLayout(wsOut, lastRow, lastOutCol, slotsPer, wb)
    
    Application.ScreenUpdating = True
    Call AutoFitAllWorksheetColumns
    Call TryAutoSaveMasterWorkbook
    
    MsgBox "?u" & SHEET_MACHINE_CALENDAR & "?v???X?V????????B" & vbCrLf & _
           "?J?n??: " & Format$(startDay, "yyyy/mm/dd") & "?i?{???j / ????: " & CStr(MACHINE_CAL_HORIZON_DAYS) & " ??" & vbCrLf & _
           "A ??????EB ??j???A1 ????P??i" & CStr(hFirst) & ":00?`" & CStr(hLast) & ":00?E1 ?? " & CStr(slotsPer) & " ?s?j?E???X?g???C?v?B" & vbCrLf & _
           "?i????X???b?g: ???C?? A12/B12?B???? 6?`22 ???j" & vbCrLf & _
           "?i???O??F: ???C?? A15/B15?EA ????B?????F????j" & vbCrLf & _
           "??????????????????A????????E????H???E?@?B??Z?????????????????????????B", vbInformation, "?@?B?J?????_?["
    Exit Sub
    
EH:
    Application.ScreenUpdating = True
    MsgBox "?@?B?J?????_?[???G???[: " & Err.Description, vbCritical, "?@?B?J?????_?["
End Sub

Sub master_?A?j???t??_?@?B?J?????_?[????()
    Call AnimateButtonPush
    Call master_?@?B?J?????_?[????
End Sub

Private Function EnsureMachineCalendarSheet(ByVal wb As Workbook, ByVal wsAfter As Worksheet) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_MACHINE_CALENDAR)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wsAfter)
        On Error Resume Next
        ws.Name = SHEET_MACHINE_CALENDAR
        On Error GoTo 0
    End If
    Set EnsureMachineCalendarSheet = ws
End Function

Private Sub MachineCalGetSkillsMode(ByVal wsSkill As Worksheet, ByRef lastC As Long, ByRef useTwoHeader As Boolean)
    Dim c As Long
    Dim pHdr As String
    Dim mHdr As String
    Dim pm As Long
    
    lastC = wsSkill.Cells(1, wsSkill.Columns.Count).End(xlToLeft).Column
    pm = 0
    For c = 2 To lastC
        pHdr = Trim$(CStr(wsSkill.Cells(1, c).Value))
        mHdr = Trim$(CStr(wsSkill.Cells(2, c).Value))
        If Len(pHdr) > 0 And Len(mHdr) > 0 Then
            If StrComp(LCase$(pHdr), "nan", vbTextCompare) <> 0 And StrComp(LCase$(mHdr), "nan", vbTextCompare) <> 0 Then
                pm = pm + 1
            End If
        End If
    Next c
    useTwoHeader = (pm > 0)
End Sub

Private Function MachineCalSkillsColumnUsable(ByVal wsSkill As Worksheet, ByVal c As Long, ByVal useTwoHeader As Boolean) As Boolean
    Dim pHdr As String
    Dim mHdr As String
    
    pHdr = Trim$(CStr(wsSkill.Cells(1, c).Value))
    If useTwoHeader Then
        mHdr = Trim$(CStr(wsSkill.Cells(2, c).Value))
        MachineCalSkillsColumnUsable = (Len(pHdr) > 0 And Len(mHdr) > 0 And StrComp(LCase$(pHdr), "nan", vbTextCompare) <> 0 And StrComp(LCase$(mHdr), "nan", vbTextCompare) <> 0)
    Else
        MachineCalSkillsColumnUsable = (Len(pHdr) > 0 And StrComp(LCase$(pHdr), "nan", vbTextCompare) <> 0)
    End If
End Function

' ?j??????i?????\??????????B???{?? Excel ??? ???????c?j
Private Function MachineCalWeekdayShort(ByVal dt As Date) As String
    On Error Resume Next
    MachineCalWeekdayShort = Format$(dt, "aaa")
    If Len(Trim$(MachineCalWeekdayShort)) = 0 Then MachineCalWeekdayShort = Format$(dt, "ddd")
    On Error GoTo 0
End Function

Private Sub PreserveMachineCalendarGrid(ByVal ws As Worksheet, ByVal outMap As Object)
    Dim lastC As Long
    Dim lastR As Long
    Dim rr As Long
    Dim cc As Long
    Dim hdr As String
    Dim procH As String
    Dim machH As String
    Dim colKey As String
    Dim note As String
    Dim startEquipCol As Long
    
    On Error Resume Next
    hdr = Trim$(CStr(ws.Cells(1, 1).Value))
    On Error GoTo 0
    If StrComp(hdr, "???t", vbTextCompare) <> 0 And StrComp(hdr, "???t????", vbTextCompare) <> 0 Then Exit Sub
    
    lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 3 Then Exit Sub
    
    If StrComp(Trim$(CStr(ws.Cells(1, 2).Value)), "?j??", vbTextCompare) = 0 Then
        startEquipCol = 3
    Else
        startEquipCol = 2
    End If
    If lastC < startEquipCol Then Exit Sub
    
    For cc = startEquipCol To lastC
        procH = Trim$(CStr(ws.Cells(1, cc).Value))
        machH = Trim$(CStr(ws.Cells(2, cc).Value))
        colKey = procH & vbTab & machH
        For rr = 3 To lastR
            If Not IsDate(ws.Cells(rr, 1).Value) Then GoTo NextPreserveCell
            note = Trim$(CStr(ws.Cells(rr, cc).Value))
            If Len(note) = 0 Then GoTo NextPreserveCell
            outMap(Format$(CDate(ws.Cells(rr, 1).Value), "yyyy-mm-dd hh:nn") & "|" & colKey) = ws.Cells(rr, cc).Value
NextPreserveCell:
        Next rr
    Next cc
End Sub

Private Sub FormatMachineCalendarGridLayout(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long, ByVal slotsPerDay As Long, ByVal wb As Workbook)
    Dim dataRng As Range
    Dim prevWs As Worksheet
    Dim prevScr As Boolean
    Dim hdrTop As Long
    Dim hdrMid As Long
    Dim hdrFont As Long
    Dim memHdr As Long
    
    If lastRow < 3 Or lastCol < 1 Then Exit Sub
    
    hdrTop = RGB(47, 84, 150)
    hdrMid = RGB(70, 115, 175)
    hdrFont = RGB(255, 255, 255)
    memHdr = RGB(180, 220, 210)
    
    On Error Resume Next
    With ws.Range("A1:A2")
        .Interior.Color = memHdr
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
    End With
    With ws.Range("B1:B2")
        .Interior.Color = memHdr
        .Font.Bold = True
        .Font.Color = RGB(0, 0, 0)
    End With
    On Error GoTo 0
    
    If lastCol >= 3 Then
        With ws.Range(ws.Cells(1, 3), ws.Cells(1, lastCol))
            .Interior.Color = hdrTop
            .Font.Bold = True
            .Font.Color = hdrFont
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        With ws.Range(ws.Cells(2, 3), ws.Cells(2, lastCol))
            .Interior.Color = hdrMid
            .Font.Bold = True
            .Font.Color = hdrFont
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
    End If
    
    Set dataRng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    On Error Resume Next
    With dataRng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(160, 160, 160)
    End With
    On Error GoTo 0
    
    On Error Resume Next
    ws.Range(ws.Cells(3, 1), ws.Cells(lastRow, 1)).NumberFormatLocal = "yyyy/mm/dd hh:mm"
    On Error GoTo 0
    
    Call ApplyMachineCalendarDayStripes(ws, lastRow, lastCol, slotsPerDay)
    Call ApplyMachineCalendarOutsideRegularHoursTint(ws, wb, lastRow, lastCol)
    
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Columns.AutoFit
    
    Set prevWs = Nothing
    On Error Resume Next
    Set prevWs = ActiveSheet
    On Error GoTo 0
    prevScr = Application.ScreenUpdating
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not ActiveWorkbook Is ThisWorkbook Then ThisWorkbook.Activate
    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .Split = False
    End With
    DoEvents
    ws.Range("A1").Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Application.GoTo ws.Range("A1"), True
    DoEvents
    ws.Range("C3").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
    If Not prevWs Is Nothing Then
        On Error Resume Next
        prevWs.Activate
        On Error GoTo 0
    End If
    Application.ScreenUpdating = prevScr
End Sub

' 3 ?s???~???u?J?????_?[???v?P?????????X?g???C?v?i?????????F?j
Private Sub ApplyMachineCalendarDayStripes(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long, ByVal slotsPerDay As Long)
    Dim rr As Long
    Dim dayIx As Long
    Dim fillColor As Long
    Dim bandA As Long
    Dim bandB As Long
    Dim spd As Long
    
    If lastRow < 3 Or lastCol < 1 Then Exit Sub
    spd = slotsPerDay
    If spd < 1 Then spd = MACHINE_CAL_SLOTS_PER_DAY
    
    bandA = RGB(237, 246, 255)
    bandB = RGB(220, 236, 252)
    
    For rr = 3 To lastRow
        dayIx = (rr - 3) \ spd
        If (dayIx Mod 2) = 0 Then
            fillColor = bandA
        Else
            fillColor = bandB
        End If
        With ws.Range(ws.Cells(rr, 1), ws.Cells(rr, lastCol)).Interior
            .Pattern = xlSolid
            .Color = fillColor
        End With
    Next rr
End Sub

' ???J?n?`?I???i???C?? A15/B15?j??d????? 1 ????X???b?g?? A ??i?????j???????I?????W??????iB ?????X?g???C?v????j
Private Sub ApplyMachineCalendarOutsideRegularHoursTint(ByVal ws As Worksheet, ByVal wb As Workbook, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim regSt As Date, regEn As Date
    Dim rr As Long
    Dim slotFull As Date
    Dim slotStartFrac As Double, slotEndFrac As Double
    Dim regStartFrac As Double, regEndFrac As Double
    Dim useColor As Long
    
    useColor = RGB(244, 224, 198)
    If lastRow < 3 Then Exit Sub
    If Not MasterMainReadRegularShiftTimes(wb, regSt, regEn) Then Exit Sub
    
    regStartFrac = TimeValue(regSt)
    regEndFrac = TimeValue(regEn)
    
    For rr = 3 To lastRow
        On Error Resume Next
        slotFull = CDate(ws.Cells(rr, 1).Value)
        On Error GoTo 0
        If Not IsDate(ws.Cells(rr, 1).Value) Then GoTo NextMcTint
        slotStartFrac = TimeValue(slotFull)
        slotEndFrac = slotStartFrac + (1# / 24#)
        ' ???J??? [slotStart, slotEnd) ?? [regStart, regEnd) ???d????? ?? ???O
        If (slotEndFrac <= regStartFrac) Or (slotStartFrac >= regEndFrac) Then
            With ws.Cells(rr, 1).Interior
                .Pattern = xlSolid
                .Color = useColor
            End With
        End If
NextMcTint:
    Next rr
End Sub

Private Function ResolveCalendarYearMonth(ByVal wsCal As Worksheet, ByRef outY As Long, ByRef outM As Long) As Boolean
    Dim v As Variant
    Dim s As String
    Dim re As Object, ms As Object
    
    v = wsCal.Range("A1").Value
    If IsDate(v) Then
        outY = Year(CDate(v))
        outM = Month(CDate(v))
        ResolveCalendarYearMonth = True
        Exit Function
    End If
    
    s = Trim$(CStr(v))
    If Len(s) > 0 Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = False
        re.IgnoreCase = True
        re.Pattern = "(\d{4}).*?(\d{1,2})"
        If re.test(s) Then
            Set ms = re.Execute(s)
            outY = CLng(ms(0).SubMatches(0))
            outM = CLng(ms(0).SubMatches(1))
            If outM >= 1 And outM <= 12 Then
                ResolveCalendarYearMonth = True
                Exit Function
            End If
        End If
    End If
    
    ' A1 ???????????A?V?[?g????un???J?????_?[?v??????p?i?N????N?j
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True
    re.Pattern = "(\d{1,2})??"
    If re.test(wsCal.Name) Then
        Set ms = re.Execute(wsCal.Name)
        outY = Year(Date)
        outM = CLng(ms(0).SubMatches(0))
        If outM >= 1 And outM <= 12 Then
            ResolveCalendarYearMonth = True
            Exit Function
        End If
    End If
End Function

Private Function TryResolveCalendarDate(ByVal v As Variant, ByVal y As Long, ByVal m As Long, ByRef outD As Date) As Boolean
    Dim d As Long
    If IsDate(v) Then
        outD = CDate(v)
        TryResolveCalendarDate = True
        Exit Function
    End If
    If IsNumeric(v) Then
        d = CLng(v)
        If d >= 1 And d <= 31 Then
            On Error Resume Next
            outD = DateSerial(y, m, d)
            If Err.Number = 0 Then TryResolveCalendarDate = True
            Err.Clear
            On Error GoTo 0
        End If
    End If
End Function

Private Function IsFactoryWorkingCell(ByVal c As Range) As Boolean
    Dim clr As Long
    Dim rr As Long, gg As Long, bb As Long
    
    ' ?h????????????????????????
    If c.Interior.Pattern = xlPatternNone Then
        IsFactoryWorkingCell = True
        Exit Function
    End If
    
    clr = c.Interior.Color
    If clr = RGB(255, 255, 255) Then
        IsFactoryWorkingCell = True
        Exit Function
    End If
    
    rr = clr Mod 256
    gg = (clr \ 256) Mod 256
    bb = (clr \ 65536) Mod 256
    ' ???F?n?i?????????j????????????
    If rr >= 235 And gg >= 220 And bb <= 190 Then
        IsFactoryWorkingCell = True
    End If
End Function

' =========================================================
' ?}?`?{?^??: ?????A?j???[?V???? ?? ?o?É£????imaster.xlsm ?p?j
' ?o???????{?????C?? A15/B15?i???j????A?S?J?????_?[?????K?p?i???? 8:45?`17:00?j?B
' =========================================================
Sub master_?A?j???t??_?J?????_?[????o?É£???()
    Call AnimateButtonPush
    Call master_?J?????_?[???????o?[?o?É£???
End Sub

' ?}?`?{?^???p?FOnAction ??{?????w????? AnimateButtonPush ???????????????b?p?[?????‰Ì???
Sub master_?A?j???t??_?S?V?[?g?t?H???g?????X?g????I?????????()
    Call AnimateButtonPush
    Call master_?S?V?[?g?t?H???g?????X?g????I?????????
End Sub

Sub master_?A?j???t??_?S?V?[?g?t?H???g???????????()
    Call AnimateButtonPush
    Call master_?S?V?[?g?t?H???g???????????
End Sub

Sub master_?A?j???t??_?S?V?[?g?t?H???g_BIZ_UDP?S?V?b?N?????()
    Call AnimateButtonPush
    Call master_?S?V?[?g?t?H???g_BIZ_UDP?S?V?b?N?????
End Sub

Sub master_?A?j???t??_??????????n??V?[?g???X?V()
    On Error GoTo EH
    Call AnimateButtonPush
    Call MasterEnsureMachineChangeoverTemplateSheets
    MsgBox "???V?[?g???X?V????????B" & vbCrLf & _
           "?E???_??????O?áq?? ?c skills ??H??+?@?B????o?^?s????i????/??n???????????s???????j" & vbCrLf & _
           "?E???_?@?B_?????n????? ?c skills ??@?B??????o?^?s????i????????s???????j", vbInformation, "??????E?????n??"
    Exit Sub
EH:
    MsgBox "?G???[: " & Err.Description, vbCritical, "??????E?????n??"
End Sub

' =========================================================
' ????F?{?^????????????A?j???[?V?????i???Y???_AI?z??e?X?g_xlsm_VBA.txt ????l?E?S?V?[?g?? Caller ???????j
' =========================================================
Private Sub AnimateButtonPush()
    Dim shpName As String
    Dim shp As Shape
    Dim ws As Worksheet
    Dim candidate As Shape
    Dim firstHit As Shape
    Dim originalTop As Single
    Dim originalLeft As Single
    Dim hasShadow As Boolean
    
    On Error Resume Next
    shpName = CStr(Application.Caller)
    On Error GoTo 0
    If Len(Trim$(shpName)) = 0 Then Exit Sub
    
    Set shp = Nothing
    Set firstHit = Nothing
    For Each ws In ThisWorkbook.Worksheets
        Err.Clear
        On Error Resume Next
        Set candidate = ws.Shapes(shpName)
        If Err.Number = 0 And Not candidate Is Nothing Then
            If firstHit Is Nothing Then Set firstHit = candidate
            If ws Is ActiveSheet Then
                Set shp = candidate
                Exit For
            End If
        End If
    Next ws
    On Error GoTo 0
    
    If shp Is Nothing Then Set shp = firstHit
    If shp Is Nothing Then Exit Sub
    
    originalTop = shp.Top
    originalLeft = shp.Left
    hasShadow = shp.Shadow.Visible
    
    shp.Top = originalTop + 2
    shp.Left = originalLeft + 2
    If hasShadow Then shp.Shadow.Visible = msoFalse
    
    DoEvents
    Sleep 150
    
    shp.Top = originalTop
    shp.Left = originalLeft
    If hasShadow Then shp.Shadow.Visible = msoTrue
    DoEvents
End Sub

' =========================================================
' ???????????{?^???????????imaster.xlsm ????: ?o?É£???????{?^???j
' =========================================================
' ?v???Z?b?g: 1=?? 2=?e?B?[?? ?c 10=?}?[???^?iCreateCoolButtonWithPreset ?? presetId?j
Private Function CoolButtonGradientTop(ByVal presetId As Long) As Long
    Select Case presetId
        Case 1: CoolButtonGradientTop = RGB(65, 105, 225)
        Case 2: CoolButtonGradientTop = RGB(0, 180, 170)
        Case 3: CoolButtonGradientTop = RGB(255, 160, 60)
        Case 4: CoolButtonGradientTop = RGB(60, 179, 113)
        Case 5: CoolButtonGradientTop = RGB(186, 85, 211)
        Case 6: CoolButtonGradientTop = RGB(100, 120, 220)
        Case 7: CoolButtonGradientTop = RGB(130, 140, 150)
        Case 8: CoolButtonGradientTop = RGB(255, 120, 120)
        Case 9: CoolButtonGradientTop = RGB(255, 200, 80)
        Case 10: CoolButtonGradientTop = RGB(230, 90, 180)
        Case Else: CoolButtonGradientTop = RGB(65, 105, 225)
    End Select
End Function

Private Function CoolButtonGradientBottom(ByVal presetId As Long) As Long
    Select Case presetId
        Case 1: CoolButtonGradientBottom = RGB(0, 0, 139)
        Case 2: CoolButtonGradientBottom = RGB(0, 100, 95)
        Case 3: CoolButtonGradientBottom = RGB(180, 80, 0)
        Case 4: CoolButtonGradientBottom = RGB(0, 90, 40)
        Case 5: CoolButtonGradientBottom = RGB(75, 0, 130)
        Case 6: CoolButtonGradientBottom = RGB(40, 50, 120)
        Case 7: CoolButtonGradientBottom = RGB(70, 75, 85)
        Case 8: CoolButtonGradientBottom = RGB(180, 50, 50)
        Case 9: CoolButtonGradientBottom = RGB(180, 120, 0)
        Case 10: CoolButtonGradientBottom = RGB(140, 30, 100)
        Case Else: CoolButtonGradientBottom = RGB(0, 0, 139)
    End Select
End Function

Private Sub CreateCoolButtonWithPreset(btnText As String, macroName As String, posX As Single, posY As Single, ByVal presetId As Long, Optional ByVal fixedShapeName As String = vbNullString)
    CreateCoolButton btnText, macroName, posX, posY, CoolButtonGradientTop(presetId), CoolButtonGradientBottom(presetId), fixedShapeName
End Sub

' ?????V?[?g??u?o?É£????v?{?^????1??z?u?i?e?B?[???n?v???Z?b?g 2?j
Sub master_???????????{?^??????()
    Dim y As Single
    Const gap As Single = 70
    y = 50
    CreateCoolButtonWithPreset "?o?É£???", "master_?A?j???t??_?J?????_?[????o?É£???", 50, y, 2
    y = y + gap
    CreateCoolButtonWithPreset "?@?B?J?????_?[??", "master_?A?j???t??_?@?B?J?????_?[????", 50, y, 3
    y = y + gap
    On Error Resume Next
    ActiveSheet.Shapes("Btn_ChangeoverTemplateStack").Delete
    On Error GoTo 0
    CreateCoolButtonWithPreset "??????E?????n??V?[?g", "master_?A?j???t??_??????????n??V?[?g???X?V", 50, y, 4, "Btn_ChangeoverTemplateStack"
    MsgBox "?{?^??????????????B" & vbCrLf & _
           "?E?o?É£??? ?c ?J?????_?[???????o?[?o?É£????????i??????C?? A15/B15?E???? 8:45?`17:00?j" & vbCrLf & _
           "?E?@?B?J?????_?[?? ?c A?????EB?j???EC?`skills?????E???C??A12/B12?????`?i????6?`22???j?E???X?g???C?v??u?@?B?J?????_?[?v1???i365???j" & vbCrLf & _
           "?E??????E?????n??V?[?g ?c skills ????u???_??????O?áq??v?u???_?@?B_?????n??????v??H?????E?@?B??????i???????????????A?j???j" & vbCrLf & _
           "?D?????u??h???b?O????z?u????????????B", vbInformation
End Sub

' ?z?F P1?`P10 ????{?i?N???b?N????????j
Sub master_???????????{?^??_?z?F?T???v????()
    Dim i As Long
    Dim x As Single
    Dim y As Single
    Const colW As Single = 232
    Const rowH As Single = 62
    Const left0 As Single = 40
    Const top0 As Single = 40
    
    For i = 1 To 10
        x = left0 + CSng((i - 1) Mod 5) * colW
        y = top0 + CSng((i - 1) \ 5) * rowH
        CreateCoolButton "P" & CStr(i), "master_???????????{?^??????", x, y, CoolButtonGradientTop(i), CoolButtonGradientBottom(i)
        On Error Resume Next
        ActiveSheet.Shapes(ActiveSheet.Shapes.Count).OnAction = ""
        On Error GoTo 0
    Next i
    MsgBox "?z?F?v???Z?b?g P1?`P10 ????{??z?u????????B?s?v????????????????B", vbInformation
End Sub

Private Sub CreateCoolButton(btnText As String, macroName As String, posX As Single, posY As Single, colorTop As Long, colorBottom As Long, Optional ByVal fixedShapeName As String = vbNullString)
    Dim shp As Shape
    
    Set shp = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, posX, posY, 220, 50)
    
    With shp
        With .TextFrame2.TextRange
            .Text = btnText
            .Font.Name = "???C???I"
            .Font.Size = 14
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        End With
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        
        With .Fill
            .Visible = msoTrue
            .TwoColorGradient Style:=msoGradientVertical, Variant:=1
            .ForeColor.RGB = colorTop
            .BackColor.RGB = colorBottom
        End With
        
        .Line.Visible = msoFalse
        
        With .ThreeD
            .BevelTopType = msoBevelSoftRound
            .BevelTopDepth = 6
            .BevelTopInset = 6
        End With
        
        With .Shadow
            .Type = msoShadow21
            .Visible = msoTrue
            .OffsetX = 3
            .OffsetY = 3
            .Transparency = 0.5
            .Blur = 4
        End With
        
        .OnAction = macroName
        
        On Error Resume Next
        If Len(Trim$(fixedShapeName)) > 0 Then
            .Name = Trim$(fixedShapeName)
        Else
            Randomize
            .Name = "CoolBtn_" & Format(Now, "yyyymmddhhnnss") & "_" & Format(Int(1000000 * Rnd), "000000")
        End If
        On Error GoTo 0
    End With
End Sub

' =========================================================
' ?S?V?[?g??Z???t?H???g???i?ò@???s?E?{?^????????j
' ???Z???O???b?h???B?}?`?E?O???t???e?L?X?g????O?B
' ???V?[?g???: ??p?X???[?h???????????t?H???g?E???????B?p?X???[?h?t????X?L?b?v?????B
' ???u???X?g?I???v?? Excel ?W????m?Z??????????n???m?t?H???g?n?_?C?A???O???g?p?B
' ??365/??????? Dialogs.Show ?? False ??G???[???Á∑?????????A???s???? InputBox ????B
' ?????{??Excel??? Font.Name ?????????????????????????????? NameFarEast?E?e?[?}???????s???B
' ???}?`??}?N?????umaster_?A?j???t??_?S?V?[?g?t?H???g?????X?g????I?????????v???w??i?????A?j???p?j?B
' ???z??u?b?N??u???C???V?[?g_A????K??_AutoFit?v?? master ???????????????B
' =========================================================
Private Function GetOrCreateFontScratchSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SCRATCH_SHEET_FONT)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        On Error Resume Next
        ws.Name = SCRATCH_SHEET_FONT
        On Error GoTo 0
        ws.Range("A1").Value = "?i?t?H???g?I??p?E?????????????????j"
        ws.Visible = xlSheetVeryHidden
    End If
    Set GetOrCreateFontScratchSheet = ws
End Function

Private Sub RestoreCellFontProps(ByVal r As Range, ByVal oldName As String, _
    ByVal oldSize As Variant, ByVal oldBold As Variant, ByVal oldItalic As Variant, _
    ByVal oldUnderline As Variant, ByVal oldColor As Variant, ByVal oldStrike As Variant)
    On Error Resume Next
    With r.Font
        .Name = oldName
        If Not IsEmpty(oldSize) Then .Size = oldSize
        If Not IsEmpty(oldBold) Then .Bold = oldBold
        If Not IsEmpty(oldItalic) Then .Italic = oldItalic
        If Not IsEmpty(oldUnderline) Then .Underline = oldUnderline
        If Not IsEmpty(oldColor) Then .Color = oldColor
        If Not IsEmpty(oldStrike) Then .Strikethrough = oldStrike
    End With
    On Error GoTo 0
End Sub

Private Function FontNameExistsInExcel(ByVal fontName As String) As Boolean
    Dim i As Long
    For i = 1 To Application.FontNames.Count
        If StrComp(Application.FontNames(i), fontName, vbTextCompare) = 0 Then
            FontNameExistsInExcel = True
            Exit Function
        End If
    Next i
End Function

' ??p?X???[?h??V?[?g????????????i?????o?[?o?É£???j?B?p?X???[?h?t???? False ????B
Private Function TryTempUnprotectBlank(ByVal ws As Worksheet) As Boolean
    TryTempUnprotectBlank = False
    If Not ws.ProtectContents Then Exit Function
    On Error Resume Next
    ws.Unprotect Password:=""
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    If Not ws.ProtectContents Then TryTempUnprotectBlank = True
    On Error GoTo 0
End Function

' ???????????V?[?g??????i??????p?X???[?h?z??B?I?v?V?????? Excel ?????????????j
Private Sub ReprotectSheetBlankPwdUI(ByVal ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True
    Err.Clear
    On Error GoTo 0
End Sub

' ?u?b?N????e???[?N?V?[?g?? UsedRange ?????I?[?g?t?B?b?g
Private Sub AutoFitAllWorksheetColumns()
    Dim ws As Worksheet
    Dim ur As Range
    Dim unlk As Boolean
    For Each ws In ThisWorkbook.Worksheets
        unlk = TryTempUnprotectBlank(ws)
        On Error Resume Next
        If Not ws.ProtectContents Then
            Set ur = ws.UsedRange
            If Not ur Is Nothing Then ur.Columns.AutoFit
        End If
        Err.Clear
        On Error GoTo 0
        If unlk Then ReprotectSheetBlankPwdUI ws
    Next ws
End Sub

' ?t?H???g??X?E?o?É£???????AThisWorkbook?imaster.xlsm?j????????
Private Sub TryAutoSaveMasterWorkbook()
    On Error Resume Next
    If ThisWorkbook.ReadOnly Then
        On Error GoTo 0
        Exit Sub
    End If
    ThisWorkbook.Save
    If Err.Number <> 0 Then
        MsgBox "master.xlsm ????????????s???????: " & Err.Description, vbExclamation
        Err.Clear
    End If
    On Error GoTo 0
End Sub

' UsedRange ????t?H???g????K?p?i?e?[?}?t?H???g?E???A?W?A??t?H???g??????O???????j
Private Function AssignFontNameToRange(ByVal tgt As Range, ByVal fontName As String) As Boolean
    On Error Resume Next
    tgt.Font.ThemeFont = xlThemeFontNone
    Err.Clear
    tgt.Font.Name = fontName
    If Err.Number <> 0 Then
        AssignFontNameToRange = False
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    tgt.Font.NameFarEast = fontName
    Err.Clear
    tgt.Font.NameAscii = fontName
    Err.Clear
    AssignFontNameToRange = True
    On Error GoTo 0
End Function

Private Sub ApplyFontToAllSheetCells(ByVal fontName As String, ByRef skippedOut As String)
    Dim ws As Worksheet
    Dim ur As Range
    Dim rangeErr As Boolean
    Dim prevCalc As XlCalculation
    Dim prevScr As Boolean
    Dim prevEv As Boolean
    Dim wasProt As Boolean
    Dim unprotectOk As Boolean
    
    skippedOut = ""
    prevCalc = Application.Calculation
    prevScr = Application.ScreenUpdating
    prevEv = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    For Each ws In ThisWorkbook.Worksheets
        wasProt = False
        unprotectOk = True
        On Error Resume Next
        If ws.ProtectContents Then
            wasProt = True
            ws.Unprotect
            If ws.ProtectContents And Len(SHEET_FONT_UNPROTECT_PASSWORD) > 0 Then
                ws.Unprotect Password:=SHEET_FONT_UNPROTECT_PASSWORD
            End If
            If ws.ProtectContents Then
                unprotectOk = False
                skippedOut = skippedOut & "?E" & ws.Name & "?i????????????X?L?b?v?j" & vbCrLf
                Err.Clear
            End If
        End If
        
        If unprotectOk Then
            rangeErr = False
            Set ur = Nothing
            Set ur = ws.UsedRange
            If Err.Number <> 0 Then
                skippedOut = skippedOut & "?E" & ws.Name & "?iUsedRange: " & Err.Description & "?j" & vbCrLf
                Err.Clear
                rangeErr = True
            End If
            If Not rangeErr Then
                If Not ur Is Nothing Then
                    If Not AssignFontNameToRange(ur, fontName) Then
                        skippedOut = skippedOut & "?E" & ws.Name & "?i?t?H???g?K?p???s?j" & vbCrLf
                    Else
                        If StrComp(Trim$(ws.Name), SHEET_RESULT_EQUIP_GANTT, vbBinaryCompare) <> 0 Then
                            ur.Columns.AutoFit
                        End If
                    End If
                End If
            End If
        End If
        
        If wasProt And unprotectOk Then
            Err.Clear
            If Len(SHEET_FONT_UNPROTECT_PASSWORD) > 0 Then
                ws.Protect Password:=SHEET_FONT_UNPROTECT_PASSWORD, DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True
            Else
                ws.Protect DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True
            End If
            If Err.Number <> 0 Then
                skippedOut = skippedOut & "?E" & ws.Name & "?i?????s: " & Err.Description & "?j" & vbCrLf
                Err.Clear
            End If
        End If
        On Error GoTo 0
    Next ws
    
    Application.Calculation = prevCalc
    Application.ScreenUpdating = prevScr
    Application.EnableEvents = prevEv
    
    Call TryAutoSaveMasterWorkbook
End Sub

Public Sub master_?S?V?[?g?t?H???g?????X?g????I?????????()
    Dim wsScratch As Worksheet
    Dim r As Range
    Dim prevWs As Worksheet
    Dim prevVis As XlSheetVisibility
    Dim oldName As String
    Dim oldSize As Variant
    Dim oldBold As Variant
    Dim oldItalic As Variant
    Dim oldUnderline As Variant
    Dim oldColor As Variant
    Dim oldStrike As Variant
    Dim picked As String
    Dim skipped As String
    Dim dlgOk As Boolean
    Dim v As Variant
    
    Set prevWs = ActiveSheet
    Set wsScratch = GetOrCreateFontScratchSheet()
    prevVis = wsScratch.Visible
    
    Set r = wsScratch.Range("A1")
    With r.Font
        oldName = .Name
        oldSize = .Size
        oldBold = .Bold
        oldItalic = .Italic
        oldUnderline = .Underline
        oldColor = .Color
        oldStrike = .Strikethrough
    End With
    
    wsScratch.Visible = xlSheetVisible
    wsScratch.Activate
    r.Select
    
    dlgOk = False
    Dim dlgRet As Variant
    On Error Resume Next
    dlgRet = Application.Dialogs(xlDialogFormatFont).Show
    If Err.Number <> 0 Then
        Err.Clear
        dlgOk = False
    Else
        Select Case VarType(dlgRet)
            Case vbBoolean
                dlgOk = dlgRet
            Case vbInteger, vbLong, vbByte
                dlgOk = (CLng(dlgRet) <> 0)
            Case Else
                dlgOk = False
        End Select
    End If
    On Error GoTo 0
    
    DoEvents
    
    picked = ""
    If dlgOk Then
        On Error Resume Next
        picked = Trim$(CStr(r.Font.Name))
        If Len(picked) = 0 Then
            picked = Trim$(CStr(r.Font.NameFarEast))
        End If
        On Error GoTo 0
    End If
    
    RestoreCellFontProps r, oldName, oldSize, oldBold, oldItalic, oldUnderline, oldColor, oldStrike
    wsScratch.Visible = prevVis
    On Error Resume Next
    prevWs.Activate
    On Error GoTo 0
    
    If Len(picked) = 0 Then
        v = Application.InputBox( _
            "?t?H???g????_?C?A???O???g??????A?????L?????Z???????????B" & vbCrLf & _
            "?K?p????t?H???g??????????????????i?z?[????t?H???g?{?b?N?X??????\?L?j?B" & vbCrLf & _
            "?i???????????umaster_?S?V?[?g?t?H???g???????????v??????????????????j", _
            "?S?V?[?g??t?H???g????", _
            oldName, _
            Type:=2)
        If VarType(v) = vbBoolean Then Exit Sub
        picked = Trim$(CStr(v))
        If Len(picked) = 0 Then
            MsgBox "?t?H???g????????????~????????B", vbExclamation
            Exit Sub
        End If
        If Not FontNameExistsInExcel(picked) Then
            If MsgBox( _
                "?t?H???g?u" & picked & "?v?????????????????????B" & vbCrLf & _
                "??????K?p???????????H", _
                vbQuestion Or vbYesNo, "?m?F") = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    On Error GoTo FailList
    ApplyFontToAllSheetCells picked, skipped
    On Error GoTo 0
    
    If Len(skipped) = 0 Then
        MsgBox "?S?V?[?g??Z???t?H???g???u" & picked & "?v??????????B", vbInformation
    Else
        MsgBox "?t?H???g?u" & picked & "?v??????????B?X?L?b?v?????V?[?g:" & vbCrLf & vbCrLf & skipped, vbExclamation
    End If
    Exit Sub
    
FailList:
    MsgBox "?t?H???g????G???[: " & Err.Description, vbCritical
End Sub

Public Sub master_?S?V?[?g?t?H???g??I?????????()
    Call master_?A?j???t??_?S?V?[?g?t?H???g?????X?g????I?????????
End Sub

Public Sub master_?S?V?[?g?t?H???g???????????()
    Dim v As Variant
    Dim fontName As String
    Dim skipped As String
    
    v = Application.InputBox( _
        "?K?p????t?H???g??????????????????B" & vbCrLf & _
        "?i?z?[????t?H???g?{?b?N?X??????\?L?j", _
        "?S?V?[?g??t?H???g????i?????j", _
        BIZ_UDP_GOTHIC_FONT_NAME, _
        Type:=2)
    If VarType(v) = vbBoolean Then Exit Sub
    
    fontName = Trim$(CStr(v))
    If Len(fontName) = 0 Then
        MsgBox "?t?H???g????????????~????????B", vbExclamation
        Exit Sub
    End If
    
    If Not FontNameExistsInExcel(fontName) Then
        If MsgBox( _
            "?t?H???g?u" & fontName & "?v?????????????????????B" & vbCrLf & _
            "??????K?p???????????H", _
            vbQuestion Or vbYesNo, "?m?F") = vbNo Then
            Exit Sub
        End If
    End If
    
    On Error GoTo FailHand
    ApplyFontToAllSheetCells fontName, skipped
    On Error GoTo 0
    
    If Len(skipped) = 0 Then
        MsgBox "?S?V?[?g??Z???t?H???g???u" & fontName & "?v??????????B", vbInformation
    Else
        MsgBox "?t?H???g?u" & fontName & "?v??????????B?X?L?b?v:" & vbCrLf & vbCrLf & skipped, vbExclamation
    End If
    Exit Sub
    
FailHand:
    MsgBox "?t?H???g????G???[: " & Err.Description, vbCritical
End Sub

Public Sub master_?S?V?[?g?t?H???g_BIZ_UDP?S?V?b?N?????()
    Dim skipped As String
    On Error GoTo FailB
    ApplyFontToAllSheetCells BIZ_UDP_GOTHIC_FONT_NAME, skipped
    On Error GoTo 0
    
    If Len(skipped) = 0 Then
        MsgBox "?S?V?[?g??Z???t?H???g???u" & BIZ_UDP_GOTHIC_FONT_NAME & "?v??????????B", vbInformation
    Else
        MsgBox "?t?H???g??????????B?X?L?b?v:" & vbCrLf & vbCrLf & skipped, vbExclamation
    End If
    Exit Sub
    
FailB:
    MsgBox "?t?H???g????G???[: " & Err.Description, vbCritical
End Sub

Public Sub master_?g??????\???X?V()
    Dim wb As Workbook
    Dim wsSkills As Worksheet
    Dim wsNeed As Worksheet
    Dim wsOut As Worksheet
    Dim rProc As Long, rMach As Long, rBase As Long
    Dim useTwoHeader As Boolean
    Dim maxReqNew As Long
    Dim maxReqOld As Long
    Dim maxReq As Long
    Dim totalRows As Long
    Dim fullRebuild As Boolean
    Dim remCol As Long
    Dim lastDataR As Long
    Dim maxId As Long
    Dim outRow As Long
    Dim sheetRowIndex As Long
    Dim existingKeys As Object
    Dim lastC As Long
    Dim c As Long
    Dim pKey As String
    Dim addedPairs As Long
    Dim appendedRows As Long
    Dim r As Long
    Dim rk As String
    Dim msgDetail As String
    
    On Error GoTo EH
    
    Set wb = ThisWorkbook
    On Error Resume Next
    Set wsSkills = wb.Worksheets(SHEET_SKILLS_DATA)
    Set wsNeed = wb.Worksheets(SHEET_NEED_DATA)
    On Error GoTo EH
    If wsSkills Is Nothing Then
        MsgBox "skills ?V?[?g??????????B", vbExclamation, "?g??????\"
        Exit Sub
    End If
    If wsNeed Is Nothing Then
        MsgBox "need ?V?[?g??????????B", vbExclamation, "?g??????\"
        Exit Sub
    End If
    
    If Not NeedFindBaseHeaders(wsNeed, rProc, rMach, rBase) Then
        MsgBox "need ?V?[?g??u?H?????v?u?@?B???v?u??{?K?v?l???v?s?????????????B" & vbCrLf & _
               "?iplanning_core ????l????C?A?E?g???K?v????j", vbExclamation, "?g??????\"
        Exit Sub
    End If
    
    maxReqNew = SkillComboScanMaxReq(wsSkills, wsNeed, rProc, rMach, rBase, useTwoHeader)
    If maxReqNew < 1 Then maxReqNew = 1
    If maxReqNew > 12 Then maxReqNew = 12
    
    If Not EnsureSkillCombinationSheetAfterSkills(wb, wsSkills, wsOut) Then
        MsgBox "?g??????\?V?[?g?????????????????B", vbExclamation, "?g??????\"
        Exit Sub
    End If
    
    fullRebuild = False
    If Trim$(CStr(wsOut.Cells(1, 1).Value)) <> "?g?????sID" Then fullRebuild = True
    remCol = SkillComboFindRemarksCol(wsOut)
    If remCol = 0 Then fullRebuild = True
    If Not fullRebuild Then
        maxReqOld = remCol - 7
        If maxReqOld < 1 Then fullRebuild = True
    End If
    
    maxReq = maxReqNew
    If Not fullRebuild Then
        If maxReqOld > maxReq Then maxReq = maxReqOld
    End If
    If maxReq < 1 Then maxReq = 1
    If maxReq > 12 Then maxReq = 12
    
    Application.ScreenUpdating = False
    
    If fullRebuild Then
        wsOut.Cells.Clear
        Call WriteSkillCombinationSheetHeader(wsOut, maxReq)
        totalRows = SkillComboFillBody(wsSkills, wsNeed, wsOut, rProc, rMach, rBase, useTwoHeader, maxReq)
        msgDetail = "?o??s??: " & CStr(totalRows) & "?i?S?s?? need / skills ???????????j"
    Else
        Call SkillComboEnsureMemberColumnsGrow(wsOut, maxReqOld, maxReq)
        Set existingKeys = CreateObject("Scripting.Dictionary")
        lastDataR = SkillComboDataLastRow(wsOut)
        If lastDataR < 2 Then lastDataR = 1
        For r = 2 To lastDataR
            rk = SkillComboRowGroupKey(wsOut, r)
            If Len(rk) > 0 Then existingKeys(rk) = True
        Next r
        maxId = SkillComboMaxNumericId(wsOut, lastDataR)
        outRow = lastDataR + 1
        sheetRowIndex = maxId
        addedPairs = 0
        lastC = wsSkills.Cells(1, wsSkills.Columns.Count).End(xlToLeft).Column
        If useTwoHeader Then
            For c = 2 To lastC
                pKey = SkillComboPairKeyFromSkillsColumn(wsSkills, c, True)
                If Len(pKey) = 0 Then GoTo NextInc1
                If existingKeys.Exists(pKey) Then GoTo NextInc1
                Call SkillComboWriteBodyForSkillsColumn(wsSkills, wsNeed, wsOut, c, rProc, rMach, rBase, True, maxReq, outRow, sheetRowIndex)
                existingKeys.Add pKey, True
                addedPairs = addedPairs + 1
NextInc1:
            Next c
        Else
            For c = 2 To lastC
                pKey = SkillComboPairKeyFromSkillsColumn(wsSkills, c, False)
                If Len(pKey) = 0 Then GoTo NextInc0
                If existingKeys.Exists(pKey) Then GoTo NextInc0
                Call SkillComboWriteBodyForSkillsColumn(wsSkills, wsNeed, wsOut, c, rProc, rMach, rBase, False, maxReq, outRow, sheetRowIndex)
                existingKeys.Add pKey, True
                addedPairs = addedPairs + 1
NextInc0:
            Next c
        End If
        appendedRows = outRow - lastDataR - 1
        If appendedRows < 0 Then appendedRows = 0
        totalRows = SkillComboDataLastRow(wsOut) - 1
        If totalRows < 0 Then totalRows = 0
        msgDetail = "??????g?????s???X??????????B" & vbCrLf & _
                    "?V?K??H????+?@?B???O???[?v: " & CStr(addedPairs) & " ????????L?i?f?[?^?s +" & CStr(appendedRows) & "?j?B" & vbCrLf & _
                    "???v?f?[?^?s??: " & CStr(totalRows)
    End If
    
    Call SkillComboSortDataRowsByMachine(wsOut, maxReq)
    Call SkillComboRenumberComboRowIds(wsOut)
    Call ApplySkillCombinationSheetFormat(wsOut, maxReq)
    Call FreezeTopRowSafe(wsOut)
    Application.ScreenUpdating = True
    
    Call TryAutoSaveMasterWorkbook
    
    MsgBox "?u" & SHEET_SKILL_COMBINATIONS & "?v???X?V????????B" & vbCrLf & msgDetail & vbCrLf & _
           "?f?[?^?s??@?B?????H???????K?v?l?????g?????D??x???H??+?@?B???g?????sID???????????i??L????ÑÑ????AA??E?g?????sID???ÉH??1????U?ëì???j?B" & vbCrLf & _
           "?g?????D??x??????A??i???W??j?B", vbInformation, "?g??????\"
    Exit Sub
    
EH:
    Application.ScreenUpdating = True
    MsgBox "?g??????\??X?V??G???[: " & Err.Description, vbCritical, "?g??????\"
End Sub

Private Sub FreezeTopRowSafe(ByVal ws As Worksheet)
    Dim wb As Workbook
    Set wb = ws.Parent
    On Error Resume Next
    wb.Activate
    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .Split = False
        .ScrollRow = 1
        .ScrollColumn = 1
    End With
    ws.Range("B2").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
End Sub

' ?H????+?@?B???i??2?E3?j?????s???4?u?H??+?@?B?v??O???[?v?L?[?????i??1??g?????sID?j
Private Function SkillComboRowGroupKey(ByVal ws As Worksheet, ByVal dataRow As Long) As String
    Dim s1 As String
    Dim s2 As String
    Dim s3 As String
    s1 = Trim$(CStr(ws.Cells(dataRow, 2).Value))
    s2 = Trim$(CStr(ws.Cells(dataRow, 3).Value))
    s3 = Trim$(CStr(ws.Cells(dataRow, 4).Value))
    If Len(s1) > 0 Or Len(s2) > 0 Then
        SkillComboRowGroupKey = s1 & vbTab & s2
    Else
        SkillComboRowGroupKey = s3
    End If
End Function

Private Sub SkillComboApplyGridBorders(ByVal rng As Range)
    Dim brColor As Long
    brColor = RGB(180, 180, 180)
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = brColor
    End With
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = brColor
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = brColor
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = brColor
    End With
    With rng.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = brColor
    End With
    With rng.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = brColor
    End With
End Sub

Private Sub ApplySkillCombinationSheetFormat(ByVal ws As Worksheet, ByVal maxReq As Long)
    Dim lastCol As Long
    Dim lastRow As Long
    Dim r As Long
    Dim hdrRange As Range
    Dim prevKey As String
    Dim curKey As String
    Dim altBand As Boolean
    Dim rowRng As Range
    Dim hdrFill As Long
    Dim hdrFont As Long
    Dim fillA As Long
    Dim fillB As Long
    Dim bodyRange As Range
    
    lastCol = 7 + maxReq
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 1 Then Exit Sub
    
    hdrFill = RGB(47, 84, 150)
    hdrFont = RGB(255, 255, 255)
    fillA = RGB(232, 240, 254)
    fillB = RGB(255, 255, 255)
    
    Set hdrRange = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    With hdrRange
        .Font.Bold = True
        .Font.Color = hdrFont
        .Interior.Pattern = xlSolid
        .Interior.Color = hdrFill
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    If lastRow >= 2 Then
        prevKey = vbNullString
        altBand = False
        For r = 2 To lastRow
            curKey = SkillComboRowGroupKey(ws, r)
            If curKey <> prevKey Then
                altBand = Not altBand
                prevKey = curKey
            End If
            Set rowRng = ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))
            With rowRng.Interior
                .Pattern = xlSolid
                If altBand Then
                    .Color = fillA
                Else
                    .Color = fillB
                End If
            End With
            With rowRng.Font
                .Bold = False
                .Color = RGB(0, 0, 0)
            End With
        Next r
    End If
    
    Set bodyRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    bodyRange.Columns.AutoFit
    Call SkillComboApplyGridBorders(bodyRange)
End Sub

Private Function EnsureSkillCombinationSheetAfterSkills(ByVal wb As Workbook, ByVal wsSkills As Worksheet, ByRef wsOut As Worksheet) As Boolean
    On Error Resume Next
    Set wsOut = wb.Worksheets(SHEET_SKILL_COMBINATIONS)
    On Error GoTo 0
    If wsOut Is Nothing Then
        Set wsOut = wb.Worksheets.Add(After:=wsSkills)
        On Error Resume Next
        wsOut.Name = SHEET_SKILL_COMBINATIONS
        On Error GoTo 0
    Else
        On Error Resume Next
        wsOut.Move After:=wsSkills
        On Error GoTo 0
    End If
    EnsureSkillCombinationSheetAfterSkills = Not (wsOut Is Nothing)
End Function

Private Function NeedFindBaseHeaders(ByVal ws As Worksheet, ByRef rProc As Long, ByRef rMach As Long, ByRef rBase As Long) As Boolean
    Dim r As Long
    Dim s0 As String
    Dim lastR As Long
    
    rProc = 0: rMach = 0: rBase = 0
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 1 Then Exit Function
    
    For r = 1 To lastR
        s0 = Trim$(CStr(ws.Cells(r, 1).Value))
        If Len(s0) = 0 Then GoTo NextNeedH
        If rProc = 0 And s0 = "?H????" Then rProc = r
        If rMach = 0 And s0 = "?@?B??" Then rMach = r
        If rBase = 0 Then
            If InStr(s0, "?K?v?l??") > 0 And Left$(s0, 4) <> "????w??" Then
                rBase = r
            End If
        End If
        If rProc > 0 And rMach > 0 And rBase > 0 Then Exit For
NextNeedH:
    Next r
    
    NeedFindBaseHeaders = (rProc > 0 And rMach > 0 And rBase > 0)
End Function

Private Function NeedGetBaseReqAtColumn(ByVal ws As Worksheet, ByVal rBase As Long, ByVal colIdx As Long) As Long
    Dim v As Variant
    Dim n As Long
    On Error Resume Next
    v = ws.Cells(rBase, colIdx).Value
    On Error GoTo 0
    If IsNumeric(v) Then
        n = CLng(CDbl(v))
        If n >= 1 Then
            NeedGetBaseReqAtColumn = n
            Exit Function
        End If
    End If
    NeedGetBaseReqAtColumn = 1
End Function

Private Function SkillComboScanMaxReq(ByVal wsSkill As Worksheet, ByVal wsNeed As Worksheet, _
    ByVal rProc As Long, ByVal rMach As Long, ByVal rBase As Long, ByRef useTwoHeader As Boolean) As Long
    Dim lastC As Long
    Dim c As Long
    Dim pm As Long
    Dim pHdr As String
    Dim mHdr As String
    Dim needCol As Long
    Dim rq As Long
    Dim mx As Long
    
    mx = 1
    useTwoHeader = False
    lastC = wsSkill.Cells(1, wsSkill.Columns.Count).End(xlToLeft).Column
    
    pm = 0
    For c = 2 To lastC
        pHdr = Trim$(CStr(wsSkill.Cells(1, c).Value))
        mHdr = Trim$(CStr(wsSkill.Cells(2, c).Value))
        If Len(pHdr) > 0 And Len(mHdr) > 0 Then
            If StrComp(LCase$(pHdr), "nan", vbTextCompare) <> 0 And StrComp(LCase$(mHdr), "nan", vbTextCompare) <> 0 Then
                pm = pm + 1
            End If
        End If
    Next c
    useTwoHeader = (pm > 0)
    
    If useTwoHeader Then
        For c = 2 To lastC
            pHdr = Trim$(CStr(wsSkill.Cells(1, c).Value))
            mHdr = Trim$(CStr(wsSkill.Cells(2, c).Value))
            If Len(pHdr) = 0 Or Len(mHdr) = 0 Then GoTo NextSC1
            needCol = NeedFindColumnForProcessMachine(wsNeed, rProc, rMach, pHdr, mHdr)
            If needCol > 0 Then
                rq = NeedGetBaseReqAtColumn(wsNeed, rBase, needCol)
                If rq > mx Then mx = rq
            End If
NextSC1:
        Next c
    Else
        ' 1?s?w?b?_: ?s1???????Bneed ???????u??o?????H??+?@?B?v?????@?B????????????v??????? ?? ??????????? need ?s???
        For c = 2 To lastC
            pHdr = Trim$(CStr(wsSkill.Cells(1, c).Value))
            If Len(pHdr) = 0 Then GoTo NextSC2
            needCol = NeedFindColumnByHeaderToken(wsNeed, rProc, rMach, pHdr)
            If needCol > 0 Then
                rq = NeedGetBaseReqAtColumn(wsNeed, rBase, needCol)
                If rq > mx Then mx = rq
            End If
NextSC2:
        Next c
    End If
    
    SkillComboScanMaxReq = mx
End Function

Private Function NeedFindColumnForProcessMachine(ByVal ws As Worksheet, ByVal rProc As Long, ByVal rMach As Long, _
    ByVal pWant As String, ByVal mWant As String) As Long
    Dim c As Long
    Dim lastC As Long
    Dim pc As String
    Dim mc As String
    
    lastC = ws.Cells(rProc, ws.Columns.Count).End(xlToLeft).Column
    For c = 4 To lastC
        pc = Trim$(CStr(ws.Cells(rProc, c).Value))
        mc = Trim$(CStr(ws.Cells(rMach, c).Value))
        If pc = pWant And mc = mWant Then
            NeedFindColumnForProcessMachine = c
            Exit Function
        End If
    Next c
    NeedFindColumnForProcessMachine = 0
End Function

' 1?s?w?b?_ skills ???? "?H??+?@?B" ?????P??g?[?N?????? need ????
Private Function NeedFindColumnByHeaderToken(ByVal ws As Worksheet, ByVal rProc As Long, ByVal rMach As Long, ByVal hdr As String) As Long
    Dim c As Long
    Dim lastC As Long
    Dim pc As String
    Dim mc As String
    Dim combo As String
    
    hdr = Trim$(hdr)
    lastC = ws.Cells(rProc, ws.Columns.Count).End(xlToLeft).Column
    For c = 4 To lastC
        pc = Trim$(CStr(ws.Cells(rProc, c).Value))
        mc = Trim$(CStr(ws.Cells(rMach, c).Value))
        combo = pc & "+" & mc
        If combo = hdr Or pc = hdr Or mc = hdr Then
            NeedFindColumnByHeaderToken = c
            Exit Function
        End If
    Next c
    NeedFindColumnByHeaderToken = 0
End Function

Private Sub WriteSkillCombinationSheetHeader(ByVal ws As Worksheet, ByVal maxReq As Long)
    Dim c As Long
    Dim col As Long
    ws.Cells(1, 1).Value = "?g?????sID"
    ws.Cells(1, 2).Value = "?H????"
    ws.Cells(1, 3).Value = "?@?B??"
    ws.Cells(1, 4).Value = "?H??+?@?B"
    ws.Cells(1, 5).Value = "?K?v?l??"
    ws.Cells(1, 6).Value = "?g?????D??x"
    col = 7
    For c = 1 To maxReq
        ws.Cells(1, col).Value = "?????o?[" & CStr(c)
        col = col + 1
    Next c
    ws.Cells(1, col).Value = "???l"
End Sub

Private Function SkillComboFindRemarksCol(ByVal ws As Worksheet) As Long
    Dim c As Long
    Dim lastC As Long
    lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastC
        If Trim$(CStr(ws.Cells(1, c).Value)) = "???l" Then
            SkillComboFindRemarksCol = c
            Exit Function
        End If
    Next c
    SkillComboFindRemarksCol = 0
End Function

Private Function SkillComboDataLastRow(ByVal ws As Worksheet) As Long
    Dim r1 As Long
    Dim r2 As Long
    Dim r3 As Long
    Dim r4 As Long
    r1 = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    r2 = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    r3 = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    r4 = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    SkillComboDataLastRow = r1
    If r2 > SkillComboDataLastRow Then SkillComboDataLastRow = r2
    If r3 > SkillComboDataLastRow Then SkillComboDataLastRow = r3
    If r4 > SkillComboDataLastRow Then SkillComboDataLastRow = r4
    If SkillComboDataLastRow < 1 Then SkillComboDataLastRow = 1
End Function

' ?f?[?^?s?i2?s???~?j???@?B????????????????i?s?S???????????K?v?l???E?g?????D??x??Z???l?????????j
' ?L?[??: ?@?B?????H???????K?v?l?????g?????D??x???H??+?@?B???g?????sID
Private Sub SkillComboSortDataRowsByMachine(ByVal ws As Worksheet, ByVal maxReq As Long)
    Dim lastRow As Long
    Dim lastCol As Long
    Dim sortRng As Range
    
    lastRow = SkillComboDataLastRow(ws)
    If lastRow < 2 Then Exit Sub
    lastCol = 7 + maxReq
    If lastCol < 7 Then lastCol = 7
    
    Set sortRng = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol))
    
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range(ws.Cells(2, 3), ws.Cells(lastRow, 3)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws.Range(ws.Cells(2, 2), ws.Cells(lastRow, 2)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws.Range(ws.Cells(2, 5), ws.Cells(lastRow, 5)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws.Range(ws.Cells(2, 6), ws.Cells(lastRow, 6)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws.Range(ws.Cells(2, 4), ws.Cells(lastRow, 4)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SortFields.Add Key:=ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1)), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange sortRng
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

' ?????????????????????AA??E?g?????sID?? 1 ????A???t??????
Private Sub SkillComboRenumberComboRowIds(ByVal ws As Worksheet)
    Dim lastRow As Long
    Dim r As Long
    Dim n As Long
    lastRow = SkillComboDataLastRow(ws)
    If lastRow < 2 Then Exit Sub
    n = 0
    For r = 2 To lastRow
        n = n + 1
        ws.Cells(r, 1).Value = n
    Next r
End Sub

Private Function SkillComboMaxNumericId(ByVal ws As Worksheet, ByVal upToRow As Long) As Long
    Dim r As Long
    Dim mx As Long
    Dim v As Variant
    mx = 0
    If upToRow < 2 Then
        SkillComboMaxNumericId = 0
        Exit Function
    End If
    For r = 2 To upToRow
        v = ws.Cells(r, 1).Value
        If IsNumeric(v) Then
            If CLng(CDbl(v)) > mx Then mx = CLng(CDbl(v))
        End If
    Next r
    SkillComboMaxNumericId = mx
End Function

Private Sub SkillComboEnsureMemberColumnsGrow(ByVal ws As Worksheet, ByVal maxReqOld As Long, ByVal maxReqNew As Long)
    Dim i As Long
    Dim ins As Long
    Dim remCol As Long
    Dim hdr As String
    If maxReqNew <= maxReqOld Then Exit Sub
    ins = maxReqNew - maxReqOld
    remCol = SkillComboFindRemarksCol(ws)
    If remCol = 0 Then Exit Sub
    For i = 1 To ins
        ws.Columns(remCol).Insert Shift:=xlToRight
        hdr = "?????o?[" & CStr(maxReqOld + i)
        ws.Cells(1, remCol).Value = hdr
        remCol = remCol + 1
    Next i
End Sub

' SkillComboRowGroupKey ??????L?[?K???i2?s?w?b?_?? ?H????+vbTab+?@?B???A1?s?w?b?_???o????????j
Private Function SkillComboPairKeyFromSkillsColumn(ByVal wsSkill As Worksheet, ByVal colIdx As Long, ByVal useTwoHeader As Boolean) As String
    Dim pHdr As String
    Dim mHdr As String
    If useTwoHeader Then
        pHdr = Trim$(CStr(wsSkill.Cells(1, colIdx).Value))
        mHdr = Trim$(CStr(wsSkill.Cells(2, colIdx).Value))
        If Len(pHdr) = 0 Or Len(mHdr) = 0 Then GoTo EmptyKey
        If StrComp(LCase$(pHdr), "nan", vbTextCompare) = 0 Then GoTo EmptyKey
        If StrComp(LCase$(mHdr), "nan", vbTextCompare) = 0 Then GoTo EmptyKey
        SkillComboPairKeyFromSkillsColumn = pHdr & vbTab & mHdr
        Exit Function
    Else
        pHdr = Trim$(CStr(wsSkill.Cells(1, colIdx).Value))
        If Len(pHdr) = 0 Then GoTo EmptyKey
        If StrComp(LCase$(pHdr), "nan", vbTextCompare) = 0 Then GoTo EmptyKey
        SkillComboPairKeyFromSkillsColumn = pHdr
        Exit Function
    End If
EmptyKey:
    SkillComboPairKeyFromSkillsColumn = ""
End Function

' skills ?? 1 ????????g??????\?? outRow ??~??o??i?g?????sID ?? sheetRowIndex ????j
Private Sub SkillComboWriteBodyForSkillsColumn( _
    ByVal wsSkill As Worksheet, ByVal wsNeed As Worksheet, ByVal wsOut As Worksheet, _
    ByVal colIdx As Long, _
    ByVal rProc As Long, ByVal rMach As Long, ByVal rBase As Long, _
    ByVal useTwoHeader As Boolean, ByVal maxReq As Long, _
    ByRef outRow As Long, ByRef sheetRowIndex As Long)
    
    Dim pHdr As String
    Dim mHdr As String
    Dim combo As String
    Dim needCol As Long
    Dim req As Long
    Dim lastRowS As Long
    Dim r As Long
    Dim mem As String
    Dim cellV As String
    Dim roleCh As String
    Dim prV As Long
    Dim okCell As Boolean
    Dim nMem As Long
    Dim names() As String
    Dim isOp() As Boolean
    Dim linesColl As Collection
    Dim dup As Object
    Dim k As Variant
    Dim parts() As String
    Dim jj As Long
    Dim prio As Long
    Dim note As String
    Dim memCol As Long
    
    If useTwoHeader Then
        pHdr = Trim$(CStr(wsSkill.Cells(1, colIdx).Value))
        mHdr = Trim$(CStr(wsSkill.Cells(2, colIdx).Value))
        If Len(pHdr) = 0 Or Len(mHdr) = 0 Then Exit Sub
        combo = pHdr & "+" & mHdr
        needCol = NeedFindColumnForProcessMachine(wsNeed, rProc, rMach, pHdr, mHdr)
        If needCol <= 0 Then
            Call WriteComboWarningRow(wsOut, outRow, pHdr, mHdr, combo, 0, maxReq, "need ?????????????i?H?????E?@?B?????v???m?F?j")
            outRow = outRow + 1
            Exit Sub
        End If
        req = NeedGetBaseReqAtColumn(wsNeed, rBase, needCol)
        If req > maxReq Then req = maxReq
        
        nMem = 0
        lastRowS = wsSkill.Cells(wsSkill.Rows.Count, 1).End(xlUp).Row
        For r = 3 To lastRowS
            mem = Trim$(CStr(wsSkill.Cells(r, 1).Value))
            If Len(mem) = 0 Or StrComp(LCase$(mem), "nan", vbTextCompare) = 0 Then GoTo NextMem1
            cellV = Trim$(CStr(wsSkill.Cells(r, colIdx).Value))
            okCell = MasterParseOpAsCell(cellV, roleCh, prV)
            If Not okCell Then GoTo NextMem1
            nMem = nMem + 1
            If nMem = 1 Then
                ReDim names(1 To 1)
                ReDim isOp(1 To 1)
            Else
                ReDim Preserve names(1 To nMem)
                ReDim Preserve isOp(1 To nMem)
            End If
            names(nMem) = mem
            isOp(nMem) = (StrComp(roleCh, "OP", vbTextCompare) = 0)
NextMem1:
        Next r
        
        If nMem < req Then
            note = "?X?L??(OP/AS)???ñb?K?v?l?? " & CStr(req) & " ?????i" & CStr(nMem) & "???j?????g???????"
            Call WriteComboWarningRow(wsOut, outRow, pHdr, mHdr, combo, req, maxReq, note)
            outRow = outRow + 1
            Exit Sub
        End If
        
        Set linesColl = New Collection
        Set dup = CreateObject("Scripting.Dictionary")
        Call BuildSkillCombinations(nMem, req, names, isOp, linesColl, dup)
        
        prio = 0
        For Each k In SortCollectionStrings(linesColl)
            prio = prio + 1
            sheetRowIndex = sheetRowIndex + 1
            wsOut.Cells(outRow, 1).Value = sheetRowIndex
            wsOut.Cells(outRow, 2).Value = pHdr
            wsOut.Cells(outRow, 3).Value = mHdr
            wsOut.Cells(outRow, 4).Value = combo
            wsOut.Cells(outRow, 5).Value = req
            wsOut.Cells(outRow, 6).Value = prio
            parts = Split(CStr(k), "|")
            For jj = LBound(parts) To UBound(parts)
                wsOut.Cells(outRow, 7 + jj).Value = parts(jj)
            Next jj
            For jj = UBound(parts) + 1 To maxReq - 1
                wsOut.Cells(outRow, 7 + jj).Value = ""
            Next jj
            wsOut.Cells(outRow, 7 + maxReq).Value = ""
            outRow = outRow + 1
        Next k
        
        If prio = 0 Then
            note = "OP1?????????g??????????????i?v???[?_?[?????j"
            Call WriteComboWarningRow(wsOut, outRow, pHdr, mHdr, combo, req, maxReq, note)
            outRow = outRow + 1
        End If
    Else
        memCol = 1
        pHdr = Trim$(CStr(wsSkill.Cells(1, colIdx).Value))
        If Len(pHdr) = 0 Then Exit Sub
        needCol = NeedFindColumnByHeaderToken(wsNeed, rProc, rMach, pHdr)
        mHdr = ""
        combo = pHdr
        If needCol <= 0 Then
            Call WriteComboWarningRow(wsOut, outRow, "", "", combo, 0, maxReq, "need ?????????????i??o???? need ?????j")
            outRow = outRow + 1
            Exit Sub
        End If
        req = NeedGetBaseReqAtColumn(wsNeed, rBase, needCol)
        If req > maxReq Then req = maxReq
        
        nMem = 0
        lastRowS = wsSkill.Cells(wsSkill.Rows.Count, memCol).End(xlUp).Row
        For r = 2 To lastRowS
            mem = Trim$(CStr(wsSkill.Cells(r, memCol).Value))
            If Len(mem) = 0 Or StrComp(LCase$(mem), "nan", vbTextCompare) = 0 Then GoTo NextMem0
            cellV = Trim$(CStr(wsSkill.Cells(r, colIdx).Value))
            okCell = MasterParseOpAsCell(cellV, roleCh, prV)
            If Not okCell Then GoTo NextMem0
            nMem = nMem + 1
            If nMem = 1 Then
                ReDim names(1 To 1)
                ReDim isOp(1 To 1)
            Else
                ReDim Preserve names(1 To nMem)
                ReDim Preserve isOp(1 To nMem)
            End If
            names(nMem) = mem
            isOp(nMem) = (StrComp(roleCh, "OP", vbTextCompare) = 0)
NextMem0:
        Next r
        
        If nMem < req Then
            note = "?X?L??(OP/AS)???ñb?K?v?l???????????g???????"
            Call WriteComboWarningRow(wsOut, outRow, "", "", combo, req, maxReq, note)
            outRow = outRow + 1
            Exit Sub
        End If
        
        Set linesColl = New Collection
        Set dup = CreateObject("Scripting.Dictionary")
        Call BuildSkillCombinations(nMem, req, names, isOp, linesColl, dup)
        
        prio = 0
        For Each k In SortCollectionStrings(linesColl)
            prio = prio + 1
            sheetRowIndex = sheetRowIndex + 1
            wsOut.Cells(outRow, 1).Value = sheetRowIndex
            wsOut.Cells(outRow, 2).Value = ""
            wsOut.Cells(outRow, 3).Value = ""
            wsOut.Cells(outRow, 4).Value = combo
            wsOut.Cells(outRow, 5).Value = req
            wsOut.Cells(outRow, 6).Value = prio
            parts = Split(CStr(k), "|")
            For jj = LBound(parts) To UBound(parts)
                wsOut.Cells(outRow, 7 + jj).Value = parts(jj)
            Next jj
            For jj = UBound(parts) + 1 To maxReq - 1
                wsOut.Cells(outRow, 7 + jj).Value = ""
            Next jj
            wsOut.Cells(outRow, 7 + maxReq).Value = ""
            outRow = outRow + 1
        Next k
        
        If prio = 0 Then
            Call WriteComboWarningRow(wsOut, outRow, "", "", combo, req, maxReq, "OP1?????????g??????????????")
            outRow = outRow + 1
        End If
    End If
End Sub

Private Function SkillComboFillBody(ByVal wsSkill As Worksheet, ByVal wsNeed As Worksheet, ByVal wsOut As Worksheet, _
    ByVal rProc As Long, ByVal rMach As Long, ByVal rBase As Long, ByVal useTwoHeader As Boolean, ByVal maxReq As Long) As Long
    Dim lastC As Long
    Dim c As Long
    Dim outRow As Long
    Dim sheetRowIndex As Long
    
    outRow = 2
    sheetRowIndex = 0
    lastC = wsSkill.Cells(1, wsSkill.Columns.Count).End(xlToLeft).Column
    If useTwoHeader Then
        For c = 2 To lastC
            Call SkillComboWriteBodyForSkillsColumn(wsSkill, wsNeed, wsOut, c, rProc, rMach, rBase, True, maxReq, outRow, sheetRowIndex)
        Next c
    Else
        For c = 2 To lastC
            Call SkillComboWriteBodyForSkillsColumn(wsSkill, wsNeed, wsOut, c, rProc, rMach, rBase, False, maxReq, outRow, sheetRowIndex)
        Next c
    End If
    SkillComboFillBody = outRow - 2
    If SkillComboFillBody < 0 Then SkillComboFillBody = 0
End Function

Private Sub WriteComboWarningRow(ByVal ws As Worksheet, ByVal outRow As Long, ByVal pH As String, ByVal mH As String, ByVal combo As String, _
    ByVal req As Long, ByVal maxReq As Long, ByVal note As String)
    ws.Cells(outRow, 1).Value = ""
    ws.Cells(outRow, 2).Value = pH
    ws.Cells(outRow, 3).Value = mH
    ws.Cells(outRow, 4).Value = combo
    If req > 0 Then ws.Cells(outRow, 5).Value = req Else ws.Cells(outRow, 5).Value = ""
    ws.Cells(outRow, 6).Value = ""
    ws.Cells(outRow, 7 + maxReq).Value = note
End Sub

Private Function MasterParseOpAsCell(ByVal s As String, ByRef roleOut As String, ByRef prOut As Long) As Boolean
    Dim t As String
    Dim tail As String
    MasterParseOpAsCell = False
    roleOut = ""
    prOut = 0
    t = Replace(Replace(UCase$(Trim$(s)), " ", ""), vbTab, "")
    If Len(t) = 0 Then Exit Function
    If Left$(t, 2) = "OP" Then
        roleOut = "OP"
        tail = Mid$(t, 3)
    ElseIf Left$(t, 2) = "AS" Then
        roleOut = "AS"
        tail = Mid$(t, 3)
    Else
        Exit Function
    End If
    If Len(tail) = 0 Then
        prOut = 1
    Else
        If Not IsNumeric(tail) Then Exit Function
        prOut = CLng(CDbl(tail))
        If prOut < 0 Then prOut = 0
    End If
    MasterParseOpAsCell = True
End Function

Private Sub BuildSkillCombinations(ByVal nMem As Long, ByVal req As Long, ByRef names() As String, ByRef isOp() As Boolean, _
    ByRef linesColl As Collection, ByRef dup As Object)
    Dim chosen() As Long
    If req < 1 Or nMem < req Then Exit Sub
    ReDim chosen(1 To req)
    Call RecurSkillCombos(nMem, req, names, isOp, 1, 1, chosen, linesColl, dup)
End Sub

Private Sub RecurSkillCombos(ByVal nMem As Long, ByVal req As Long, ByRef names() As String, ByRef isOp() As Boolean, _
    ByVal depth As Long, ByVal startAt As Long, ByRef chosen() As Long, ByRef linesColl As Collection, ByRef dup As Object)
    Dim i As Long
    Dim j As Long
    Dim hasOp As Boolean
    Dim parts() As String
    Dim line As String
    
    If depth > req Then
        hasOp = False
        For j = 1 To req
            If isOp(chosen(j)) Then hasOp = True
        Next j
        If Not hasOp Then Exit Sub
        ReDim parts(1 To req)
        For j = 1 To req
            parts(j) = names(chosen(j))
        Next j
        Call SortStringArrayAsc1(parts)
        line = Join(parts, "|")
        If Not dup.Exists(line) Then
            dup.Add line, True
            linesColl.Add line
        End If
        Exit Sub
    End If
    
    For i = startAt To nMem
        chosen(depth) = i
        Call RecurSkillCombos(nMem, req, names, isOp, depth + 1, i + 1, chosen, linesColl, dup)
    Next i
End Sub

Private Sub SortStringArrayAsc1(ByRef a() As String)
    Dim i As Long
    Dim j As Long
    Dim t As String
    Dim lb As Long
    Dim ub As Long
    lb = LBound(a)
    ub = UBound(a)
    For i = lb To ub - 1
        For j = i + 1 To ub
            If StrComp(a(i), a(j), vbBinaryCompare) > 0 Then
                t = a(i): a(i) = a(j): a(j) = t
            End If
        Next j
    Next i
End Sub

Private Function SortCollectionStrings(ByVal coll As Collection) As Collection
    Dim arr() As String
    Dim i As Long
    Dim n As Long
    Dim outC As Collection
    Set outC = New Collection
    n = coll.Count
    If n = 0 Then
        Set SortCollectionStrings = outC
        Exit Function
    End If
    ReDim arr(1 To n)
    For i = 1 To n
        arr(i) = coll(i)
    Next i
    Call SortStringArrayAsc1(arr)
    For i = 1 To n
        outC.Add arr(i)
    Next i
    Set SortCollectionStrings = outC
End Function

' -------------------------------------------------------------------------
' ???_??????O?áq?? / ???_?@?B_?????n????? ?c skills ??H???~?@?B??]?L????????????????
' -------------------------------------------------------------------------

Private Function MasterChangeoverFindRow(ByVal wsCh As Worksheet, ByVal procNm As String, ByVal machNm As String) As Long
    Dim r As Long
    Dim lr As Long
    Dim a As String
    Dim b As String
    MasterChangeoverFindRow = 0
    lr = wsCh.Cells(wsCh.Rows.Count, 1).End(xlUp).Row
    If lr < 2 Then Exit Function
    For r = 2 To lr
        a = Trim$(CStr(wsCh.Cells(r, 1).Value))
        b = Trim$(CStr(wsCh.Cells(r, 2).Value))
        If StrComp(a, procNm, vbTextCompare) = 0 And StrComp(b, machNm, vbTextCompare) = 0 Then
            MasterChangeoverFindRow = r
            Exit Function
        End If
    Next r
End Function

Private Function MasterDailyStartupFindRow(ByVal wsSu As Worksheet, ByVal machNm As String) As Long
    Dim r As Long
    Dim lr As Long
    MasterDailyStartupFindRow = 0
    lr = wsSu.Cells(wsSu.Rows.Count, 1).End(xlUp).Row
    If lr < 2 Then Exit Function
    For r = 2 To lr
        If StrComp(Trim$(CStr(wsSu.Cells(r, 1).Value)), machNm, vbTextCompare) = 0 Then
            MasterDailyStartupFindRow = r
            Exit Function
        End If
    Next r
End Function

' skills 1 ?s????H?????E2 ?s????@?B?????????i?@?B?J?????_?[?????|?j
Private Sub MasterSyncChangeoverRowsFromSkills(ByVal wb As Workbook)
    Dim wsSk As Worksheet
    Dim wsCh As Worksheet
    Dim lastC As Long
    Dim c As Long
    Dim pHdr As String
    Dim mHdr As String
    Dim rNew As Long
    On Error Resume Next
    Set wsSk = wb.Worksheets(SHEET_SKILLS_DATA)
    On Error GoTo 0
    If wsSk Is Nothing Then Exit Sub
    On Error Resume Next
    Set wsCh = wb.Worksheets(SHEET_MACHINE_CHANGEOVER)
    On Error GoTo 0
    If wsCh Is Nothing Then Exit Sub

    lastC = wsSk.Cells(1, wsSk.Columns.Count).End(xlToLeft).Column
    If lastC < 2 Then Exit Sub

    For c = 2 To lastC
        pHdr = Trim$(CStr(wsSk.Cells(1, c).Value))
        mHdr = Trim$(CStr(wsSk.Cells(2, c).Value))
        If Len(pHdr) = 0 Or Len(mHdr) = 0 Then GoTo NextSkillCol
        If StrComp(LCase$(pHdr), "nan", vbTextCompare) = 0 Then GoTo NextSkillCol
        If StrComp(LCase$(mHdr), "nan", vbTextCompare) = 0 Then GoTo NextSkillCol
        If MasterChangeoverFindRow(wsCh, pHdr, mHdr) > 0 Then GoTo NextSkillCol
        rNew = wsCh.Cells(wsCh.Rows.Count, 1).End(xlUp).Row + 1
        If rNew < 2 Then rNew = 2
        wsCh.Cells(rNew, 1).Value = pHdr
        wsCh.Cells(rNew, 2).Value = mHdr
        ' C,D ??????i???[?U?[?????????j
NextSkillCol:
    Next c
End Sub

Private Sub MasterSyncDailyStartupRowsFromSkills(ByVal wb As Workbook)
    Dim wsSk As Worksheet
    Dim wsSu As Worksheet
    Dim lastC As Long
    Dim c As Long
    Dim mHdr As String
    Dim dict As Object
    Dim k As Variant
    Dim arr() As String
    Dim n As Long
    Dim i As Long
    Dim rNew As Long
    On Error Resume Next
    Set wsSk = wb.Worksheets(SHEET_SKILLS_DATA)
    On Error GoTo 0
    If wsSk Is Nothing Then Exit Sub
    On Error Resume Next
    Set wsSu = wb.Worksheets(SHEET_MACHINE_DAILY_STARTUP)
    On Error GoTo 0
    If wsSu Is Nothing Then Exit Sub

    Set dict = CreateObject("Scripting.Dictionary")
    lastC = wsSk.Cells(2, wsSk.Columns.Count).End(xlToLeft).Column
    If lastC < 2 Then Exit Sub
    For c = 2 To lastC
        mHdr = Trim$(CStr(wsSk.Cells(2, c).Value))
        If Len(mHdr) = 0 Then GoTo NextMachCol
        If StrComp(LCase$(mHdr), "nan", vbTextCompare) = 0 Then GoTo NextMachCol
        If Not dict.Exists(mHdr) Then dict.Add mHdr, True
NextMachCol:
    Next c
    If dict.Count = 0 Then Exit Sub

    ReDim arr(1 To dict.Count)
    n = 0
    For Each k In dict.Keys
        n = n + 1
        arr(n) = CStr(k)
    Next k
    Call SortStringArrayAsc1(arr)

    For i = 1 To UBound(arr)
        mHdr = arr(i)
        If MasterDailyStartupFindRow(wsSu, mHdr) > 0 Then GoTo NextMachNm
        rNew = wsSu.Cells(wsSu.Rows.Count, 1).End(xlUp).Row + 1
        If rNew < 2 Then rNew = 2
        wsSu.Cells(rNew, 1).Value = mHdr
        ' B ????i???[?U?[?????????j
NextMachNm:
    Next i
End Sub

' ???NO???????/??n???E?@?B?????????n??????B
' ?E?V?[?g?????????????A1 ?s?????o?????????iA1 ????????????o?????????j?B
' ?Eskills ?? 1?`2 ?s??i?H?????E?@?B???j????A???o?^??s?????????i?????s?? C ???~???X??????j?B
Public Sub MasterEnsureMachineChangeoverTemplateSheets()
    Dim wb As Workbook
    Dim ws As Worksheet
    On Error Resume Next
    Set wb = ThisWorkbook
    On Error GoTo 0
    If wb Is Nothing Then Exit Sub

    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_MACHINE_CHANGEOVER)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        On Error Resume Next
        ws.Name = SHEET_MACHINE_CHANGEOVER
        On Error GoTo 0
    End If
    If Not ws Is Nothing Then
        If Len(Trim$(CStr(ws.Cells(1, 1).Value))) = 0 Then
            ws.Cells(1, 1).Value = "?H????"
            ws.Cells(1, 2).Value = "?@?B??"
            ws.Cells(1, 3).Value = "????????_??"
            ws.Cells(1, 4).Value = "??n??????_??"
            ws.Rows(1).Font.Bold = True
        End If
    End If

    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_MACHINE_DAILY_STARTUP)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        On Error Resume Next
        ws.Name = SHEET_MACHINE_DAILY_STARTUP
        On Error GoTo 0
    End If
    If Not ws Is Nothing Then
        If Len(Trim$(CStr(ws.Cells(1, 1).Value))) = 0 Then
            ws.Cells(1, 1).Value = "?@?B??"
            ws.Cells(1, 2).Value = "?????n?????_??"
            ws.Rows(1).Font.Bold = True
        End If
    End If

    Call MasterSyncChangeoverRowsFromSkills(wb)
    Call MasterSyncDailyStartupRowsFromSkills(wb)
End Sub

' ???C???????V?[?g?? H4 ?????AMasterEnsureMachineChangeoverTemplateSheets ?p??u???????????{?^???v??1??z?u????B
' ?}?`?? Btn_ChangeoverTemplateMain?i????s????????ëì???j?BOnAction = master_?A?j???t??_??????????n??V?[?g???X?V?i????????A?j???t???j?B
Public Sub master_??????????n??_?A?j???t???{?^??????()
    Dim ws As Worksheet
    Const shpName As String = "Btn_ChangeoverTemplateMain"
    On Error Resume Next
    Set ws = MasterGetMainWorksheet(ThisWorkbook)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "???C???V?[?g?????????????B?A?N?e?B?u?V?[?g?????????B", vbExclamation, "??????E?????n??"
        Set ws = ActiveSheet
    End If
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Shapes(shpName).Delete
    On Error GoTo 0
    ws.Activate
    Dim posX As Single
    Dim posY As Single
    posX = ws.Range("H4").Left
    posY = ws.Range("H4").Top
    CreateCoolButtonWithPreset "??????E?????n????X?V", "master_?A?j???t??_??????????n??V?[?g???X?V", posX, posY, 4, shpName
    MsgBox "?{?^????z?u????????B" & vbCrLf & _
           "?E?}?N??: MasterEnsureMachineChangeoverTemplateSheets?i???o???Eskills ???????j" & vbCrLf & _
           "?E?}?`??: " & shpName & "?i????s??u???????j" & vbCrLf & _
           "?E?????????????????????{?^??????l?????????A?j?????t??????B", vbInformation, "??????E?????n??"
End Sub

' ?u?b?N???J????????? Ctrl+Shift+?e???L?[ - ??o?^?i?????? ThisWorkbook ?? BeforeClose?B?S???? master_xlsm_ThisWorkbook_VBA.txt?j
Sub Auto_Open()
    MasterShortcutMainSheet_OnKeyRegister
End Sub

