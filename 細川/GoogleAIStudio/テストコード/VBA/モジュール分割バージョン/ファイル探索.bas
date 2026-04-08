Attribute VB_Name = "ファイル探索"
Option Explicit

Public Sub メインシート_masterブックを開く()
    Dim path As String
    Dim folder As String
    Dim wb As Workbook
    Dim wbMaster As Workbook
    
    folder = ThisWorkbook.path
    If Len(folder) = 0 Then
        MsgBox "ブックを一度保存してから実行してください。", vbExclamation
        Exit Sub
    End If
    path = folder & "\" & MASTER_WORKBOOK_FILE
    If Len(Dir(path)) = 0 Then
        MsgBox "次のファイルが見つかりません。" & vbCrLf & path, vbExclamation
        Exit Sub
    End If
    
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then
            wb.Activate
            MacroSplash_SetStep "master.xlsm は既に開いています（アクティブにしました）。"
            m_animMacroSucceeded = True
            Exit Sub
        End If
    Next wb
    
    On Error GoTo OpenFail
    MacroSplash_SetStep "master.xlsm を開いています…"
    Set wbMaster = Application.Workbooks.Open(Filename:=path)
    wbMaster.Activate
    MacroSplash_SetStep "master.xlsm を開きました。"
    m_animMacroSucceeded = True
    Exit Sub
OpenFail:
    MsgBox "master.xlsm を開けませんでした: " & Err.Description, vbCritical
End Sub

Private Sub マスタメイン_工場稼働と定常を取得( _
    ByRef facOk As Boolean, ByRef facS As Date, ByRef facE As Date, _
    ByRef regOk As Boolean, ByRef regS As Date, ByRef regE As Date)
    
    Dim folder As String
    Dim p As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim openedHere As Boolean
    Dim v As Variant
    Dim tS As Date
    Dim tE As Date
    
    facOk = False
    regOk = False
    folder = ThisWorkbook.path
    If Len(folder) = 0 Then Exit Sub
    p = folder & "\" & MASTER_WORKBOOK_FILE
    If Len(Dir(p)) = 0 Then Exit Sub
    
    openedHere = False
    Set wb = Nothing
    On Error Resume Next
    Set wb = Workbooks(MASTER_WORKBOOK_FILE)
    On Error GoTo 0
    If wb Is Nothing Then
        On Error Resume Next
        Set wb = Workbooks.Open(Filename:=p, ReadOnly:=True, UpdateLinks:=0)
        openedHere = Not (wb Is Nothing)
        On Error GoTo 0
    End If
    If wb Is Nothing Then Exit Sub
    
    Set ws = マスタブック_メイン設定シートを取得(wb)
    If ws Is Nothing Then GoTo CloseMasterWb
    
    v = マスタメイン_結合左上の値(ws, "A12")
    If マスタメイン_セルを時刻Dateへ(v, tS) Then
        v = マスタメイン_結合左上の値(ws, "B12")
        If マスタメイン_セルを時刻Dateへ(v, tE) Then
            If TimeValue(tS) < TimeValue(tE) Then
                facOk = True
                facS = tS
                facE = tE
            End If
        End If
    End If
    
    v = マスタメイン_結合左上の値(ws, "A15")
    If マスタメイン_セルを時刻Dateへ(v, tS) Then
        v = マスタメイン_結合左上の値(ws, "B15")
        If マスタメイン_セルを時刻Dateへ(v, tE) Then
            If TimeValue(tS) < TimeValue(tE) Then
                regOk = True
                regS = tS
                regE = tE
            End If
        End If
    End If

CloseMasterWb:
    If openedHere Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
End Sub

Private Function マスタメイン_工場標準勤怠表示文字列() As String
    Const FB As String = "08:45 / 17:00"
    Dim folder As String
    Dim p As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim openedHere As Boolean
    Dim vS As Variant, vE As Variant
    Dim tS As Date, tE As Date
    
    マスタメイン_工場標準勤怠表示文字列 = FB
    On Error GoTo CleanExit
    
    folder = ThisWorkbook.path
    If Len(folder) = 0 Then Exit Function
    p = folder & "\" & MASTER_WORKBOOK_FILE
    If Len(Dir(p)) = 0 Then Exit Function
    
    openedHere = False
    Set wb = Nothing
    On Error Resume Next
    Set wb = Workbooks(MASTER_WORKBOOK_FILE)
    On Error GoTo CleanExit
    If wb Is Nothing Then
        On Error Resume Next
        Set wb = Workbooks.Open(Filename:=p, ReadOnly:=True, UpdateLinks:=0)
        On Error GoTo CleanExit
        openedHere = Not (wb Is Nothing)
    End If
    If wb Is Nothing Then Exit Function
    
    Set ws = マスタブック_メイン設定シートを取得(wb)
    If ws Is Nothing Then GoTo CloseWb
    
    vS = マスタメイン_結合左上の値(ws, "A15")
    vE = マスタメイン_結合左上の値(ws, "B15")
    If Not マスタメイン_セルを時刻Dateへ(vS, tS) Then GoTo CloseWb
    If Not マスタメイン_セルを時刻Dateへ(vE, tE) Then GoTo CloseWb
    If TimeValue(tS) >= TimeValue(tE) Then GoTo CloseWb
    
    マスタメイン_工場標準勤怠表示文字列 = Format$(tS, "hh:nn") & " / " & Format$(tE, "hh:nn")

CloseWb:
    If openedHere Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    Exit Function
CleanExit:
    On Error Resume Next
    If openedHere And Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Function

Private Function MacroCompleteChime_EnsureWavPath() As String
    Dim p As String
    Dim dirSounds As String
    p = MacroCompleteChime_LocalWavPath()
    If Len(p) = 0 Then Exit Function
    If Len(Dir(p)) > 0 Then
        MacroCompleteChime_EnsureWavPath = p
        Exit Function
    End If
    dirSounds = ThisWorkbook.path & "\" & MACRO_COMPLETE_CHIME_REL_DIR
    On Error Resume Next
    MkDir dirSounds
    On Error GoTo 0
    If MacroCompleteChime_HttpDownloadBinary(MACRO_COMPLETE_CHIME_DOWNLOAD_URL, p) Then
        If Len(Dir(p)) > 0 Then MacroCompleteChime_EnsureWavPath = p
    End If
End Function

Private Sub MacroCompleteChime()
    On Error Resume Next
    If Not m_splashAllowMacroSound Then Exit Sub
    Dim track As Long
    Dim mp3 As String
    Dim wav As String
    track = SettingsSheet_GetCompleteChimeTrack1to4()
    mp3 = MacroCompleteChime_LocalMp3Path(track)
    If Len(mp3) > 0 And Len(Dir(mp3)) > 0 Then
        If MacroCompleteChime_MciPlayMp3(mp3) Then Exit Sub
    End If
    wav = MacroCompleteChime_EnsureWavPath()
    If Len(wav) > 0 Then
        PlaySoundW StrPtr(wav), 0&, SND_FILENAME Or SND_ASYNC
    Else
        PlaySound "SystemAsterisk", 0&, SND_ALIAS Or SND_ASYNC
    End If
End Sub

Private Function FileHasUtf8Bom(ByVal filePath As String) As Boolean
    Dim ff As Integer
    Dim b1 As Byte, b2 As Byte, b3 As Byte
    On Error GoTo CleanFail
    If Len(Dir(filePath)) = 0 Then FileHasUtf8Bom = False: Exit Function
    ff = FreeFile
    Open filePath For Binary Access Read As #ff
    Get #ff, 1, b1
    Get #ff, 2, b2
    Get #ff, 3, b3
    Close #ff
    FileHasUtf8Bom = (b1 = &HEF And b2 = &HBB And b3 = &HBF)
    Exit Function
CleanFail:
    On Error Resume Next
    If ff <> 0 Then Close #ff
    FileHasUtf8Bom = False
End Function

Private Function ReadCmdCaptureLogText(ByVal filePath As String) As String
    On Error GoTo EH
    If Len(Dir(filePath)) = 0 Then Exit Function
    If FileHasUtf8Bom(filePath) Then
        ReadCmdCaptureLogText = ReadTextFileWithCharset(filePath, "utf-8")
    Else
        ReadCmdCaptureLogText = ReadTextFileWithCharset(filePath, "Windows-932")
    End If
    Exit Function
EH:
    ReadCmdCaptureLogText = ""
End Function

Public Sub 設定_配台不要工程_E列_TSVから反映()
    Dim targetDir As String
    Dim tsvPath As String
    Dim jsonPath As String
    Dim stm As Object
    Dim text As String
    Dim lines() As String
    Dim i As Long
    Dim hdrEnd As Long
    Dim ln As String
    Dim wbExpected As String
    Dim colE As Long
    Dim ws As Worksheet
    Dim parts() As String
    Dim rNum As Long
    Dim cellVal As String

    On Error GoTo CleanFail

    targetDir = ThisWorkbook.path
    If Len(targetDir) = 0 Then Exit Sub

    tsvPath = targetDir & "\log\exclude_rules_e_column_vba.tsv"
    If Len(Dir(tsvPath)) = 0 Then Exit Sub

    Set stm = CreateObject("ADODB.Stream")
    stm.charset = "UTF-8"
    stm.Open
    stm.LoadFromFile tsvPath
    text = stm.ReadText
    stm.Close
    Set stm = Nothing

    text = Replace(text, vbCrLf, vbLf)
    lines = Split(text, vbLf)

    wbExpected = ""
    colE = 5
    hdrEnd = -1
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(lines(i))
        If Len(ln) = 0 Then GoTo NextHdr
        If ln = "---" Then
            hdrEnd = i
            Exit For
        End If
        If Left$(ln, 9) = "workbook" & vbTab Then wbExpected = Mid$(ln, 10)
        If Left$(ln, 9) = "column_e" & vbTab Then colE = CLng(Trim$(Mid$(ln, 10)))
NextHdr:
    Next i

    If hdrEnd < 0 Then Exit Sub
    If Len(wbExpected) = 0 Then Exit Sub
    If NormalizeWorkbookPathForCompare(wbExpected) <> NormalizeWorkbookPathForCompare(ThisWorkbook.FullName) Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_EXCLUDE_ASSIGNMENT)
    On Error GoTo CleanFail
    If ws Is Nothing Then Exit Sub

    For i = hdrEnd + 1 To UBound(lines)
        ln = lines(i)
        If Len(Trim$(ln)) = 0 Then GoTo NextData
        parts = Split(ln, vbTab)
        If UBound(parts) < 1 Then GoTo NextData
        On Error GoTo NextData
        rNum = CLng(Trim$(parts(0)))
        cellVal = DecodeBase64Utf8(Trim$(parts(1)))
        On Error GoTo CleanFail
        If rNum >= 2 And Len(cellVal) > 0 Then ws.Cells(rNum, colE).Value = cellVal
NextData:
    Next i

    On Error Resume Next
    Kill tsvPath
    jsonPath = targetDir & "\json\exclude_rules_e_column_pending.json"
    Kill jsonPath
    On Error GoTo 0

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0
    Exit Sub
CleanFail:
    On Error Resume Next
    Set stm = Nothing
    On Error GoTo 0
End Sub

Public Sub 設定_配台不要工程_AからE_TSVから反映()
    Dim targetDir As String
    Dim matrixPath As String
    Dim stm As Object
    Dim text As String
    Dim lines() As String
    Dim i As Long
    Dim hdrEnd As Long
    Dim ln As String
    Dim wbExpected As String
    Dim ws As Worksheet
    Dim parts() As String
    Dim rNum As Long
    Dim c As Long
    Dim cellTxt As String
    Dim jsonPath As String
    Dim tsvEPath As String

    On Error GoTo MatrixFail

    targetDir = ThisWorkbook.path
    If Len(targetDir) = 0 Then Exit Sub

    matrixPath = targetDir & "\log\exclude_rules_matrix_vba.tsv"
    If Len(Dir(matrixPath)) = 0 Then Exit Sub

    Set stm = CreateObject("ADODB.Stream")
    stm.charset = "UTF-8"
    stm.Open
    stm.LoadFromFile matrixPath
    text = stm.ReadText
    stm.Close
    Set stm = Nothing

    text = Replace(text, vbCrLf, vbLf)
    lines = Split(text, vbLf)

    wbExpected = ""
    hdrEnd = -1
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(lines(i))
        If Len(ln) = 0 Then GoTo NextHdrM
        If ln = "---" Then
            hdrEnd = i
            Exit For
        End If
        If Left$(ln, 9) = "workbook" & vbTab Then wbExpected = Mid$(ln, 10)
NextHdrM:
    Next i

    If hdrEnd < 0 Then Exit Sub
    If Len(wbExpected) = 0 Then Exit Sub
    If NormalizeWorkbookPathForCompare(wbExpected) <> NormalizeWorkbookPathForCompare(ThisWorkbook.FullName) Then Exit Sub

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_EXCLUDE_ASSIGNMENT)
    On Error GoTo MatrixFail
    If ws Is Nothing Then Exit Sub

    For i = hdrEnd + 1 To UBound(lines)
        ln = lines(i)
        If Len(Trim$(ln)) = 0 Then GoTo NextRowM
        parts = Split(ln, vbTab)
        If UBound(parts) < 5 Then GoTo NextRowM
        On Error GoTo NextRowM
        rNum = CLng(Trim$(parts(0)))
        On Error GoTo MatrixFail
        If rNum < 1 Then GoTo NextRowM
        For c = 1 To 5
            cellTxt = DecodeBase64Utf8(Trim$(parts(c)))
            If Len(cellTxt) = 0 Then
                ws.Cells(rNum, c).ClearContents
            Else
                ws.Cells(rNum, c).Value = cellTxt
            End If
        Next c
NextRowM:
    Next i

    On Error Resume Next
    Kill matrixPath
    tsvEPath = targetDir & "\log\exclude_rules_e_column_vba.tsv"
    Kill tsvEPath
    jsonPath = targetDir & "\json\exclude_rules_e_column_pending.json"
    Kill jsonPath
    On Error GoTo 0

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0
    Exit Sub
MatrixFail:
    On Error Resume Next
    Set stm = Nothing
    On Error GoTo 0
End Sub

Private Function ValidateMasterSkillsOpAsPriorityUnique(ByVal targetDir As String, ByRef errOut As String) As Boolean
    Dim wbPath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim openedHere As Boolean
    Dim lastCol As Long
    Dim lastRow As Long
    Dim c As Long
    Dim r As Long
    Dim pHdr As String
    Dim mHdr As String
    Dim combo As String
    Dim mem As String
    Dim cellV As String
    Dim roleCh As String
    Dim prVal As Long
    Dim okCell As Boolean
    Dim pmCount As Long
    Dim dict As Object
    Dim headerRow As Long
    Dim memCol As Long
    
    errOut = ""
    ValidateMasterSkillsOpAsPriorityUnique = False
    wbPath = targetDir & "\master.xlsm"
    If Len(Dir(wbPath)) = 0 Then
        errOut = "master.xlsm が見つかりません: " & wbPath
        Exit Function
    End If
    
    openedHere = False
    Set wb = Nothing
    On Error Resume Next
    Set wb = Workbooks("master.xlsm")
    On Error GoTo 0
    If wb Is Nothing Then
        On Error GoTo OpenFailSkills
        Set wb = Workbooks.Open(Filename:=wbPath, ReadOnly:=True, UpdateLinks:=0)
        openedHere = True
        On Error GoTo 0
    End If
    
    On Error Resume Next
    Set ws = wb.Worksheets("skills")
    On Error GoTo 0
    If ws Is Nothing Then
        errOut = "master.xlsm に「skills」シートがありません。"
        If openedHere Then
            On Error Resume Next
            wb.Close SaveChanges:=False
            On Error GoTo 0
        End If
        Exit Function
    End If
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastCol < 2 Or lastRow < 2 Then GoTo CloseOkSkills
    
    pmCount = 0
    For c = 2 To lastCol
        pHdr = Trim$(CStr(ws.Cells(1, c).Value))
        mHdr = Trim$(CStr(ws.Cells(2, c).Value))
        If Len(pHdr) > 0 And Len(mHdr) > 0 Then
            If StrComp(LCase$(pHdr), "nan", vbTextCompare) <> 0 And StrComp(LCase$(mHdr), "nan", vbTextCompare) <> 0 Then
                pmCount = pmCount + 1
            End If
        End If
    Next c
    
    If pmCount > 0 Then
        For c = 2 To lastCol
            pHdr = Trim$(CStr(ws.Cells(1, c).Value))
            mHdr = Trim$(CStr(ws.Cells(2, c).Value))
            If Len(pHdr) = 0 Or Len(mHdr) = 0 Then GoTo NextColTwo
            If StrComp(LCase$(pHdr), "nan", vbTextCompare) = 0 Or StrComp(LCase$(mHdr), "nan", vbTextCompare) = 0 Then GoTo NextColTwo
            combo = pHdr & "+" & mHdr
            Set dict = CreateObject("Scripting.Dictionary")
            For r = 3 To lastRow
                mem = Trim$(CStr(ws.Cells(r, 1).Value))
                If Len(mem) = 0 Or StrComp(LCase$(mem), "nan", vbTextCompare) = 0 Then GoTo NextRowTwo
                cellV = Trim$(CStr(ws.Cells(r, c).Value))
                okCell = ParseOpAsSkillCellForValidate(cellV, roleCh, prVal)
                If Not okCell Then GoTo NextRowTwo
                If Not dict.Exists(CStr(prVal)) Then
                    dict.Add CStr(prVal), mem & "(" & roleCh & ")"
                Else
                    errOut = "マスタ skills の優先度の数値が重複しています。" & vbCrLf & _
                        "列「" & combo & "」: 優先度 " & CStr(prVal) & " が重複（" & dict(CStr(prVal)) & " と " & mem & "(" & roleCh & ")）" & vbCrLf & _
                        "master.xlsm を修正してから再実行してください。"
                    Set dict = Nothing
                    GoTo CloseFailSkills
                End If
NextRowTwo:
            Next r
            Set dict = Nothing
NextColTwo:
        Next c
    Else
        headerRow = 1
        memCol = 1
        For c = 2 To lastCol
            combo = Trim$(CStr(ws.Cells(headerRow, c).Value))
            If Len(combo) = 0 Or StrComp(LCase$(combo), "nan", vbTextCompare) = 0 Then GoTo NextColOne
            Set dict = CreateObject("Scripting.Dictionary")
            For r = headerRow + 1 To lastRow
                mem = Trim$(CStr(ws.Cells(r, memCol).Value))
                If Len(mem) = 0 Or StrComp(LCase$(mem), "nan", vbTextCompare) = 0 Then GoTo NextRowOne
                cellV = Trim$(CStr(ws.Cells(r, c).Value))
                okCell = ParseOpAsSkillCellForValidate(cellV, roleCh, prVal)
                If Not okCell Then GoTo NextRowOne
                If Not dict.Exists(CStr(prVal)) Then
                    dict.Add CStr(prVal), mem & "(" & roleCh & ")"
                Else
                    errOut = "マスタ skills の優先度の数値が重複しています。" & vbCrLf & _
                        "列「" & combo & "」: 優先度 " & CStr(prVal) & " が重複（" & dict(CStr(prVal)) & " と " & mem & "(" & roleCh & ")）" & vbCrLf & _
                        "master.xlsm を修正してから再実行してください。"
                    Set dict = Nothing
                    GoTo CloseFailSkills
                End If
NextRowOne:
            Next r
            Set dict = Nothing
NextColOne:
        Next c
    End If

CloseOkSkills:
    If openedHere Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    ValidateMasterSkillsOpAsPriorityUnique = True
    Exit Function

CloseFailSkills:
    If openedHere Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    Exit Function

OpenFailSkills:
    errOut = "master.xlsm を開けませんでした: " & wbPath
End Function

Private Function 段階1_マスタ勤怠と機械カレンダーを同期し保護(ByVal targetDir As String) As String
    Dim wbPath As String
    Dim wb As Workbook
    Dim openedHere As Boolean
    Dim wsSkill As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim pmCount As Long
    Dim c As Long
    Dim pHdr As String
    Dim mHdr As String
    Dim r As Long
    Dim mem As String
    Dim startRow As Long
    Dim seen As Object
    Dim nOk As Long
    Dim nSkip As Long
    Dim wsMc As Worksheet
    Dim wm As Worksheet
    Dim parts As String
    
    段階1_マスタ勤怠と機械カレンダーを同期し保護 = ""
    wbPath = targetDir & "\" & MASTER_WORKBOOK_FILE
    If Len(Dir(wbPath)) = 0 Then Exit Function
    
    openedHere = False
    Set wb = Nothing
    On Error Resume Next
    Set wb = Workbooks(MASTER_WORKBOOK_FILE)
    On Error GoTo 0
    If wb Is Nothing Then
        On Error Resume Next
        Set wb = Workbooks.Open(Filename:=wbPath, ReadOnly:=True, UpdateLinks:=0)
        openedHere = Not (wb Is Nothing)
        On Error GoTo 0
    End If
    If wb Is Nothing Then
        段階1_マスタ勤怠と機械カレンダーを同期し保護 = "マスタ同期: ブックを開けませんでした"
        Exit Function
    End If
    
    On Error Resume Next
    Set wsMc = wb.Worksheets(SHEET_MACHINE_CALENDAR)
    On Error GoTo 0
    If Not wsMc Is Nothing Then
        If 段階1_マスタシートを本ブックへ置換コピー(wb, SHEET_MACHINE_CALENDAR) Then
            nOk = nOk + 1
        Else
            nSkip = nSkip + 1
        End If
    End If
    
    On Error Resume Next
    Set wsSkill = wb.Worksheets("skills")
    On Error GoTo 0
    If wsSkill Is Nothing Then GoTo CloseMasterWbSt1
    
    lastCol = wsSkill.Cells(1, wsSkill.Columns.Count).End(xlToLeft).Column
    lastRow = wsSkill.Cells(wsSkill.Rows.Count, 1).End(xlUp).Row
    If lastCol < 2 Or lastRow < 2 Then GoTo CloseMasterWbSt1
    
    pmCount = 0
    For c = 2 To lastCol
        pHdr = Trim$(CStr(wsSkill.Cells(1, c).Value))
        mHdr = Trim$(CStr(wsSkill.Cells(2, c).Value))
        If Len(pHdr) > 0 And Len(mHdr) > 0 Then
            If StrComp(LCase$(pHdr), "nan", vbTextCompare) <> 0 And StrComp(LCase$(mHdr), "nan", vbTextCompare) <> 0 Then
                pmCount = pmCount + 1
            End If
        End If
    Next c
    
    If pmCount > 0 Then
        startRow = 3
    Else
        startRow = 2
    End If
    
    Set seen = CreateObject("Scripting.Dictionary")
    For r = startRow To lastRow
        mem = Trim$(CStr(wsSkill.Cells(r, 1).Value))
        If Len(mem) = 0 Or StrComp(LCase$(mem), "nan", vbTextCompare) = 0 Then GoTo NextMemberSt1
        If seen.Exists(mem) Then GoTo NextMemberSt1
        seen.Add mem, True
        Set wm = Nothing
        On Error Resume Next
        Set wm = wb.Worksheets(mem)
        On Error GoTo 0
        If Not wm Is Nothing Then
            If 段階1_マスタシートを本ブックへ置換コピー(wb, mem) Then
                nOk = nOk + 1
            Else
                nSkip = nSkip + 1
            End If
        End If
NextMemberSt1:
    Next r
    
CloseMasterWbSt1:
    If openedHere Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    parts = "マスタ同期: 「" & SHEET_MACHINE_CALENDAR & "」+ メンバー勤怠をコピーし保護（成功シート数 " & CStr(nOk)
    If nSkip > 0 Then
        parts = parts & "・失敗 " & CStr(nSkip)
    End If
    parts = parts & "）※シート保護はマクロ終了時にまとめて適用"
    段階1_マスタ勤怠と機械カレンダーを同期し保護 = parts
End Function

Private Sub 配台_マスタSkillsから勤怠シート名を辞書に追加(ByVal targetDir As String, ByVal dict As Object)
    Dim wbPath As String
    Dim wb As Workbook
    Dim openedHere As Boolean
    Dim wsSkill As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim pmCount As Long
    Dim c As Long
    Dim pHdr As String
    Dim mHdr As String
    Dim r As Long
    Dim mem As String
    Dim startRow As Long
    Dim seen As Object
    Dim wm As Worksheet
    
    wbPath = targetDir & "\" & MASTER_WORKBOOK_FILE
    If Len(Dir(wbPath)) = 0 Then Exit Sub
    
    openedHere = False
    Set wb = Nothing
    On Error Resume Next
    Set wb = Workbooks(MASTER_WORKBOOK_FILE)
    On Error GoTo 0
    If wb Is Nothing Then
        On Error Resume Next
        Set wb = Workbooks.Open(Filename:=wbPath, ReadOnly:=True, UpdateLinks:=0)
        openedHere = Not (wb Is Nothing)
        On Error GoTo 0
    End If
    If wb Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set wsSkill = wb.Worksheets("skills")
    On Error GoTo 0
    If wsSkill Is Nothing Then GoTo CloseMasterWbProt
    
    lastCol = wsSkill.Cells(1, wsSkill.Columns.Count).End(xlToLeft).Column
    lastRow = wsSkill.Cells(wsSkill.Rows.Count, 1).End(xlUp).Row
    If lastCol < 2 Or lastRow < 2 Then GoTo CloseMasterWbProt
    
    pmCount = 0
    For c = 2 To lastCol
        pHdr = Trim$(CStr(wsSkill.Cells(1, c).Value))
        mHdr = Trim$(CStr(wsSkill.Cells(2, c).Value))
        If Len(pHdr) > 0 And Len(mHdr) > 0 Then
            If StrComp(LCase$(pHdr), "nan", vbTextCompare) <> 0 And StrComp(LCase$(mHdr), "nan", vbTextCompare) <> 0 Then
                pmCount = pmCount + 1
            End If
        End If
    Next c
    
    If pmCount > 0 Then
        startRow = 3
    Else
        startRow = 2
    End If
    
    Set seen = CreateObject("Scripting.Dictionary")
    For r = startRow To lastRow
        mem = Trim$(CStr(wsSkill.Cells(r, 1).Value))
        If Len(mem) = 0 Or StrComp(LCase$(mem), "nan", vbTextCompare) = 0 Then GoTo NextMemberProt
        If seen.Exists(mem) Then GoTo NextMemberProt
        seen.Add mem, True
        Set wm = Nothing
        On Error Resume Next
        Set wm = wb.Worksheets(mem)
        On Error GoTo 0
        If Not wm Is Nothing Then
            If Not dict.Exists(mem) Then dict.Add mem, True
        End If
NextMemberProt:
    Next r

CloseMasterWbProt:
    If openedHere Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
End Sub

Private Sub 段階1_コア実行()
    Dim wsh As Object
    Dim runBat As String
    Dim targetDir As String
    Dim wsLog As Worksheet
    Dim logFilePath As String
    Dim exitCode As Long
    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim adoStream As Object
    Dim outputText As String
    Dim logLines() As String
    Dim i As Long
    Dim warnRow As Long
    Dim st1XwErr As Long
    Dim st1XwDesc As String
    Dim missSt1 As String
    Dim st1DidUnlock As Boolean

    On Error GoTo ErrStage1
    m_lastStage1ExitCode = -1
    m_lastStage1ErrMsg = ""

    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    targetDir = ThisWorkbook.path
    If targetDir = "" Then
        m_lastStage1ErrMsg = "先にこのExcelファイルを保存してください。"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' 接続更新より先に設定シートを確保（Refresh で止まる・失敗して Exit したとき無言でシート未作成になるのを防ぐ）
    MacroSplash_SetStep "段階1: 「設定_配台不要工程」シートを確認・作成・見出しを整えています…"
    設定_配台不要工程_シートを確保
    MacroSplash_SetStep "段階1: 「設定_環境変数」シートを確認・作成し不足キーのみ追記しています…"
    設定_環境変数_シートを確保
    MacroSplash_SetStep "段階1: 「設定_シート表示」シートを確認・作成しています…"
    設定_シート表示_シートを確保
    MacroSplash_SetStep "段階1: データ接続（Power Query 等）を更新しています…"

    If Not TryRefreshWorkbookQueries() Then
        m_lastStage1ErrMsg = "データ接続の更新に失敗したため段階1を中断しました。（「設定_配台不要工程」シートは作成済みの可能性があります）"
        If Len(m_lastRefreshQueriesErrMsg) > 0 Then
            m_lastStage1ErrMsg = m_lastStage1ErrMsg & vbCrLf & m_lastRefreshQueriesErrMsg
        End If
        Application.ScreenUpdating = prevScreenUpdating
        Application.DisplayAlerts = prevDisplayAlerts
        Exit Sub
    End If

    MacroSplash_SetStep "段階1: マスタ skills の運用優先度を検証しています…"
    Dim skErrSt1 As String
    If Not ValidateMasterSkillsOpAsPriorityUnique(targetDir, skErrSt1) Then
        m_lastStage1ErrMsg = skErrSt1
        m_lastStage1ExitCode = -1
        Application.ScreenUpdating = prevScreenUpdating
        Application.DisplayAlerts = prevDisplayAlerts
        Exit Sub
    End If

    MacroSplash_SetStep "段階1: ブックを保存し LOG シートを初期化します…"
    Application.StatusBar = "ブックを保存しています..."
    DoEvents
    On Error Resume Next
    ThisWorkbook.Save
    Application.StatusBar = False
    On Error GoTo ErrStage1

    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("LOG")
    On Error GoTo ErrStage1
    If wsLog Is Nothing Then
        m_lastStage1ErrMsg = "「LOG」シートが見つかりません。"
        Application.ScreenUpdating = prevScreenUpdating
        Application.DisplayAlerts = prevDisplayAlerts
        Exit Sub
    End If

    st1DidUnlock = False
    配台マクロ_全シート保護を試行解除
    st1DidUnlock = True

    MacroSplash_SetStep "段階1: LOG シートをクリアしヘッダを書き込んでいます…"
    wsLog.Cells.Clear
    wsLog.Cells(1, 1).Value = "実行ブック: " & ThisWorkbook.FullName
    If Not GeminiCredentialsJsonPathIsConfigured() Then
        wsLog.Cells(1, 2).Value = "【要設定】シート「設定」B1 に Gemini 認証 JSON のフルパス（例: Z:\社内\gemini_credentials.json）。gemini_credentials.example.json 参照。"
    End If
    Dim st1MasterSync As String
    If Stage1SyncMasterSheetsToMacroBookEffective() Then
        MacroSplash_SetStep "段階1: master.xlsm から勤怠・機械カレンダーを同期しています…"
        st1MasterSync = 段階1_マスタ勤怠と機械カレンダーを同期し保護(targetDir)
    Else
        MacroSplash_SetStep "段階1: master 勤怠のマクロブックへコピーをスキップ（配台は master.xlsm 直読み）…"
        st1MasterSync = "マスタ同期: スキップ（STAGE1_SYNC_MASTER_SHEETS_TO_MACRO_BOOK=0。配台は master.xlsm を直接参照）"
    End If
    If Len(st1MasterSync) > 0 Then
        wsLog.Cells(1, 3).Value = st1MasterSync
    End If
    Set wsh = CreateObject("WScript.Shell")
    wsh.Environment("Process")("TASK_INPUT_WORKBOOK") = ThisWorkbook.FullName

    On Error Resume Next
    Kill targetDir & "\log\execution_log.txt"
    Kill targetDir & "\log\stage_vba_exitcode.txt"
    On Error GoTo ErrStage1

    MacroSplash_SetStep "段階1: Python（task_extract）でタスク抽出を実行しています。完了までお待ちください…（詳細は LOG シート・log\execution_log.txt）"
    If STAGE12_USE_XLWINGS_RUNPYTHON And Not STAGE12_USE_XLWINGS_SPLASH_LOG Then
        wsh.Environment("Process")("PM_AI_SPLASH_XLWINGS") = ""
        m_splashExecutionLogPath = targetDir & "\log\execution_log.txt"
        m_stageVbaExitCodeLogDir = ""
        MacroSplash_ClearExecutionLogPane
        On Error Resume Next
        Err.Clear
        XwRunConsoleRunner "run_stage1_for_xlwings"
        If Err.Number <> 0 Then
            st1XwErr = Err.Number
            st1XwDesc = Err.Description
            Err.Clear
            On Error GoTo ErrStage1
            m_splashExecutionLogPath = ""
            m_stageVbaExitCodeLogDir = ""
            m_lastStage1ExitCode = -1
            m_lastStage1ErrMsg = "段階1: xlwings RunPython が失敗しました (" & CStr(st1XwErr) & "): " & st1XwDesc
            Application.StatusBar = False
            Application.ScreenUpdating = prevScreenUpdating
            Application.DisplayAlerts = prevDisplayAlerts
            If st1DidUnlock Then 配台マクロ_対象シートを条件どおりに保護 targetDir
            Exit Sub
        End If
        On Error GoTo ErrStage1
        exitCode = ReadStageVbaExitCodeFromFile(targetDir & "\log\stage_vba_exitcode.txt")
        If exitCode = &H7FFFFFFF Then exitCode = 1
        m_splashExecutionLogPath = ""
        m_stageVbaExitCodeLogDir = ""
        m_lastStage1ExitCode = exitCode
        MacroSplash_LoadExecutionLogFromPath targetDir & "\log\execution_log.txt"
    Else
        Dim hideStage12CmdSt1 As Boolean
        hideStage12CmdSt1 = Stage12CmdHideWindowEffective()
        wsh.CurrentDirectory = Environ("TEMP")
        ' 遅延環境変数で py 終了コードを exit /b し VBA に返す（一時 .cmd を cmd.exe /c で実行）
        ' 進捗表示は execution_log.txt のポーリングのみ（PM_AI_SPLASH_XLWINGS は使わない＝二重表示防止）
        runBat = "@echo off" & vbCrLf & "setlocal EnableDelayedExpansion" & vbCrLf & "pushd """ & targetDir & """" & vbCrLf & _
                 "if not exist log mkdir log" & vbCrLf & _
                 "chcp 65001>nul" & vbCrLf & _
                 "echo [stage1] Running Python... Progress below. See also LOG sheet and log\execution_log.txt" & vbCrLf & _
                 "py -3 -u python\task_extract_stage1.py" & vbCrLf & _
                 "set STAGE1_PY_EXIT=!ERRORLEVEL!" & vbCrLf & _
                 "echo." & vbCrLf & _
                 "echo [stage1] Finished. ERRORLEVEL=!STAGE1_PY_EXIT!" & vbCrLf & _
                 "(echo !STAGE1_PY_EXIT!)>log\stage_vba_exitcode.txt" & vbCrLf & _
                 "exit /b !STAGE1_PY_EXIT!"
        m_splashExecutionLogPath = targetDir & "\log\execution_log.txt"
        m_stageVbaExitCodeLogDir = ""
        MacroSplash_ClearExecutionLogPane
        exitCode = RunTempCmdWithConsoleLayout(wsh, runBat, Not hideStage12CmdSt1, hideStage12CmdSt1)
        m_splashExecutionLogPath = ""
        m_stageVbaExitCodeLogDir = ""
        m_lastStage1ExitCode = exitCode
    End If

    MacroSplash_SetStep "段階1: 配台不要工程シートへ TSV（A?D）を反映しています…"
    ' Gemini サマリ・設定シートは planning_core が openpyxl で保存（ブックが閉じているとき）。
    ' Excel で開いたままのときは log の TSV/テキストをマクロで反映する。
    On Error Resume Next
    Call 設定_配台不要工程_AからE_TSVから反映
    MacroSplash_SetStep "段階1: 配台不要工程シートの E 列（ロジック式）を TSV から反映しています…"
    Call 設定_配台不要工程_E列_TSVから反映
    MacroSplash_SetStep "段階1: メインシートの Gemini 利用サマリ（P 列）を反映しています…"
    Call メインシート_Gemini利用サマリをP列に反映(targetDir)
    On Error GoTo ErrStage1

    logFilePath = targetDir & "\log\execution_log.txt"
    If Len(Dir(logFilePath)) = 0 Then
        wsLog.Range("A2").Value = "execution_log.txt が見つかりませんでした。exitCode=" & CStr(exitCode)
        wsLog.Range("A3").Value = "xlwings 経路（STAGE12_USE_XLWINGS_RUNPYTHON=True）では Show Console の Python 出力も参照してください。runner は planning_core 読込前に log を作成するよう修正済みです。"
        missSt1 = Trim$(GeminiReadUtf8File(targetDir & "\log\stage2_blocking_message.txt"))
        If Len(missSt1) > 0 Then
            wsLog.Range("A4").Value = "log\stage2_blocking_message.txt: " & missSt1
        End If
    Else
        Set adoStream = CreateObject("ADODB.Stream")
        adoStream.charset = "UTF-8"
        adoStream.Open
        adoStream.LoadFromFile logFilePath
        outputText = adoStream.ReadText
        adoStream.Close
        Set adoStream = Nothing
        outputText = Replace(outputText, vbCrLf, vbLf)
        logLines = Split(outputText, vbLf)
        MacroSplash_SetStep "段階1: execution_log.txt の全文を LOG シートへ書き込んでいます…（行数 " & CStr(UBound(logLines) - LBound(logLines) + 1) & "）"
        Application.ScreenUpdating = False
        For i = LBound(logLines) To UBound(logLines)
            wsLog.Cells(i + 2, 1).Value = logLines(i)
        Next i
        If exitCode <> 0 Then
            warnRow = UBound(logLines) - LBound(logLines) + 3
            If warnRow < 1 Then warnRow = 2
            wsLog.Cells(warnRow, 1).Value = "■ Pythonの終了コード: " & CStr(exitCode) & " （詳細は上記ログ参照）"
        End If
    End If

    ' Python 失敗時はこの先（取り込み・フォント）をスキップ。フォント手前まで進んでから MsgBox すると原因が誤解されやすい。
    If exitCode <> 0 Then
        Application.ScreenUpdating = prevScreenUpdating
        Application.DisplayAlerts = prevDisplayAlerts
        If st1DidUnlock Then 配台マクロ_対象シートを条件どおりに保護 targetDir
        Exit Sub
    End If

    MacroSplash_SetStep "段階1: output\plan_input_tasks.xlsx を開き「配台計画_タスク入力」へ取り込んでいます…"
    If Not ImportPlanInputTasksFromOutput(targetDir) Then
        m_lastStage1ExitCode = -1
        Application.ScreenUpdating = prevScreenUpdating
        Application.DisplayAlerts = prevDisplayAlerts
        If st1DidUnlock Then 配台マクロ_対象シートを条件どおりに保護 targetDir
        Exit Sub
    End If
    On Error Resume Next
    配台計画_タスク入力を前へ並べ替え
    On Error GoTo 0

    MacroSplash_SetStep "段階1: フォント統一と表示調整を行っています…"
    Application.ScreenUpdating = True
    DoEvents
    On Error Resume Next
    配台_全シートフォントBIZ_UDP_自動適用
    On Error GoTo 0

    On Error Resume Next
    配台計画_タスク入力_A1を選択
    On Error GoTo 0

    MacroSplash_SetStep "段階1: 「設定_シート表示」を一覧更新しブックへ適用しています…"
    On Error Resume Next
    設定_シート表示_一覧をブックから再取得
    Err.Clear
    設定_シート表示_ブックへ適用
    Err.Clear
    On Error GoTo 0

    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts
    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0
    If st1DidUnlock Then 配台マクロ_対象シートを条件どおりに保護 targetDir
    Exit Sub

ErrStage1:
    m_lastStage1ExitCode = -1
    m_lastStage1ErrMsg = "段階1: " & Err.Description
    Application.StatusBar = False
    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts
    If st1DidUnlock Then 配台マクロ_対象シートを条件どおりに保護 targetDir
End Sub

Private Function ImportPlanInputTasksFromOutput(ByVal targetDir As String) As Boolean
    Const PLAN_SHEET As String = "配台計画_タスク入力"
    Dim path As String
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim ws As Worksheet
    Dim da As Boolean
    Dim prevSUImp As Boolean
    Dim preserveFontName As String
    Dim preserveFontSize As Double
    Dim havePreserveFont As Boolean

    path = targetDir & "\output\plan_input_tasks.xlsx"
    If Len(Dir(path)) = 0 Then
        m_lastStage1ErrMsg = "plan_input_tasks.xlsx が見つかりません: " & path
        ImportPlanInputTasksFromOutput = False
        Exit Function
    End If

    da = Application.DisplayAlerts
    prevSUImp = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    MacroSplash_SetStep "段階1: plan_input_tasks.xlsx を開いています…"
    Set srcWb = Workbooks.Open(path)
    Set srcWs = srcWb.Sheets(1)

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(PLAN_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        MacroSplash_SetStep "段階1: 「配台計画_タスク入力」シートが無いため、出力ブックから新規シートとしてコピーしています…"
        srcWs.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Set ws = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        On Error Resume Next
        ws.Name = PLAN_SHEET
        On Error GoTo 0
    Else
        MacroSplash_SetStep "段階1: 既存の「配台計画_タスク入力」をクリアし、出力データを貼り付けています…"
        preserveFontName = "": preserveFontSize = 0: havePreserveFont = False
        配台計画_タスク入力_既存シートの基準フォントを取得 ws, preserveFontName, preserveFontSize, havePreserveFont
        ws.Cells.Clear
        srcWs.UsedRange.Copy Destination:=ws.Range("A1")
    End If

    MacroSplash_SetStep "段階1: 取り込み元ブックを閉じ、列幅・罫線・配台試行順番ソートを適用しています…"
    srcWb.Close SaveChanges:=False
    Set srcWb = Nothing

    On Error Resume Next
    ws.UsedRange.Columns.AutoFit
    If havePreserveFont Then
        配台計画_タスク入力_UsedRangeにフォント名とサイズを適用 ws, preserveFontName, preserveFontSize
    End If
    With ws.UsedRange.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(226, 239, 218)
    End With
    ws.UsedRange.Borders.LineStyle = 1
    ws.UsedRange.Borders.Weight = 2

    ' 配台試行順番（昇順）でソートし、オートフィルタを有効化。列が無いときのみ従来どおり「指定納期」。
    ' ※ Python は to_excel 直前に試行順で並べ替え済みだが、貼り付け後にここで一度ソートすることで
    '   UsedRange の列検出・表示を安定させ、かつ試行順を正とする（指定納期だけだと順序が崩れる）。
    Dim colTrialOrder As Long
    Dim colSpecifiedDue As Long
    Dim sortCol As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim keyRange As Range
    Dim sortDataOpt As XlSortDataOption
    colTrialOrder = FindColHeader(ws, "配台試行順番")
    colSpecifiedDue = FindColHeader(ws, "指定納期")
    sortCol = 0
    sortDataOpt = xlSortNormal
    If colTrialOrder > 0 Then
        sortCol = colTrialOrder
        sortDataOpt = xlSortTextAsNumbers
    ElseIf colSpecifiedDue > 0 Then
        sortCol = colSpecifiedDue
    End If
    If sortCol > 0 Then
        lastRow = ws.Cells(ws.Rows.Count, sortCol).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If lastRow >= 2 And lastCol >= 1 Then
            On Error Resume Next
            If ws.AutoFilterMode Then ws.AutoFilterMode = False
            On Error GoTo 0

            Set keyRange = ws.Range(ws.Cells(2, sortCol), ws.Cells(lastRow, sortCol))
            ws.Sort.SortFields.Clear
            ws.Sort.SortFields.Add Key:=keyRange, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=sortDataOpt
            With ws.Sort
                .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .Apply
            End With

            ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter
        End If
    End If
    On Error GoTo 0

    ' 上書き入力列に薄い黄色（Python planning_core と同系色）? 取り込み後も確実に付与
    配台計画_タスク入力_上書き列に入力色を付与 ws

    Application.DisplayAlerts = da
    Application.ScreenUpdating = prevSUImp
    ImportPlanInputTasksFromOutput = True
End Function

Private Sub ApplyPlanningConflictHighlightSidecar()
    Const SIDECAR As String = "planning_conflict_highlight.tsv"
    Dim p As String
    Dim adoStream As Object
    Dim txt As String
    Dim lines() As String
    Dim i As Long
    Dim ws As Worksheet
    Dim sheetName As String
    Dim numData As Long
    Dim clearCols() As String
    Dim headerMap As Object
    Dim r As Long
    Dim c As Long
    Dim colName As Variant
    Dim ci As Long
    Dim oneLine As String
    Dim parts() As String
    Dim cn As String
    Dim hv As Variant
    Dim lastRow As Long
    Dim lastCol As Long
    Dim prevSU As Boolean

    p = ThisWorkbook.path & "\log\" & SIDECAR
    If Len(Dir(p)) = 0 Then Exit Sub

    On Error GoTo CleanFail

    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.charset = "UTF-8"
    adoStream.Open
    adoStream.LoadFromFile p
    txt = adoStream.ReadText
    adoStream.Close
    Set adoStream = Nothing

    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)

    If UBound(lines) < 3 Then GoTo CleanDelete

    If Trim$(lines(0)) <> "V1" Then GoTo CleanDelete

    sheetName = Trim$(lines(1))
    numData = CLng(Val(Trim$(lines(2))))
    clearCols = Split(Trim$(lines(3)), vbTab)

    Set ws = ThisWorkbook.Sheets(sheetName)

    Set headerMap = CreateObject("Scripting.Dictionary")
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        hv = ws.Cells(1, c).Value
        If Not IsError(hv) Then
            If Not IsEmpty(hv) Then
                headerMap(Trim$(CStr(hv))) = c
            End If
        End If
    Next c

    lastRow = 1 + numData
    If lastRow < 2 Then lastRow = 2

    prevSU = Application.ScreenUpdating
    Application.ScreenUpdating = False

    ' 矛盾のないセルは段階1と同じ薄黄色へ。フォントは触らない（体裁維持）。
    ' AI解析列は黄色対象外（段階1の仕様に合わせる）
    For r = 2 To lastRow
        For Each colName In clearCols
            cn = Trim$(CStr(colName))
            If Len(cn) > 0 Then
                If headerMap.Exists(cn) Then
                    ci = headerMap(cn)
                    With ws.Cells(r, ci)
                        If StrComp(cn, "AI特別指定_解析", vbBinaryCompare) = 0 Then
                            .Interior.Pattern = xlNone
                        Else
                            .Interior.Color = RGB(255, 242, 204)
                        End If
                    End With
                End If
            End If
        Next colName
    Next r

    For i = 4 To UBound(lines)
        oneLine = Trim$(lines(i))
        If Len(oneLine) > 0 Then
            parts = Split(oneLine, vbTab, 2)
            If UBound(parts) >= 1 Then
                r = CLng(Val(parts(0)))
                cn = Trim$(parts(1))
                If r >= 2 Then
                    If headerMap.Exists(cn) Then
                        ci = headerMap(cn)
                        With ws.Cells(r, ci)
                            .Interior.Color = RGB(255, 0, 0)
                            .Font.Color = RGB(255, 255, 255)
                            .Font.Bold = True
                        End With
                    End If
                End If
            End If
        End If
    Next i

    Application.ScreenUpdating = prevSU
    Kill p
    Exit Sub

CleanDelete:
    On Error Resume Next
    Kill p
    Exit Sub

CleanFail:
    On Error Resume Next
    If Not adoStream Is Nothing Then
        adoStream.Close
        Set adoStream = Nothing
    End If
    Application.ScreenUpdating = True
End Sub

Private Sub 段階2_コア実行(Optional ByVal preserveStage1LogOnLogSheet As Boolean = False)
    Dim wsh As Object
    Dim runBat As String
    Dim outputText As String
    Dim targetDir As String
    Dim wsLog As Worksheet
    Dim logLines() As String
    Dim i As Long
    Dim adoStream As Object
    Dim logFilePath As String
    Dim cmdLogPath As String
    Dim outputFilePath As String
    Dim exitCode As Long
    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim sourceWb As Workbook
    Dim targetWb As Workbook
    Dim sourceWs As Worksheet
    Dim ws As Worksheet
    Dim sheetName As String
    Dim memberWb As Workbook
    Dim memberPath As String
    Dim newSheetName As String
    Dim planImported As Boolean
    Dim memberImported As Boolean
    Dim warnRow2 As Long
    Dim preserved As Collection
    Dim logStartRow As Long
    Dim logWriteRow As Long
    Dim lastLogR As Long
    Dim r As Long
    Dim pr As Long
    Dim st2XwErr As Long
    Dim st2XwDesc As String
    Dim missSt2 As String
    Dim st2DidUnlock As Boolean
    
    On Error GoTo ErrHandler
    
    m_lastStage2ErrMsg = ""
    m_lastStage2ExitCode = -1
    m_stage2PlanImported = False
    m_stage2MemberImported = False
    
    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    
    ' 1. 現在のExcelファイルの場所を取得 (UNCパス対応のため)
    targetDir = ThisWorkbook.path
    If targetDir = "" Then
        m_lastStage2ErrMsg = "先にこのExcelファイルを保存してください。"
        Exit Sub
    End If
    
    ' 2. 「LOG」シートが存在するか確認
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("LOG")
    On Error GoTo 0
    If wsLog Is Nothing Then
        m_lastStage2ErrMsg = "「LOG」シートが見つかりません。"
        Exit Sub
    End If

    Dim skErrSt2 As String
    If Not ValidateMasterSkillsOpAsPriorityUnique(targetDir, skErrSt2) Then
        m_lastStage2ErrMsg = skErrSt2
        Exit Sub
    End If

    設定_配台不要工程_シートを確保
    設定_環境変数_シートを確保
    設定_シート表示_シートを確保
    MacroSplash_SetStep "段階2: データ接続（Power Query 等）を更新しています…"

    If Not TryRefreshWorkbookQueries() Then
        m_lastStage2ErrMsg = "データ接続の更新に失敗したため段階2を中断しました。"
        If Len(m_lastRefreshQueriesErrMsg) > 0 Then
            m_lastStage2ErrMsg = m_lastStage2ErrMsg & vbCrLf & m_lastRefreshQueriesErrMsg
        End If
        Exit Sub
    End If
    
    MacroSplash_SetStep "段階2: LOG シートを準備しています（段階1ログの連結含む）…"
    Set preserved = New Collection
    If preserveStage1LogOnLogSheet Then
        lastLogR = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
        For r = 2 To lastLogR
            preserved.Add wsLog.Cells(r, 1).Value
        Next r
    End If
    
    ' ここでLOGシートは一旦クリア（連続実行時は直前に段階1行を退避済み）
    wsLog.Cells.Clear
    wsLog.Cells(1, 1).Value = "実行ブック: " & ThisWorkbook.FullName
    
    logStartRow = 2
    If preserveStage1LogOnLogSheet And preserved.Count > 0 Then
        wsLog.Cells(2, 1).Value = "---- 段階1（task_extract_stage1）----"
        logStartRow = 3
        For pr = 1 To preserved.Count
            wsLog.Cells(logStartRow, 1).Value = preserved(pr)
            logStartRow = logStartRow + 1
        Next pr
        wsLog.Cells(logStartRow, 1).Value = "---- 段階2（plan_simulation_stage2）----"
        logStartRow = logStartRow + 1
    End If
    
    Set wsh = CreateObject("WScript.Shell")
    
    If Not GeminiCredentialsJsonPathIsConfigured() Then
        wsLog.Cells(1, 2).Value = "【要設定】シート「設定」B1 に Gemini 認証 JSON のフルパス（例: Z:\社内\gemini_credentials.json）。gemini_credentials.example.json 参照。"
    End If

    ' タスク入力：TASK_INPUT_WORKBOOK でブックパスを渡す
    ' 段階2(plan_simulation_stage2.py) は「配台計画_タスク入力」シートを読みます
    ' 任意: シート「列設定_結果_タスク一覧」（列名・表示）で結果_タスク一覧の列順と表示/非表示を変更可。
    '       表示=FALSE の列は結果シートで列非表示。マクロ「列設定_結果_タスク一覧_チェックボックスを配置」でチェックボックスを表示列(B)に連動可能。
    wsh.Environment("Process")("TASK_INPUT_WORKBOOK") = ThisWorkbook.FullName
    
    ' ★修正：削除対象のログファイルのパスを output フォルダ配下に変更
    On Error Resume Next
    Kill targetDir & "\log\execution_log.txt"
    Kill targetDir & "\log\stage_vba_exitcode.txt"
    Kill targetDir & "\log\cmd_stage2.log"
    On Error GoTo 0
    
    ' ---------------------------------------------------------
    ' 【重要】UNCパス警告を回避する確実な方法
    ' ---------------------------------------------------------
    ' ① コマンドプロンプトが起動する瞬間の警告を防ぐため、裏で一時フォルダを指定
    wsh.CurrentDirectory = Environ("TEMP")
    
    ' 段階2: Python は TASK_INPUT_WORKBOOK のディスク上のファイルを読むため、
    ' 未保存の「配台計画_タスク入力」等が反映されないのを防ぐ
    MacroSplash_SetStep "段階2: ブックを保存しています…"
    Application.StatusBar = "ブックを保存しています..."
    DoEvents
    ThisWorkbook.Save
    Application.StatusBar = False
    
    st2DidUnlock = False
    配台マクロ_全シート保護を試行解除
    st2DidUnlock = True
    
    ' ② 「pushd」コマンドを使用し、UNCパスに一時的なドライブ文字を割り当てて確実に移動してからPythonを実行する
    ' リダイレクトは付けない（付けるとコンソールが真っ黒になる）。ログは Python が execution_log.txt にも出力する。
    cmdLogPath = targetDir & "\log\cmd_stage2.log"
    ' /v:on … py の終了コードを exit /b で返す（末尾の echo だけだと ERRORLEVEL=0 になり VBA が正常終了と誤認しがち）
    MacroSplash_SetStep "段階2: Python（plan_simulation）で計画シミュレーションを実行しています。完了までお待ちください…（詳細は LOG シート・log\execution_log.txt）"
    If STAGE12_USE_XLWINGS_RUNPYTHON And Not STAGE12_USE_XLWINGS_SPLASH_LOG Then
        wsh.Environment("Process")("PM_AI_SPLASH_XLWINGS") = ""
        m_splashExecutionLogPath = targetDir & "\log\execution_log.txt"
        m_stageVbaExitCodeLogDir = ""
        MacroSplash_ClearExecutionLogPane
        On Error Resume Next
        Err.Clear
        XwRunConsoleRunner "run_stage2_for_xlwings"
        If Err.Number <> 0 Then
            st2XwErr = Err.Number
            st2XwDesc = Err.Description
            Err.Clear
            On Error GoTo ErrHandler
            m_splashExecutionLogPath = ""
            m_stageVbaExitCodeLogDir = ""
            m_lastStage2ExitCode = -1
            m_lastStage2ErrMsg = "段階2: xlwings RunPython が失敗しました (" & CStr(st2XwErr) & "): " & st2XwDesc
            GoTo Finish
        End If
        On Error GoTo ErrHandler
        exitCode = ReadStageVbaExitCodeFromFile(targetDir & "\log\stage_vba_exitcode.txt")
        If exitCode = &H7FFFFFFF Then exitCode = 1
        m_splashExecutionLogPath = ""
        m_stageVbaExitCodeLogDir = ""
        m_lastStage2ExitCode = exitCode
        MacroSplash_LoadExecutionLogFromPath targetDir & "\log\execution_log.txt"
    Else
        Dim hideStage12CmdSt2 As Boolean
        hideStage12CmdSt2 = Stage12CmdHideWindowEffective()
        runBat = "@echo off" & vbCrLf & "setlocal EnableDelayedExpansion" & vbCrLf & "pushd """ & targetDir & """" & vbCrLf & _
                 "if not exist log mkdir log" & vbCrLf & _
                 "chcp 65001>nul" & vbCrLf & _
                 "echo [stage2] Running plan simulation... Progress below. Log file: log\execution_log.txt" & vbCrLf & _
                 "py -3 -u python\plan_simulation_stage2.py" & vbCrLf & _
                 "set STAGE2_PY_EXIT=!ERRORLEVEL!" & vbCrLf & _
                 "echo." & vbCrLf & _
                 "echo [stage2] Finished. ERRORLEVEL=!STAGE2_PY_EXIT!" & vbCrLf & _
                 "(echo !STAGE2_PY_EXIT!)>log\stage_vba_exitcode.txt" & vbCrLf
        ' コンソール表示時のみ: Python 失敗後にウィンドウがすぐ閉じないよう pause（非表示・headless では付けない）
        If Not hideStage12CmdSt2 Then
            runBat = runBat & "if not !STAGE2_PY_EXIT! equ 0 (" & vbCrLf & _
                     "echo." & vbCrLf & _
                     "echo [stage2] Python error. Press any key to close this window..." & vbCrLf & _
                     "pause" & vbCrLf & _
                     ")" & vbCrLf
        End If
        runBat = runBat & "exit /b !STAGE2_PY_EXIT!"
        ' 4. cmd 完了まで待機（execution_log を txtExecutionLog へポーリング）
        m_splashExecutionLogPath = targetDir & "\log\execution_log.txt"
        m_stageVbaExitCodeLogDir = ""
        MacroSplash_ClearExecutionLogPane
        exitCode = RunTempCmdWithConsoleLayout(wsh, runBat, Not hideStage12CmdSt2, hideStage12CmdSt2)
        m_splashExecutionLogPath = ""
        m_stageVbaExitCodeLogDir = ""
        m_lastStage2ExitCode = exitCode
    End If
    ' Python が検証エラー（例: exit 3）のとき log\stage2_blocking_message.txt に1行メッセージを残す。計画生成の MsgBox 用。
    If exitCode <> 0 Then
        Dim stage2Block As String
        stage2Block = Trim$(GeminiReadUtf8File(targetDir & "\log\stage2_blocking_message.txt"))
        If Len(stage2Block) > 0 Then
            m_lastStage2ErrMsg = stage2Block
        Else
            m_lastStage2ErrMsg = "Python の終了コードが " & CStr(exitCode) & " です。LOG シートおよび log\execution_log.txt を確認してください。（優先度重複などの検証中止時は log\stage2_blocking_message.txt も参照）"
        End If
    End If

    MacroSplash_SetStep "段階2: ログ・設定（配台不要工程・Gemini）をブックへ反映しています…"
    ' Gemini サマリ・設定シートは Python が openpyxl で保存を試みる。開きっぱなし時は log をマクロで反映。
    On Error Resume Next
    Call 設定_配台不要工程_AからE_TSVから反映
    Call 設定_配台不要工程_E列_TSVから反映
    Call メインシート_Gemini利用サマリをP列に反映(targetDir)
    On Error GoTo ErrHandler
    
    LOG_AIシートへ特別指定Geminiファイルを反映 targetDir
    
    ' 5. Python側で生成したログファイル(UTF-8)を読み込む
    logFilePath = targetDir & "\log\execution_log.txt"
    
    If Len(Dir(logFilePath)) = 0 Then
        wsLog.Cells(logStartRow, 1).Value = "execution_log.txt が見つかりませんでした。exitCode=" & CStr(exitCode)
        wsLog.Cells(logStartRow + 1, 1).Value = "xlwings 経路（STAGE12_USE_XLWINGS_RUNPYTHON=True）では Show Console の Python 出力も参照してください。runner は planning_core 読込前に log を作成するよう修正済みです。"
        missSt2 = Trim$(GeminiReadUtf8File(targetDir & "\log\stage2_blocking_message.txt"))
        If Len(missSt2) > 0 Then
            wsLog.Cells(logStartRow + 2, 1).Value = "log\stage2_blocking_message.txt: " & missSt2
        End If
        GoTo Finish
    End If
    
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.charset = "UTF-8"
    adoStream.Open
    
    adoStream.LoadFromFile logFilePath
    
    outputText = adoStream.ReadText
    adoStream.Close
    Set adoStream = Nothing
    
    ' 5. 改行コードを統一して配列に分割
    outputText = Replace(outputText, vbCrLf, vbLf)
    logLines = Split(outputText, vbLf)
    
    ' 6. LOGシートに一行ずつ書き出す（段階1退避があるときは logStartRow から）
    Application.ScreenUpdating = False
    logWriteRow = logStartRow
    For i = LBound(logLines) To UBound(logLines)
        wsLog.Cells(logWriteRow, 1).Value = logLines(i)
        logWriteRow = logWriteRow + 1
    Next i
    Application.ScreenUpdating = prevScreenUpdating
    
    If exitCode <> 0 Then
        warnRow2 = logWriteRow
        If warnRow2 < 1 Then warnRow2 = 2
        wsLog.Cells(warnRow2, 1).Value = "■ Pythonの終了コード: " & CStr(exitCode) & " （詳細は上記・実行時のコンソール・log\execution_log.txt を参照）"
    End If

    ' cmd.exe の標準出力/標準エラーも LOG シート末尾に追記（リダイレクトは環境により UTF-8 または Shift_JIS）
    If Len(Dir(cmdLogPath)) > 0 Then
        Dim cmdText As String
        Dim cmdLines() As String
        Dim baseRow As Long
        cmdText = ReadCmdCaptureLogText(cmdLogPath)
        cmdText = Replace(cmdText, vbCrLf, vbLf)
        cmdLines = Split(cmdText, vbLf)
        baseRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 2
        wsLog.Cells(baseRow, 1).Value = "---- cmd.exe stdout/stderr ----"
        For i = LBound(cmdLines) To UBound(cmdLines)
            wsLog.Cells(baseRow + 1 + i, 1).Value = cmdLines(i)
        Next i
    End If

    ' 計画検証エラー等で m_lastStage2ErrMsg が設定されたときは結果ブック取り込みをスキップ（誤って前回出力を取り込まない）
    If exitCode <> 0 And Len(m_lastStage2ErrMsg) > 0 Then
        GoTo Finish
    End If

    ' ブックが開いたままだと Python 側の openpyxl 保存が失敗することがある → TSV 経由でハイライトを反映
    ApplyPlanningConflictHighlightSidecar
    
    ' ---------------------------------------------------------
    ' 7. 生成されたExcelファイルのシートをこのブックに取り込む
    ' ---------------------------------------------------------
    MacroSplash_SetStep "段階2: 出力 xlsx から結果シート・個人シートを取り込みます…"
    planImported = False
    memberImported = False
    Set targetWb = ThisWorkbook
    
    ' 7a. production_plan_multi_day_*.xlsx（結果_* シート）
    outputFilePath = GetLatestOutputFile(targetDir & "\output", "production_plan_multi_day_*.xlsx")
    
    If outputFilePath <> "" Then
        ' 画面描画と警告を一時停止（削除確認ダイアログ等を非表示）
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        ' 列設定シートは削除せずに取り込むと同名で (2) が増えるため、事前に掃除
        列設定結果タスク一覧_番号付き重複シートを削除 targetWb
        
        ' 出力されたブックを開く
        Set sourceWb = Workbooks.Open(outputFilePath)
        
        For Each sourceWs In sourceWb.Sheets
            sheetName = Trim$(sourceWs.Name)
            
            ' Python 出力と同名のシートがマクロブックに残っていると、Copy 時に Excel が (2) を付けて複製する。
            ' 従来は「結果_*」と列設定のみ事前削除していたため、TEMP_設備毎の時間割・ブロックテーブル等が重複した。
            ' 既に残っている「名前 (2)」だけの場合もあるため、同源名（正確一致 + 「名前 (」で始まる複製）をまとめて削除する。
            マクロブックから計画取込シート同源名シートを削除 targetWb, sheetName
            Set ws = Nothing
            
            ' シートをコピー（ターゲットブックの末尾に）
            sourceWs.Copy After:=targetWb.Sheets(targetWb.Sheets.Count)
            
            ' コピーしたシートの書式設定（列幅、罫線、見出し）
            ' ※ Sheets(Count) だけだと末尾が _FontPick のとき誤参照するため、取り込み元と同名で引き直す
            Set ws = 取込ブック内のコピー先シートを取得(targetWb, sheetName)
            
            ' (1) セルフォントは上書きしない（Python 出力・ユーザーが「全シートフォント」で変更した体裁を段階2で維持する）
            
            ' (1b) 列幅: Python 出力では列幅を書かない。設備ガントは専用、それ以外は AutoFit。
            '     結果_タスク一覧 は非表示列があるため、全列 Select+AutoFit すると非表示が解除される。
            If StrComp(sheetName, "結果_設備ガント", vbBinaryCompare) = 0 Then
                結果_設備ガント_列幅を設定 ws
            ElseIf StrComp(sheetName, "結果_タスク一覧", vbBinaryCompare) = 0 Then
                結果シート_列幅_AutoFit非表示を維持 ws
                結果_タスク一覧_配完回答指定16時_いいえを強調 ws
            Else
                結果シート_列幅_AutoFit安定 ws
            End If
            
            ' (2) 使用している範囲全体に罫線(実線・細線)を引く
            ws.UsedRange.Borders.LineStyle = 1 ' xlContinuous
            ws.UsedRange.Borders.Weight = 2    ' xlThin
            ' 罫線付与で列幅が変わる環境があるため、設備ガントは専用幅を再適用
            If StrComp(sheetName, "結果_設備ガント", vbBinaryCompare) = 0 Then
                結果_設備ガント_列幅を設定 ws
            End If
            
            ' (3) 見出し（1行目）：太字・薄い黄緑（表形式シート向け）
            '     結果_設備ガント は 1 行目がレポートタイトル（Python でサイズ・背景を指定）のため上書きしない
            If StrComp(sheetName, "結果_設備ガント", vbBinaryCompare) <> 0 Then
                With ws.UsedRange.Rows(1)
                    .Font.Bold = True
                    .Interior.Color = RGB(226, 239, 218) ' 薄い黄緑色
                End With
            End If
            
            ' (3b) 結果_設備ガントのみ：タイトル A1（結合先頭）を強制的に左寄せ
            If StrComp(sheetName, "結果_設備ガント", vbBinaryCompare) = 0 Then
                結果_設備ガント_タイトルA1を左寄せに固定 ws
            End If
            
            ' (4) 結果_* のみ：メインシートへ戻るリンク（1行目・見出し行の右余白）
            If Left$(sheetName, 3) = "結果_" Then
                On Error Resume Next
                結果シート_メインへ戻るリンクを付与 ws
                Err.Clear
                On Error GoTo ErrHandler
            End If
            
            ' (5) 結果_* の保護は段階2 終了時（Finish）の 配台マクロ_対象シートを条件どおりに保護 でまとめて適用（処理中は全シート解除済み）
            
        Next sourceWs
        
        ' (6) master.xlsm メインの工場稼働(A12/B12)・定常(A15/B15)を結果_設備毎の時間割・結果_設備毎の時間割_機械名毎・結果_設備ガントに反映（UserInterfaceOnly 保護後もマクロから可。依頼NO薄緑は機械名毎のみ追加）
        On Error Resume Next
        取込後_結果シートへマスタ時刻を反映 targetWb
        Err.Clear
        ' マスタ反映・保護後も、設備ガントの列幅は専用設定に戻す（AutoFit 混入防止）
        Set ws = Nothing
        Set ws = targetWb.Worksheets("結果_設備ガント")
        If Err.Number = 0 Then
            結果_設備ガント_列幅を設定 ws
            結果_設備ガント_タイトルA1を左寄せに固定 ws
        End If
        Err.Clear
        On Error GoTo ErrHandler
        
        ' ソースブックを閉じる（保存しない）
        sourceWb.Close SaveChanges:=False
        Set sourceWb = Nothing
        
        planImported = True
        
        ' 画面描画と警告を元に戻す
        Application.DisplayAlerts = prevDisplayAlerts
        Application.ScreenUpdating = prevScreenUpdating
        
        ' 最初（一番左）のシートを選択状態にする（お好みで）
        targetWb.Sheets(1).Activate
    End If
    
    MacroSplash_SetStep "段階2: 個人別スケジュール（member_schedule）を取り込んでいます…"
    ' 7b. member_schedule_*.xlsx（メンバー名シート → 個人_プレフィックスで取り込み）
    memberPath = GetLatestOutputFile(targetDir & "\output", "member_schedule_*.xlsx")
    If Len(memberPath) > 0 Then
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        Set memberWb = Workbooks.Open(memberPath)
        
        For Each sourceWs In memberWb.Sheets
            sheetName = sourceWs.Name
            newSheetName = SafePersonalSheetName(sheetName)
            
            ' 既に「個人_*」シートがある場合は削除せず、内容をクリアしてから上書き
            On Error Resume Next
            Set ws = targetWb.Sheets(newSheetName)
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo ErrHandler
                sourceWs.Copy After:=targetWb.Sheets(targetWb.Sheets.Count)
                Set ws = 取込ブック内のコピー先シートを取得(targetWb, sheetName)
                On Error Resume Next
                ws.Name = newSheetName
                On Error GoTo ErrHandler
            Else
                Err.Clear
                On Error GoTo ErrHandler
                ws.Cells.Clear
                sourceWs.UsedRange.Copy Destination:=ws.Range("A1")
            End If
            
            ' 個人_* もセルフォントは上書きしない（同上）
            ' 個人シートの列幅も Python 側では設定しない（同上 AutoFit）
            結果シート_列幅_AutoFit安定 ws
            ws.UsedRange.Borders.LineStyle = 1
            ws.UsedRange.Borders.Weight = 2
            With ws.UsedRange.Rows(1)
                .Font.Bold = True
                .Interior.Color = RGB(226, 239, 218)
            End With
        Next sourceWs
        
        memberWb.Close SaveChanges:=False
        Set memberWb = Nothing
        memberImported = True
        
        Application.DisplayAlerts = prevDisplayAlerts
        Application.ScreenUpdating = prevScreenUpdating
    End If
    
    MacroSplash_SetStep "段階2: メインシート・シート順・フォント後処理を実行しています…"
    ' メインシート：メンバーへのリンク ＋ 前日から12日間の出退勤（失敗しても本処理は継続）
    On Error Resume Next
    メインシート_メンバー一覧と出勤表示 True
    ' 個人_* シートをブック末尾へ（失敗しても継続）
    個人シートを末尾へ並べ替え
    ' 「設定」の一つ前に列設定シートを置く（取り込みでは末尾に付くため）
    On Error Resume Next
    列設定_結果_タスク一覧を設定の直前へ移動 ThisWorkbook
    On Error GoTo ErrHandler

    MacroSplash_SetStep "段階2: 「設定_シート表示」を一覧更新しブックへ適用しています…"
    On Error Resume Next
    設定_シート表示_一覧をブックから再取得
    Err.Clear
    設定_シート表示_ブックへ適用
    Err.Clear
    On Error GoTo ErrHandler
    
    ' 完了ダイアログ直前はメインシートを表示（A1）
    On Error Resume Next
    Application.ScreenUpdating = True
    メインシートA1を選択
    DoEvents
    On Error GoTo ErrHandler
    
    m_stage2PlanImported = planImported
    m_stage2MemberImported = memberImported

    On Error Resume Next
    配台_全シートフォントBIZ_UDP_自動適用
    On Error GoTo 0

Finish:
    On Error Resume Next
    If Not adoStream Is Nothing Then
        adoStream.Close
        Set adoStream = Nothing
    End If
    If Not sourceWb Is Nothing Then
        sourceWb.Close SaveChanges:=False
        Set sourceWb = Nothing
    End If
    If Not memberWb Is Nothing Then
        memberWb.Close SaveChanges:=False
        Set memberWb = Nothing
    End If
    On Error GoTo 0
    
    Application.DisplayAlerts = prevDisplayAlerts
    Application.ScreenUpdating = prevScreenUpdating
    
    If st2DidUnlock Then
        On Error Resume Next
        配台マクロ_対象シートを条件どおりに保護 targetDir
        On Error GoTo 0
    End If
    
    On Error Resume Next
    If planImported Then
        結果プレフィックスシートの表示倍率を設定 ThisWorkbook, 100
        結果_設備ガント_表示倍率を設定 ThisWorkbook, 85
        結果_設備毎の時間割_B2選択して窓枠固定
        結果_タスク一覧_F2選択して窓枠固定
        結果_カレンダー出勤簿_A2選択して窓枠固定
    End If
    メインシートA1を選択
    On Error GoTo 0
    
    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0
    
    Exit Sub
    
ErrHandler:
    m_lastStage2ErrMsg = "VBAエラー: " & Err.Number & " / " & Err.Description
    If Not wsLog Is Nothing Then
        wsLog.Cells(1, 1).Value = m_lastStage2ErrMsg
    End If
    Resume Finish
End Sub

Function GetLatestOutputFile(folderPath As String, filePattern As String) As String
    Dim latestPath As String
    Dim latestDate As Date

    If Len(Dir(folderPath, vbDirectory)) = 0 Then
        GetLatestOutputFile = ""
        Exit Function
    End If

    latestDate = 0
    latestPath = ""
    CollectLatestOutputFileRecursive folderPath, filePattern, latestPath, latestDate
    GetLatestOutputFile = latestPath
End Function

Private Sub CollectLatestOutputFileRecursive(ByVal folderPath As String, ByVal filePattern As String, ByRef latestPath As String, ByRef latestDate As Date)
    Dim fso As Object
    Dim fldr As Object
    Dim subFldr As Object
    Dim fil As Object
    Dim currDate As Date
    If Len(folderPath) = 0 Then Exit Sub

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Sub
    If Not fso.FolderExists(folderPath) Then Exit Sub
    Set fldr = fso.GetFolder(folderPath)
    If fldr Is Nothing Then Exit Sub
    On Error GoTo 0

    For Each fil In fldr.Files
        If LCase$(fil.Name) Like LCase$(filePattern) Then
            On Error Resume Next
            currDate = fil.DateLastModified
            If Err.Number = 0 Then
                If currDate > latestDate Then
                    latestDate = currDate
                    latestPath = CStr(fil.path)
                End If
            Else
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next fil

    For Each subFldr In fldr.SubFolders
        CollectLatestOutputFileRecursive CStr(subFldr.path), filePattern, latestPath, latestDate
    Next subFldr
End Sub

