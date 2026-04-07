Private Function ParseStage12CmdHideWindowBool(ByVal s As String, ByVal defaultVal As Boolean) As Boolean
    Dim t As String
    t = LCase$(Trim$(s))
    If Len(t) = 0 Then ParseStage12CmdHideWindowBool = defaultVal: Exit Function
    If t = "1" Or t = "true" Or t = "yes" Or t = "on" Or t = "y" Then
        ParseStage12CmdHideWindowBool = True
        Exit Function
    End If
    If t = "0" Or t = "false" Or t = "no" Or t = "off" Or t = "n" Then
        ParseStage12CmdHideWindowBool = False
        Exit Function
    End If
    If Trim$(s) = "はい" Then ParseStage12CmdHideWindowBool = True: Exit Function
    If Trim$(s) = "いいえ" Then ParseStage12CmdHideWindowBool = False: Exit Function
    ParseStage12CmdHideWindowBool = defaultVal
End Function

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

Private Function ReadTextFileWithCharset(ByVal filePath As String, ByVal charset As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.charset = charset
    stm.Open
    stm.LoadFromFile filePath
    ReadTextFileWithCharset = stm.ReadText
    stm.Close
    Set stm = Nothing
End Function

' cmd.exe が生成した capture ログ用（UTF-8 BOM が無ければ日本語環境では Shift_JIS として読む）
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

' Excel で式として解釈される先頭 "=" を文字列として保持する
Private Function EscapeExcelFormulaText(ByVal s As String) As String
    If Len(s) > 0 Then
        If Left$(s, 1) = "=" Then
            EscapeExcelFormulaText = "'" & s
            Exit Function
        End If
    End If
    EscapeExcelFormulaText = s
End Function

' 段階2 完了後: 特別指定_備考用 Gemini のプロンプト・応答ログを LOG_AI シートに転記（pause の代わりにブック内で確認）
Private Function NormalizeWorkbookPathForCompare(ByVal p As String) As String
    NormalizeWorkbookPathForCompare = LCase$(Replace(Replace(Trim$(p), "/", "\"), vbTab, ""))
End Function

' NodeTypedValue は Variant（Byte 配列）のため、引数は Variant にしてから Byte() へ代入する。
Private Function Utf8BytesToString(ByVal data As Variant) As String
    Dim stm As Object
    Dim bytes() As Byte
    On Error GoTo CleanFail
    bytes = data
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.Write bytes
    stm.Position = 0
    stm.Type = 2
    stm.charset = "UTF-8"
    Utf8BytesToString = stm.ReadText
    stm.Close
    Exit Function
CleanFail:
    Utf8BytesToString = ""
End Function

Private Function DecodeBase64Utf8(ByVal b64 As String) As String
    On Error GoTo Fail
    Dim xml As Object, node As Object
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.text = b64
    DecodeBase64Utf8 = Utf8BytesToString(node.NodeTypedValue)
    Exit Function
Fail:
    On Error GoTo Fail2
    Set xml = CreateObject("MSXML2.DOMDocument.3.0")
    Set node = xml.createElement("b64")
    node.DataType = "bin.base64"
    node.text = b64
    DecodeBase64Utf8 = Utf8BytesToString(node.NodeTypedValue)
    Exit Function
Fail2:
    DecodeBase64Utf8 = ""
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

' =========================================================
' Python が出力した log\exclude_rules_matrix_vba.tsv から
' 「設定_配台不要工程」の A?E を書き込む（ブックが Excel で開いたまま openpyxl が保存できないとき）。
' =========================================================
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

' =========================================================
' マスタ master.xlsm「skills」: 同一列で OP/AS の優先度の数値が重複していないか検証
' （planning_core._validate_skills_op_as_priority_numbers_unique と同趣旨・2段/1段ヘッダ両対応）
' =========================================================
Private Function ParseOpAsSkillCellForValidate(ByVal s As String, ByRef roleOut As String, ByRef prOut As Long) As Boolean
    Dim t As String
    Dim tail As String
    ParseOpAsSkillCellForValidate = False
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
    ParseOpAsSkillCellForValidate = True
End Function
