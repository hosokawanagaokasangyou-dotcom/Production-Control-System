Attribute VB_Name = "Gemini連携"
Option Explicit

Sub アニメ付き_Gemini認証を暗号化してB1に保存()
    Call AnimateButtonPush
    ' InputBox 等があるためグリッド操作ブロックは使わない（スプラッシュのみ）
    アニメ付き_スプラッシュ付きで実行 "Gemini 認証を暗号化して保存しています…", "設定_Gemini認証を暗号化してB1に保存", , , False
End Sub

Public Sub 設定_Gemini認証を暗号化してB1に保存()
    Dim apiKey As String
    Dim pass1 As String
    Dim pass2 As String
    Dim wbPath As String
    Dim outPath As String
    Dim plainPath As String
    Dim passPath As String
    Dim errPath As String
    Dim jsonBody As String
    Dim wsh As Object
    Dim gemBat As String
    Dim exitCode As Long
    Dim wsSet As Worksheet
    Dim errLog As String
    Dim pyScript As String
    
    On Error GoTo EH
    
    Set wsSet = Nothing
    On Error Resume Next
    Set wsSet = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    On Error GoTo EH
    If wsSet Is Nothing Then
        MsgBox "シート「" & SHEET_SETTINGS & "」がありません。先に作成してください。", vbExclamation
        Exit Sub
    End If
    
    wbPath = ThisWorkbook.path
    If Len(wbPath) = 0 Then
        MsgBox "ブックを一度保存してから実行してください（保存フォルダに暗号化 JSON を出力します）。", vbExclamation
        Exit Sub
    End If
    
    pyScript = wbPath & "\python\encrypt_gemini_credentials.py"
    If Len(Dir(pyScript)) = 0 Then
        MsgBox "次のファイルが見つかりません。" & vbCrLf & pyScript & vbCrLf & vbCrLf & _
               "テストコード直下に python\ フォルダがあり、上記スクリプトがあるか確認してください。", vbCritical
        Exit Sub
    End If
    
    apiKey = InputBox( _
        "Gemini API キー（AIza...）を貼り付けてください。" & vbCrLf & _
        "キャンセルで中断します。", _
        "Gemini 認証の暗号化 (1/3)")
    If Len(Trim$(apiKey)) = 0 Then Exit Sub
    
    pass1 = InputBox( _
        "暗号化に使うパスフレーズを入力してください。" & vbCrLf & _
        "社内で案内されている値を使用し、次の画面でもう一度同じものを入力します。", _
        "Gemini 認証の暗号化 (2/3)")
    If Len(pass1) = 0 Then
        MsgBox "パスフレーズが空のため中断しました。", vbInformation
        Exit Sub
    End If
    
    pass2 = InputBox( _
        "パスフレーズをもう一度入力してください（確認用）。", _
        "Gemini 認証の暗号化 (3/3)")
    If StrComp(pass1, pass2, vbBinaryCompare) <> 0 Then
        MsgBox "2回のパスフレーズが一致しません。やり直してください。", vbExclamation
        Exit Sub
    End If
    
    Randomize
    plainPath = Environ("TEMP") & "\gemini_plain_" & Format(Now, "yyyymmddhhnnss") & "_" & CStr(Int(1000000 * Rnd)) & ".json"
    passPath = Environ("TEMP") & "\gemini_pass_" & Format(Now, "yyyymmddhhnnss") & "_" & CStr(Int(1000000 * Rnd)) & ".txt"
    errPath = Environ("TEMP") & "\gemini_encrypt_stderr.txt"
    outPath = wbPath & "\gemini_credentials.encrypted.json"
    
    If Len(Dir(outPath)) > 0 Then
        If MsgBox("既に次のファイルがあります。上書きしますか？" & vbCrLf & outPath, vbYesNo Or vbExclamation, "確認") <> vbYes Then
            Exit Sub
        End If
    End If
    
    jsonBody = "{" & """gemini_api_key"": """ & GeminiJsonStringEscape(Trim$(apiKey)) & """}"
    Call GeminiWriteUtf8File(plainPath, jsonBody)
    Call GeminiWriteUtf8File(passPath, pass1)
    
    On Error Resume Next
    Kill errPath
    On Error GoTo EH
    
    MacroSplash_SetStep "Gemini: Python で認証 JSON を暗号化しています…"
    Set wsh = CreateObject("WScript.Shell")
    gemBat = "@echo off" & vbCrLf & "pushd """ & wbPath & """" & vbCrLf & "chcp 65001>nul" & vbCrLf & _
             "py -3 -u python\encrypt_gemini_credentials.py """ & plainPath & """ """ & outPath & """ --passphrase-file """ & passPath & """ 2> """ & errPath & """" & vbCrLf & _
             "exit /b %ERRORLEVEL%"
    exitCode = RunTempCmdWithConsoleLayout(wsh, gemBat)
    
    On Error Resume Next
    Kill plainPath
    Kill passPath
    On Error GoTo EH
    
    If Len(Dir(outPath)) = 0 Then
        errLog = Trim$(GeminiReadUtf8File(errPath))
        If Len(errLog) > 2500 Then errLog = Left$(errLog, 2500) & vbCrLf & "…（省略）"
        If Len(errLog) = 0 Then errLog = "（標準エラーに出力なし。py -3 が PATH に無い、または別のエラーの可能性があります）"
        MsgBox "暗号化ファイルができませんでした。（終了コード " & CStr(exitCode) & "）" & vbCrLf & vbCrLf & _
               "【Python のメッセージ】" & vbCrLf & errLog & vbCrLf & vbCrLf & _
               "よくある対処: py -3 -m pip install cryptography" & vbCrLf & _
               "または: py -3 -m pip install -r python\requirements.txt", vbCritical
        Exit Sub
    End If
    
    wsSet.Range("B1").Value = outPath
    wsSet.Range("B2").ClearContents
    
    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0
    
    MacroSplash_SetStep "Gemini 認証の暗号化が完了しました。設定 B1 にパスを保存しました。"
    m_animMacroSucceeded = True
    Exit Sub
EH:
    MsgBox "エラー: " & Err.Description, vbCritical
End Sub

Public Sub メインシート_Gemini利用サマリをP列に反映(ByVal targetDir As String)
    Const START_ROW As Long = 16
    Const USAGE_COL As Long = 16 ' P
    Const CLEAR_ROWS As Long = 120
    
    Dim wsMain As Worksheet
    Dim fp As String
    Dim adoStream As Object
    Dim outputText As String
    Dim logLines() As String
    Dim i As Long
    Dim r As Long
    Dim lastClearRow As Long
    
    Set wsMain = GetMainWorksheet()
    If wsMain Is Nothing Then Exit Sub
    
    lastClearRow = START_ROW + CLEAR_ROWS - 1
    wsMain.Range(wsMain.Cells(START_ROW, USAGE_COL), wsMain.Cells(lastClearRow, USAGE_COL)).ClearContents
    
    fp = targetDir & "\log\gemini_usage_summary_for_main.txt"
    If Len(Dir(fp)) = 0 Then Exit Sub
    
    On Error GoTo GeminiUsageP_Fail
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.charset = "UTF-8"
    adoStream.Open
    adoStream.LoadFromFile fp
    outputText = adoStream.ReadText
    adoStream.Close
    Set adoStream = Nothing
    On Error GoTo 0
    
    outputText = Replace(outputText, vbCrLf, vbLf)
    If Len(Trim$(outputText)) = 0 Then Exit Sub
    
    logLines = Split(outputText, vbLf)
    
    Application.ScreenUpdating = False
    For i = LBound(logLines) To UBound(logLines)
        r = START_ROW + i
        If r > lastClearRow Then Exit For
        With wsMain.Cells(r, USAGE_COL)
            .Value = logLines(i)
            .WrapText = True
            .VerticalAlignment = xlTop
        End With
    Next i
    Application.ScreenUpdating = True
    Exit Sub
    
GeminiUsageP_Fail:
    On Error Resume Next
    If Not adoStream Is Nothing Then
        adoStream.Close
        Set adoStream = Nothing
    End If
    On Error GoTo 0
End Sub

Private Function GeminiCredentialsJsonPathIsConfigured() As Boolean
    Dim rng As Range
    GeminiCredentialsJsonPathIsConfigured = False
    On Error Resume Next
    Set rng = ThisWorkbook.Worksheets(SHEET_SETTINGS).Range("B1")
    If Err.Number = 0 And Not rng Is Nothing Then
        If Len(Trim$(CStr(rng.Value))) > 0 Then
            GeminiCredentialsJsonPathIsConfigured = True
        End If
    End If
    On Error GoTo 0
End Function

Private Sub LOG_AIシートへ特別指定Geminiファイルを反映(ByVal targetDir As String)
    Const SH_LOG_AI As String = "LOG_AI"
    Const MAX_CELL As Long = 32700
    Dim ws As Worksheet
    Dim wasProtected As Boolean
    Dim promptPath As String
    Dim remarkPath As String
    Dim fileBody As String
    Dim lines() As String
    Dim r As Long
    Dim i As Long
    
    promptPath = targetDir & "\log\ai_task_special_last_prompt.txt"
    remarkPath = targetDir & "\log\ai_task_special_remark_last.txt"
    
    ' ※呼び出し元で On Error Resume Next の直後だと Err が残っていることがある。
    ' Set ws = Worksheets(...) 成功時も Err は自動クリアされないため、
    ' Err.Number 判定で「無い」と誤認し別シートへ書くと LOG_AI が空のままになる。
    Set ws = Nothing
    Err.Clear
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_LOG_AI)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        On Error Resume Next
        ws.Name = SH_LOG_AI
        On Error GoTo 0
    End If
    If ws Is Nothing Then Exit Sub

    ' 保護シートだと Cells(...).Value で 1004 になるため、書き込み前に解除（再保護はしない）
    wasProtected = ws.ProtectContents
    If wasProtected Then
        On Error Resume Next
        ws.Unprotect
        On Error GoTo 0
        If ws.ProtectContents Then
            MsgBox "LOG_AI シートが保護されているため、AIログを書き込めません。保護を解除してください。", vbExclamation
            Exit Sub
        End If
    End If
    
    ws.Cells.Clear
    r = 1
    
    ws.Cells(r, 1).Value = "[log\ai_task_special_last_prompt.txt]"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
    If Len(Dir(promptPath)) > 0 Then
        fileBody = ReadTextFileWithCharset(promptPath, "utf-8")
        fileBody = Replace(fileBody, vbCrLf, vbLf)
        lines = Split(fileBody, vbLf)
        For i = LBound(lines) To UBound(lines)
            If Len(lines(i)) > MAX_CELL Then
                ws.Cells(r, 1).Value = EscapeExcelFormulaText(Left$(lines(i), MAX_CELL) & "…(切り詰め)")
            Else
                ws.Cells(r, 1).Value = EscapeExcelFormulaText(lines(i))
            End If
            r = r + 1
        Next i
    Else
        ws.Cells(r, 1).Value = "(ファイルなし: " & promptPath & ")"
        r = r + 1
    End If
    
    r = r + 1
    ws.Cells(r, 1).Value = "[log\ai_task_special_remark_last.txt]"
    ws.Cells(r, 1).Font.Bold = True
    r = r + 1
    If Len(Dir(remarkPath)) > 0 Then
        fileBody = ReadTextFileWithCharset(remarkPath, "utf-8")
        fileBody = Replace(fileBody, vbCrLf, vbLf)
        lines = Split(fileBody, vbLf)
        For i = LBound(lines) To UBound(lines)
            If Len(lines(i)) > MAX_CELL Then
                ws.Cells(r, 1).Value = EscapeExcelFormulaText(Left$(lines(i), MAX_CELL) & "…(切り詰め)")
            Else
                ws.Cells(r, 1).Value = EscapeExcelFormulaText(lines(i))
            End If
            r = r + 1
        Next i
    Else
        ws.Cells(r, 1).Value = "(ファイルなし: " & remarkPath & ")"
        r = r + 1
    End If
    
    ws.Columns(1).ColumnWidth = 100
End Sub

