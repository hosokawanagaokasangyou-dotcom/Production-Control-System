Attribute VB_Name = "業務ロジック"
Option Explicit

Private Function AugmentCmdBodyWithConsoleTitle(ByVal body As String, ByVal titleText As String) As String
    Const echoOffCrLf As String = "@echo off" & vbCrLf
    If Len(body) >= Len(echoOffCrLf) And LCase$(Left$(body, 9)) = "@echo off" Then
        If Mid$(body, 10, 2) = vbCrLf Then
            AugmentCmdBodyWithConsoleTitle = echoOffCrLf & "title " & titleText & vbCrLf & Mid$(body, 12)
            Exit Function
        End If
    End If
    AugmentCmdBodyWithConsoleTitle = echoOffCrLf & "title " & titleText & vbCrLf & body
End Function

Private Function EnsureStageBatchStdoutRedirect(ByVal body As String) As String
    Dim t As String
    Dim lines() As String
    Dim i As Long
    Dim s As String
    t = Replace(Replace(body, vbCrLf, vbLf), vbCr, vbLf)
    lines = Split(t, vbLf)
    For i = LBound(lines) To UBound(lines)
        s = lines(i)
        If Len(s) > 0 Then
            If InStr(1, LTrim$(s), "py ", vbTextCompare) = 1 Then
                If InStr(1, s, "1>>", vbTextCompare) = 0 And InStr(1, s, ">nul", vbTextCompare) = 0 Then
                    lines(i) = RTrim$(s) & " 1>nul 2>&1"
                End If
                EnsureStageBatchStdoutRedirect = Join(lines, vbCrLf)
                Exit Function
            End If
        End If
    Next i
    EnsureStageBatchStdoutRedirect = body
End Function

Private Sub ConsoleApplyBorderlessIfNeeded(ByVal hwnd As LongPtr)
    On Error Resume Next
    If Not STAGE12_CMD_OVERLAY_BORDERLESS Then Exit Sub
    #If Win64 Then
        Dim ns64 As LongPtr
        Dim nsLo As Long
        ns64 = SplashGetWindowLongPtr(hwnd, GWL_STYLE)
        nsLo = CLng(ns64)
        nsLo = nsLo And Not WS_CONSOLE_OVERLAY_STRIP
        Call SplashSetWindowLongPtr(hwnd, GWL_STYLE, nsLo)
    #Else
        Dim ns As Long
        ns = SplashGetWindowLongPtr(hwnd, GWL_STYLE)
        Call SplashSetWindowLongPtr(hwnd, GWL_STYLE, ns And Not WS_CONSOLE_OVERLAY_STRIP)
    #End If
    Call SetWindowPos(hwnd, 0&, 0&, 0&, 0&, 0&, SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED)
    On Error GoTo 0
End Sub

Private Sub ConsoleApplyBorderlessIfNeeded(ByVal hwnd As Long)
End Sub

Sub アニメ付き_計画生成を実行()
    Call AnimateButtonPush
    アニメ付き_スプラッシュ付きで実行 "シミュレーション（計画生成）を実行しています…", "RunPython", False, , True, True
End Sub

Sub アニメ付き_タスク抽出を実行()
    Call AnimateButtonPush
    アニメ付き_スプラッシュ付きで実行 "タスク抽出（段階1）を実行しています…", "RunPythonStage1", , , True, True
End Sub

Sub アニメ付き_段階1と段階2を連続実行()
    Call AnimateButtonPush
    アニメ付き_スプラッシュ付きで実行 "段階1と段階2を連続実行しています…", "RunPythonStage1ThenStage2", , , True, True
End Sub

Sub アニメ付き_環境構築を実行()
    Const ENV_BUILD_PASSWORD As String = "1111"
    Dim userInput As String
    
    ' 誤操作防止用（セキュリティ目的ではないため、パスワードを明示）
    userInput = InputBox( _
        "環境構築は初回のみ実行してください。" & vbCrLf & vbCrLf & _
        "Python 3 が無ければインストールし、setup_environment.py で requirements.txt を導入します。" & vbCrLf & _
        "　pandas / openpyxl / google-genai / cryptography / xlwings 等" & vbCrLf & _
        "　xlwings の Excel アドイン（xlwings.xlam）も配置します。" & vbCrLf & vbCrLf & _
        "誤操作防止のため、下記パスワードを入力してから OK を押してください。" & vbCrLf & _
        "【パスワード】" & ENV_BUILD_PASSWORD & vbCrLf & vbCrLf & _
        "キャンセルすると実行しません。", _
        "環境構築の実行確認")
    
    If StrComp(Trim$(userInput), ENV_BUILD_PASSWORD, vbBinaryCompare) <> 0 Then
        MsgBox "パスワードが一致しないため、環境構築は実行されませんでした。", vbInformation
        Exit Sub
    End If
    
    Call AnimateButtonPush
    アニメ付き_スプラッシュ付きで実行 "環境構築を実行しています…", "InstallComponents"
End Sub

Sub アニメ付き_列設定_結果_タスク一覧_列順表示をPython適用()
    Call AnimateButtonPush
    アニメ付き_スプラッシュ付きで実行 "列設定を結果タスク一覧に反映しています…", "列設定_結果_タスク一覧_列順表示をPython適用"
End Sub

Sub アニメ付き_列設定_結果_タスク一覧_重複列名を整理()
    Call AnimateButtonPush
    アニメ付き_スプラッシュ付きで実行 "列設定シートの重複列名を整理しています…", "列設定_結果_タスク一覧_重複列名を整理"
End Sub

Sub メインシート_メンバー一覧と出勤表示_手動()
    メインシート_メンバー一覧と出勤表示 False
End Sub

Sub アニメ付き_メインシート_masterブックを開く()
    Call AnimateButtonPush
    メインシート_masterブックを開く
End Sub

Public Sub メインシート_master開くボタンを配置()
    Dim ws As Worksheet
    
    Set ws = GetMainWorksheet()
    If ws Is Nothing Then
        MsgBox "「メイン」「Main」、または名前に「メイン」を含むシートが見つかりません。", vbExclamation
        Exit Sub
    End If
    ws.Activate
    CreateCoolButtonWithPreset "master.xlsm を開く", "アニメ付き_メインシート_masterブックを開く", 380, 12, 2
    MsgBox "メインシートにボタンを配置しました。位置はドラッグで調整できます。", vbInformation
End Sub

Private Function マスタメイン_セルを時刻Dateへ(ByVal v As Variant, ByRef outT As Date) As Boolean
    On Error GoTo Fail
    If IsEmpty(v) Or VarType(v) = vbError Then GoTo Fail
    
    Select Case VarType(v)
    Case vbDate
        outT = CDate(v)
        マスタメイン_セルを時刻Dateへ = True
        Exit Function
    Case vbDouble, vbSingle, vbCurrency, vbLong, vbInteger
        outT = CDate(CDbl(v))
        マスタメイン_セルを時刻Dateへ = True
        Exit Function
    Case vbString
        If Len(Trim$(v)) = 0 Then GoTo Fail
        outT = CDate(Trim$(v))
        マスタメイン_セルを時刻Dateへ = True
        Exit Function
    Case Else
        If IsDate(v) Then
            outT = CDate(v)
            マスタメイン_セルを時刻Dateへ = True
            Exit Function
        End If
    End Select
Fail:
    マスタメイン_セルを時刻Dateへ = False
End Function

Private Function 時刻を分に(ByVal t As Date) As Long
    時刻を分に = CLng(Hour(t)) * 60& + CLng(Minute(t))
End Function

Private Function 半開区間が重なる分(ByVal a0 As Long, ByVal a1 As Long, ByVal b0 As Long, ByVal b1 As Long) As Boolean
    半開区間が重なる分 = (a0 < b1) And (a1 > b0)
End Function

Private Function 日時帯文字列を時刻範囲に(ByVal v As Variant, ByRef t0 As Date, ByRef t1 As Date) As Boolean
    Dim s As String
    Dim sep As String
    Dim parts() As String
    Dim leftS As String
    Dim rightS As String
    
    If IsEmpty(v) Or VarType(v) = vbError Then Exit Function
    s = Trim$(Replace(Replace(CStr(v), vbCr, ""), vbLf, ""))
    If Len(s) = 0 Then Exit Function
    If InStr(s, "■") > 0 Then Exit Function
    
    sep = vbNullString
    If InStr(s, "-") > 0 Then sep = "-"
    If InStr(s, "－") > 0 Then sep = "－"
    If Len(sep) = 0 And InStr(s, "~") > 0 Then sep = "~"
    If Len(sep) = 0 And InStr(s, "?") > 0 Then sep = "?"
    If Len(sep) = 0 Then Exit Function
    
    parts = Split(s, sep, 2)
    If UBound(parts) < 1 Then Exit Function
    leftS = Trim$(Replace(parts(0), "：", ":"))
    rightS = Trim$(Replace(parts(1), "：", ":"))
    
    If Not マスタメイン_セルを時刻Dateへ(leftS, t0) Then Exit Function
    If Not マスタメイン_セルを時刻Dateへ(rightS, t1) Then Exit Function
    If 時刻を分に(t0) >= 時刻を分に(t1) Then Exit Function
    日時帯文字列を時刻範囲に = True
End Function

Private Function マスタブック_メイン設定シートを取得(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    Dim sh As Worksheet
    Dim best As Worksheet
    Dim bestLen As Long
    Dim L As Long
    If wb Is Nothing Then Exit Function
    On Error Resume Next
    Set ws = wb.Worksheets("メイン")
    If ws Is Nothing Then Set ws = wb.Worksheets("メイン_")
    If ws Is Nothing Then Set ws = wb.Worksheets("Main")
    On Error GoTo 0
    If Not ws Is Nothing Then
        Set マスタブック_メイン設定シートを取得 = ws
        Exit Function
    End If
    Set best = Nothing
    bestLen = 10000
    For Each sh In wb.Worksheets
        If InStr(sh.Name, "メイン") > 0 Then
            If InStr(sh.Name, "カレンダー") > 0 Then GoTo NextMastMainPick
            L = Len(sh.Name)
            If L < bestLen Then
                bestLen = L
                Set best = sh
            End If
        End If
NextMastMainPick:
    Next sh
    Set マスタブック_メイン設定シートを取得 = best
End Function

Private Function マスタメイン_結合左上の値(ByVal ws As Worksheet, ByVal cellAddr As String) As Variant
    Dim rng As Range
    On Error GoTo FailMMTL
    Set rng = ws.Range(cellAddr)
    マスタメイン_結合左上の値 = rng.MergeArea.Cells(1, 1).Value
    Exit Function
FailMMTL:
    マスタメイン_結合左上の値 = Empty
End Function

Private Sub 結果_設備毎の時間割_マスタ時刻反映( _
    ByVal ws As Worksheet, _
    ByVal regOk As Boolean, ByVal regS As Date, ByVal regE As Date, _
    ByVal facOk As Boolean, ByVal facS As Date, ByVal facE As Date)
    
    Dim colTB As Long
    Dim lastR As Long
    Dim r As Long
    Dim t0 As Date
    Dim t1 As Date
    Dim b0 As Long
    Dim b1 As Long
    Dim r0 As Long
    Dim r1 As Long
    Dim f0 As Long
    Dim f1 As Long
    
    On Error GoTo CleanExit
    If ws Is Nothing Then Exit Sub
    colTB = FindColHeader(ws, "日時帯")
    If colTB = 0 Then Exit Sub
    
    If regOk Then
        r0 = 時刻を分に(regS)
        r1 = 時刻を分に(regE)
    End If
    If facOk Then
        f0 = 時刻を分に(facS)
        f1 = 時刻を分に(facE)
    End If
    
    lastR = ws.Cells(ws.Rows.Count, colTB).End(xlUp).Row
    If lastR < 2 Then GoTo CleanExit
    
    For r = 2 To lastR
        If 日時帯文字列を時刻範囲に(ws.Cells(r, colTB).Value, t0, t1) Then
            b0 = 時刻を分に(t0)
            b1 = 時刻を分に(t1)
            With ws.Cells(r, colTB)
                If facOk And Not 半開区間が重なる分(b0, b1, f0, f1) Then
                    .Interior.Pattern = xlSolid
                    .Interior.Color = RGB(221, 235, 247)
                ElseIf regOk And Not 半開区間が重なる分(b0, b1, r0, r1) Then
                    .Interior.Pattern = xlSolid
                    .Interior.Color = RGB(252, 228, 214)
                Else
                    .Interior.Pattern = xlNone
                End If
            End With
        End If
    Next r
CleanExit:
    On Error GoTo 0
End Sub

Private Sub 結果_機械名毎時間割_依頼NOセルを薄緑(ByVal ws As Worksheet)
    Dim colTB As Long
    Dim lastR As Long
    Dim lastC As Long
    Dim r As Long
    Dim c As Long
    Dim v As Variant
    Dim s As String
    
    On Error GoTo CleanExit2
    If ws Is Nothing Then Exit Sub
    colTB = FindColHeader(ws, "日時帯")
    If colTB = 0 Then Exit Sub
    
    lastR = ws.Cells(ws.Rows.Count, colTB).End(xlUp).Row
    If lastR < 2 Then GoTo CleanExit2
    lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastC <= colTB Then GoTo CleanExit2
    
    For r = 2 To lastR
        For c = colTB + 1 To lastC
            v = ws.Cells(r, c).Value
            If IsEmpty(v) Or VarType(v) = vbError Then GoTo NextC2
            s = Trim$(Replace(Replace(CStr(v), vbCr, ""), vbLf, ""))
            If Len(s) = 0 Then GoTo NextC2
            If StrComp(s, "（休憩）", vbBinaryCompare) = 0 Then GoTo NextC2
            With ws.Cells(r, c)
                .Interior.Pattern = xlSolid
                .Interior.Color = RGB(198, 239, 206)
            End With
NextC2:
        Next c
    Next r
CleanExit2:
    On Error GoTo 0
End Sub

Private Sub 結果_設備時間割_準備後始末セルを薄緑(ByVal ws As Worksheet)
    Dim colTB As Long
    Dim lastR As Long
    Dim lastC As Long
    Dim r As Long
    Dim c As Long
    Dim v As Variant
    Dim s As String
    Dim hdr As String
    
    On Error GoTo CleanExit3
    If ws Is Nothing Then Exit Sub
    colTB = FindColHeader(ws, "日時帯")
    If colTB = 0 Then Exit Sub
    
    lastR = ws.Cells(ws.Rows.Count, colTB).End(xlUp).Row
    If lastR < 2 Then GoTo CleanExit3
    lastC = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastC <= colTB Then GoTo CleanExit3
    
    For r = 2 To lastR
        For c = colTB + 1 To lastC
            hdr = Trim$(CStr(ws.Cells(1, c).Value))
            If Len(hdr) >= 2 Then
                If Right$(hdr, 2) = "進度" Then GoTo NextC3
            End If
            v = ws.Cells(r, c).Value
            If IsEmpty(v) Or VarType(v) = vbError Then GoTo NextC3
            s = Trim$(Replace(Replace(CStr(v), vbCr, ""), vbLf, ""))
            If Len(s) = 0 Then GoTo NextC3
            If InStr(1, s, "(日次始業準備)", vbTextCompare) > 0 _
                Or InStr(1, s, "(加工前準備)", vbTextCompare) > 0 _
                Or InStr(1, s, "(依頼切替後始末)", vbTextCompare) > 0 Then
                With ws.Cells(r, c)
                    .Interior.Pattern = xlSolid
                    .Interior.Color = RGB(198, 239, 206)
                End With
            End If
NextC3:
        Next c
    Next r
CleanExit3:
    On Error GoTo 0
End Sub

Private Sub 結果_設備ガント_マスタ時刻反映( _
    ByVal ws As Worksheet, _
    ByVal regOk As Boolean, ByVal regS As Date, ByVal regE As Date, _
    ByVal facOk As Boolean, ByVal facS As Date, ByVal facE As Date)
    
    Dim lastR As Long
    Dim r As Long
    Dim c As Long
    Dim lastC As Long
    Dim slotStart As Date
    Dim slotEnd As Date
    Dim s0 As Long
    Dim s1 As Long
    Dim r0 As Long
    Dim r1 As Long
    Dim f0 As Long
    Dim f1 As Long
    Dim v As Variant
    Dim ur As Range
    
    On Error GoTo CleanExit
    If ws Is Nothing Then Exit Sub
    
    If regOk Then
        r0 = 時刻を分に(regS)
        r1 = 時刻を分に(regE)
    End If
    If facOk Then
        f0 = 時刻を分に(facS)
        f1 = 時刻を分に(facE)
    End If
    
    Set ur = ws.UsedRange
    If ur Is Nothing Then GoTo CleanExit
    lastR = ur.Row + ur.Rows.Count - 1
    
    For r = 1 To lastR
        If Trim$(CStr(ws.Cells(r, 2).Value)) = "機械名" _
            And Trim$(CStr(ws.Cells(r, 3).Value)) = "工程名" _
            And Trim$(CStr(ws.Cells(r, 4).Value)) = "担当者" _
            And Trim$(CStr(ws.Cells(r, 5).Value)) = "タスク概要" Then
            
            lastC = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
            For c = 6 To lastC
                v = ws.Cells(r, c).Value
                If Not IsEmpty(v) And VarType(v) <> vbError Then
                    If マスタメイン_セルを時刻Dateへ(v, slotStart) Then
                        slotEnd = slotStart + TimeSerial(0, 15, 0)
                        s0 = 時刻を分に(slotStart)
                        s1 = 時刻を分に(slotEnd)
                        With ws.Cells(r, c)
                            If facOk And Not 半開区間が重なる分(s0, s1, f0, f1) Then
                                .Interior.Pattern = xlSolid
                                .Interior.Color = RGB(221, 235, 247)
                            ElseIf regOk And Not 半開区間が重なる分(s0, s1, r0, r1) Then
                                .Interior.Pattern = xlSolid
                                .Interior.Color = RGB(252, 228, 214)
                            Else
                                .Interior.Pattern = xlSolid
                                .Interior.Color = RGB(217, 217, 217)
                            End If
                        End With
                    End If
                End If
            Next c
        End If
    Next r
CleanExit:
    On Error GoTo 0
End Sub

Private Sub 取込後_結果シートへマスタ時刻を反映(ByVal wb As Workbook)
    Dim facOk As Boolean
    Dim regOk As Boolean
    Dim facS As Date
    Dim facE As Date
    Dim regS As Date
    Dim regE As Date
    Dim ws As Worksheet
    
    If wb Is Nothing Then Exit Sub
    マスタメイン_工場稼働と定常を取得 facOk, facS, facE, regOk, regS, regE
    
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_RESULT_EQUIP_SCHEDULE)
    If Not ws Is Nothing Then
        結果_設備毎の時間割_マスタ時刻反映 ws, regOk, regS, regE, facOk, facS, facE
        結果_設備時間割_準備後始末セルを薄緑 ws
    End If
    Set ws = Nothing
    Set ws = wb.Worksheets("TEMP_設備毎の時間割")
    If Not ws Is Nothing Then
        結果_設備時間割_準備後始末セルを薄緑 ws
    End If
    Set ws = Nothing
    Set ws = wb.Worksheets(SHEET_RESULT_EQUIP_BY_MACHINE)
    If Not ws Is Nothing Then
        結果_設備毎の時間割_マスタ時刻反映 ws, regOk, regS, regE, facOk, facS, facE
        結果_機械名毎時間割_依頼NOセルを薄緑 ws
    End If
    Set ws = Nothing
    Set ws = wb.Worksheets("結果_設備ガント")
    If Not ws Is Nothing Then
        結果_設備ガント_マスタ時刻反映 ws, regOk, regS, regE, facOk, facS, facE
    End If
    On Error GoTo 0
End Sub

Public Sub 結果シート_マスタ工場稼働と定常を再適用()
    取込後_結果シートへマスタ時刻を反映 ThisWorkbook
End Sub

Private Function メインシート_勤怠表示先頭ゼロ無し(ByVal labeled As String) As String
    Dim parts() As String
    Dim a As String, b As String
    parts = Split(labeled, " / ")
    If UBound(parts) <> 1 Then
        メインシート_勤怠表示先頭ゼロ無し = labeled
        Exit Function
    End If
    a = Trim$(parts(0))
    b = Trim$(parts(1))
    If Len(a) >= 4 And Left$(a, 1) = "0" Then a = Mid$(a, 2)
    If Len(b) >= 4 And Left$(b, 1) = "0" Then b = Mid$(b, 2)
    メインシート_勤怠表示先頭ゼロ無し = a & " / " & b
End Function

Private Function メインシート_勤怠表示が通常勤務か(ByVal txt As String, Optional ByVal stdDispCached As String) As Boolean
    Dim t As String
    Dim exp As String
    t = Trim$(Replace(Replace(txt, vbCr, ""), vbLf, ""))
    t = Replace(t, "：", ":")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    If Len(stdDispCached) > 0 Then
        exp = stdDispCached
    Else
        exp = マスタメイン_工場標準勤怠表示文字列()
    End If
    If StrComp(t, exp, vbTextCompare) = 0 Then
        メインシート_勤怠表示が通常勤務か = True
    ElseIf StrComp(t, メインシート_勤怠表示先頭ゼロ無し(exp), vbTextCompare) = 0 Then
        メインシート_勤怠表示が通常勤務か = True
    Else
        メインシート_勤怠表示が通常勤務か = False
    End If
End Function

Private Sub メインシート_勤怠セルに背景色を設定(ByVal c As Range, ByVal displayVal As String, ByVal stdDispCached As String)
    Dim s As String
    s = Trim$(CStr(displayVal))
    On Error Resume Next
    If s = "" Or s = "-" Then
        c.Interior.Pattern = xlSolid
        c.Interior.Color = RGB(242, 242, 242)
    ElseIf メインシート_勤怠表示が通常勤務か(s, stdDispCached) Then
        c.Interior.Pattern = xlNone
    Else
        c.Interior.Pattern = xlSolid
        c.Interior.Color = RGB(255, 242, 204)
    End If
    On Error GoTo 0
End Sub

Private Sub メインシート_メンバー勤怠ブロックに罫線を設定(ByVal wsMain As Worksheet, ByVal lastMemberRow As Long)
    Dim rng As Range
    Const lastCol As Long = 14   ' N列（B=メンバー、C～N=12日分）
    If wsMain Is Nothing Then Exit Sub
    If lastMemberRow < 7 Then Exit Sub
    On Error Resume Next
    Set rng = wsMain.Range(wsMain.Cells(7, 2), wsMain.Cells(lastMemberRow, lastCol))
    With rng
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone
        With .Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .Weight = xlThin
        End With
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .Weight = xlThin
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .Weight = xlThin
        End With
        With .Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .Weight = xlThin
        End With
        With .Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .Weight = xlThin
        End With
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .ColorIndex = xlAutomatic
            .Weight = xlThin
        End With
    End With
    On Error GoTo 0
End Sub

Public Sub メインシート_メンバー一覧と出勤表示(Optional ByVal Silent As Boolean = False)
    Dim wb As Workbook
    Dim wsMain As Worksheet
    Dim wsCal As Worksheet
    Dim ws As Worksheet
    Dim dict As Object
    Dim members As Object
    Dim keys As Variant
    Dim keysArr() As String
    Dim i As Long, j As Long, r As Long, col As Long
    Dim lastR As Long
    Dim mn As String
    Dim sheetName As String
    Dim d As Date
    Dim k As String
    Dim colDate As Long, colMem As Long, colIn As Long, colOut As Long
    Dim wkStr As String
    Dim temp As String
    Dim cnt As Long
    Dim srcHdr As Range, srcMem As Range
    Dim bHdrFn As String, bHdrFs As Double, bHdrFc As Variant
    Dim bHdrBold As Boolean, bHdrIt As Boolean, bHdrUl As Long
    Dim bMemFn As String, bMemFs As Double, bMemFc As Variant
    Dim bMemBold As Boolean, bMemIt As Boolean, bMemUl As Long
    Dim lastMemberRow As Long
    Dim stdDispCached As String
    
    lastMemberRow = 0
    On Error GoTo EH
    
    Set wb = ThisWorkbook
    Set wsMain = GetMainWorksheet()
    If wsMain Is Nothing Then
        If Not Silent Then MsgBox "「メイン」「Main」、または名前に「メイン」を含むシートが見つかりません。", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' クリア前に B 列・見出しの見本フォントを記憶（無ければ日付列 C から）
    Set srcHdr = wsMain.Cells(7, 2)
    If Len(Trim$(CStr(srcHdr.Value))) = 0 Then Set srcHdr = wsMain.Cells(7, 3)
    メインシート_フォント属性を取得 srcHdr, bHdrFn, bHdrFs, bHdrFc, bHdrBold, bHdrIt, bHdrUl
    Set srcMem = wsMain.Cells(8, 2)
    If Len(Trim$(CStr(srcMem.Value))) = 0 Then Set srcMem = wsMain.Cells(8, 3)
    If Len(Trim$(CStr(srcMem.Value))) = 0 Then Set srcMem = wsMain.Cells(7, 3)
    メインシート_フォント属性を取得 srcMem, bMemFn, bMemFs, bMemFc, bMemBold, bMemIt, bMemUl
    
    ' Clear だとフォント等の書式まで消える → ClearContents のみ。B列の個人リンクは削除してから再付与
    メインシート_指定範囲のハイパーリンクを削除 wsMain, wsMain.Range("B7:B500")
    wsMain.Range("B7:N500").ClearContents
    On Error Resume Next
    wsMain.Range("C7:N500").Interior.Pattern = xlNone
    On Error GoTo EH
    
    ' 見出し行（前日から12日間）
    wsMain.Cells(7, 2).Value = "メンバー"
    メインシート_フォント属性を適用 wsMain.Cells(7, 2), bHdrFn, bHdrFs, bHdrFc, bHdrBold, bHdrIt, bHdrUl
    For i = 0 To 11
        d = DateAdd("d", i - 1, Date)
        wkStr = Split("月,火,水,木,金,土,日", ",")(Weekday(d, vbMonday) - 1)
        wsMain.Cells(7, 3 + i).Value = Format$(d, "m/d") & "(" & wkStr & ")"
        wsMain.Cells(7, 3 + i).HorizontalAlignment = xlCenter
    Next i
    
    ' 個人_* シートからメンバー名一覧（重複なし）
    Set members = CreateObject("Scripting.Dictionary")
    For Each ws In wb.Worksheets
        If Left$(ws.Name, 3) = "個人_" Then
            mn = Mid$(ws.Name, 4)
            If Len(mn) > 0 And Not members.Exists(mn) Then members.Add mn, mn
        End If
    Next ws
    
    ' 結果_カレンダー(出勤簿) から (メンバー,日付) → 出勤/退勤
    Set dict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set wsCal = wb.Worksheets("結果_カレンダー(出勤簿)")
    On Error GoTo EH
    
    If Not wsCal Is Nothing Then
        colDate = FindColHeader(wsCal, "日付")
        colMem = FindColHeader(wsCal, "メンバー")
        colIn = FindColHeader(wsCal, "出勤")
        colOut = FindColHeader(wsCal, "退勤")
        If colDate > 0 And colMem > 0 And colIn > 0 And colOut > 0 Then
            lastR = wsCal.Cells(wsCal.Rows.Count, colDate).End(xlUp).Row
            For r = 2 To lastR
                If IsDate(wsCal.Cells(r, colDate).Value) Then
                    d = CDate(wsCal.Cells(r, colDate).Value)
                    mn = Trim$(CStr(wsCal.Cells(r, colMem).Value))
                    If Len(mn) > 0 Then
                        k = mn & "|" & Format$(d, "yyyy-mm-dd")
                        dict(k) = Trim$(CStr(wsCal.Cells(r, colIn).Value)) & " / " & Trim$(CStr(wsCal.Cells(r, colOut).Value))
                    End If
                End If
            Next r
        End If
    End If
    
    cnt = members.Count
    If cnt = 0 Then
        wsMain.Cells(8, 2).Value = "（個人_* のシートがありません）"
        メインシート_フォント属性を適用 wsMain.Cells(8, 2), bMemFn, bMemFs, bMemFc, bMemBold, bMemIt, bMemUl
        lastMemberRow = 8
        GoTo CleanExit
    End If
    
    keys = members.keys
    ReDim keysArr(0 To UBound(keys))
    For i = 0 To UBound(keys)
        keysArr(i) = CStr(keys(i))
    Next i
    ' 単純ソート（表示順）
    For i = 0 To UBound(keysArr) - 1
        For j = i + 1 To UBound(keysArr)
            If keysArr(i) > keysArr(j) Then
                temp = keysArr(i): keysArr(i) = keysArr(j): keysArr(j) = temp
            End If
        Next j
    Next i
    
    ' master.xlsm の定常表示はセルごとに読むと都度 Open/Close になり得るため、ここで1回だけ取得して勤怠セル着色に渡す
    stdDispCached = マスタメイン_工場標準勤怠表示文字列()
    
    r = 8
    For i = 0 To UBound(keysArr)
        mn = keysArr(i)
        sheetName = SafePersonalSheetName(mn)
        On Error Resume Next
        wsMain.Hyperlinks.Add anchor:=wsMain.Cells(r, 2), Address:="", SubAddress:="'" & Replace(sheetName, "'", "''") & "'!A1", TextToDisplay:=mn
        On Error GoTo EH
        メインシート_フォント属性を適用 wsMain.Cells(r, 2), bMemFn, bMemFs, bMemFc, bMemBold, bMemIt, bMemUl
        
        For col = 0 To 11
            d = DateAdd("d", col - 1, Date)
            k = mn & "|" & Format$(d, "yyyy-mm-dd")
            If dict.Exists(k) Then
                wsMain.Cells(r, 3 + col).Value = dict(k)
                メインシート_勤怠セルに背景色を設定 wsMain.Cells(r, 3 + col), CStr(dict(k)), stdDispCached
            Else
                wsMain.Cells(r, 3 + col).Value = "-"
                メインシート_勤怠セルに背景色を設定 wsMain.Cells(r, 3 + col), "-", stdDispCached
            End If
            wsMain.Cells(r, 3 + col).HorizontalAlignment = xlCenter
        Next col
        r = r + 1
    Next i
    lastMemberRow = r - 1

CleanExit:
    On Error Resume Next
    メインシート_メンバー勤怠ブロックに罫線を設定 wsMain, lastMemberRow
    メインシート_結果シートリンクを更新 wsMain
    メインシート_AからK列_AutoFitOnSheet wsMain
    Application.ScreenUpdating = True
    Exit Sub
EH:
    If Not Silent Then MsgBox "メインシート更新エラー: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Private Sub メインシート_AからK列_AutoFitOnSheet(ByVal wsMain As Worksheet)
    On Error Resume Next
    If wsMain Is Nothing Then Exit Sub
    wsMain.Columns("A:N").AutoFit
    On Error GoTo 0
End Sub

Public Sub メインシート_AからK列_AutoFit()
    Dim ws As Worksheet
    Dim su As Boolean
    On Error Resume Next
    Set ws = GetMainWorksheet()
    If ws Is Nothing Then Exit Sub
    su = Application.ScreenUpdating
    Application.ScreenUpdating = True
    メインシート_AからK列_AutoFitOnSheet ws
    Application.ScreenUpdating = su
    On Error GoTo 0
End Sub

Private Function GetMainWorksheet() As Worksheet
    ' 配台ブックのメイン UI はシート名「メイン_」固定（旧「メイン」「Main」や部分一致は使わない）
    On Error Resume Next
    Set GetMainWorksheet = ThisWorkbook.Worksheets("メイン_")
    On Error GoTo 0
End Function

Private Sub メインシートA1を選択()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetMainWorksheet()
    If ws Is Nothing Then Exit Sub
    ws.Activate
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

Private Function FindColHeader(ws As Worksheet, ByVal headerText As String) As Long
    Dim c As Long
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If Trim$(CStr(ws.Cells(1, c).Value)) = headerText Then
            FindColHeader = c
            Exit Function
        End If
    Next c
    FindColHeader = 0
End Function

Private Sub メインシート_指定範囲のハイパーリンクを削除(ByVal wsMain As Worksheet, ByVal Target As Range)
    Dim c As Range
    If wsMain Is Nothing Or Target Is Nothing Then Exit Sub
    On Error Resume Next
    For Each c In Target.Cells
        If c.Hyperlinks.Count > 0 Then c.Hyperlinks.Delete
    Next c
    On Error GoTo 0
End Sub

Private Sub メインシート_結果シートリンクを更新(ByVal wsMain As Worksheet)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim coll As Collection
    Dim arr() As String
    Dim i As Long, j As Long, r As Long
    Dim n As Long
    Dim temp As String
    Dim sn As String
    Dim afn As String, afs As Double, afc As Variant
    Dim aBold As Boolean, aIt As Boolean, aUl As Long
    Dim srcA As Range
    
    ' クリア前に A 列リンクの見本フォントを記憶（無ければ日付見出し C7）
    Set srcA = wsMain.Cells(2, 1)
    If Len(Trim$(CStr(srcA.Value))) = 0 Then Set srcA = wsMain.Cells(1, 1)
    If Len(Trim$(CStr(srcA.Value))) = 0 Then Set srcA = wsMain.Cells(7, 3)
    メインシート_フォント属性を取得 srcA, afn, afs, afc, aBold, aIt, aUl
    
    Set wb = wsMain.Parent
    Set coll = New Collection
    For Each ws In wb.Worksheets
        If Left$(ws.Name, 3) = "結果_" Then coll.Add ws.Name
    Next ws
    
    ' Clear だと A 列のフォントが既定に戻る → リンク削除＋内容のみクリア
    ' 結果_* が増えても取りこぼさないよう A 列の確保行を広げる（「計画結果」見出し＋リンク列）
    メインシート_指定範囲のハイパーリンクを削除 wsMain, wsMain.Range("A1:A120")
    wsMain.Range("A1:A120").ClearContents
    
    If coll.Count = 0 Then
        wsMain.Range("A1").Value = "（結果_* シートなし）"
        メインシート_フォント属性を適用 wsMain.Cells(1, 1), afn, afs, afc, aBold, aIt, aUl
        Exit Sub
    End If
    
    n = coll.Count
    ReDim arr(1 To n)
    For i = 1 To n
        arr(i) = coll(i)
    Next i
    For i = 1 To n - 1
        For j = i + 1 To n
            If arr(i) > arr(j) Then
                temp = arr(i): arr(i) = arr(j): arr(j) = temp
            End If
        Next j
    Next i
    
    wsMain.Cells(1, 1).Value = "計画結果（シートへ）"
    メインシート_フォント属性を適用 wsMain.Cells(1, 1), afn, afs, afc, True, aIt, aUl
    r = 2
    For i = 1 To n
        sn = arr(i)
        wsMain.Hyperlinks.Add anchor:=wsMain.Cells(r, 1), Address:="", SubAddress:="'" & Replace(sn, "'", "''") & "'!A1", TextToDisplay:=sn
        メインシート_フォント属性を適用 wsMain.Cells(r, 1), afn, afs, afc, False, aIt, aUl
        r = r + 1
    Next i
End Sub

Private Sub 結果シート_列幅_AutoFit安定(ByVal targetWs As Worksheet)
    Dim su As Boolean
    If StrComp(targetWs.Name, SCRATCH_SHEET_FONT, vbBinaryCompare) = 0 Then Exit Sub
    ' 結果_設備ガントは専用列幅（時刻グリッド）のため絶対に EntireColumn.AutoFit しない
    If StrComp(Trim$(targetWs.Name), "結果_設備ガント", vbBinaryCompare) = 0 Then Exit Sub
    su = Application.ScreenUpdating
    On Error Resume Next
    Application.ScreenUpdating = True
    targetWs.Activate
    DoEvents
    targetWs.Cells.Select
    DoEvents
    targetWs.Cells.EntireColumn.AutoFit
    targetWs.Range("A1").Select
    Application.ScreenUpdating = su
    On Error GoTo 0
End Sub

Private Sub 結果シート_列幅_AutoFit非表示を維持(ByVal targetWs As Worksheet)
    Dim su As Boolean
    Dim lastCol As Long
    Dim c As Long
    
    If StrComp(targetWs.Name, SCRATCH_SHEET_FONT, vbBinaryCompare) = 0 Then Exit Sub
    
    On Error Resume Next
    lastCol = targetWs.UsedRange.Column + targetWs.UsedRange.Columns.Count - 1
    If lastCol < 1 Then lastCol = targetWs.Cells(1, targetWs.Columns.Count).End(xlToLeft).Column
    On Error GoTo 0
    If lastCol < 1 Then Exit Sub
    
    su = Application.ScreenUpdating
    Application.ScreenUpdating = True
    On Error Resume Next
    targetWs.Activate
    DoEvents
    
    For c = 1 To lastCol
        If Not targetWs.Columns(c).Hidden Then
            targetWs.Columns(c).AutoFit
        End If
    Next c
    
    targetWs.Range("A1").Select
    Application.ScreenUpdating = su
    On Error GoTo 0
End Sub

Private Sub 結果_タスク一覧_配完回答指定16時_いいえを強調(ByVal ws As Worksheet)
    Dim c As Long
    Dim lastRow As Long
    Dim r As Long
    Dim v As Variant
    
    If ws Is Nothing Then Exit Sub
    If StrComp(ws.Name, SHEET_RESULT_TASK_LIST, vbBinaryCompare) <> 0 Then Exit Sub
    
    c = FindColHeader(ws, "配完_回答指定16時まで")
    If c <= 0 Then c = FindColHeader(ws, "配完_基準16時まで")
    If c <= 0 Then Exit Sub
    
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If lastRow < 2 Then Exit Sub
    
    On Error Resume Next
    For r = 2 To lastRow
        v = ws.Cells(r, c).Value
        If IsError(v) Then
            ' skip
        ElseIf Trim$(CStr(v)) = "いいえ" Then
            With ws.Cells(r, c)
                .Interior.Color = RGB(255, 0, 0)
                .Font.Color = RGB(255, 255, 255)
                .Font.Bold = True
            End With
        End If
    Next r
    On Error GoTo 0
End Sub

Private Sub 列設定結果タスク一覧_番号付き重複シートを削除(ByVal wb As Workbook)
    Dim i As Long
    Dim ws As Worksheet
    Dim pfx As String
    Dim prevDA As Boolean
    
    pfx = SHEET_COL_CONFIG_RESULT_TASK & " ("
    prevDA = Application.DisplayAlerts
    Application.DisplayAlerts = False
    For i = wb.Worksheets.Count To 1 Step -1
        Set ws = wb.Worksheets(i)
        If InStr(1, ws.Name, pfx, vbBinaryCompare) = 1 Then
            ws.Delete
        End If
    Next i
    Application.DisplayAlerts = prevDA
End Sub

Private Sub 列設定_結果_タスク一覧を設定の直前へ移動(ByVal wb As Workbook)
    Dim wsCfg As Worksheet
    Dim wsSet As Worksheet
    
    On Error Resume Next
    Set wsCfg = wb.Worksheets(SHEET_COL_CONFIG_RESULT_TASK)
    Set wsSet = wb.Worksheets(SHEET_SETTINGS)
    On Error GoTo 0
    If wsCfg Is Nothing Or wsSet Is Nothing Then Exit Sub
    
    On Error Resume Next
    If wsCfg.Index <> wsSet.Index - 1 Then
        wsCfg.Move Before:=wsSet
    End If
    On Error GoTo 0
End Sub

Private Sub 結果シート_メインへ戻るリンクを付与(ByVal ws As Worksheet)
    Dim wsMain As Worksheet
    Dim mainName As String
    Dim lastCol As Long
    Dim anchor As Range
    
    Set wsMain = GetMainWorksheet()
    If wsMain Is Nothing Then Exit Sub
    If StrComp(ws.Name, wsMain.Name, vbBinaryCompare) = 0 Then Exit Sub
    
    mainName = wsMain.Name
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then lastCol = 1
    Set anchor = ws.Cells(1, lastCol + 2)
    
    On Error Resume Next
    anchor.Hyperlinks.Delete
    On Error GoTo 0
    ws.Hyperlinks.Add anchor:=anchor, Address:="", SubAddress:="'" & Replace(mainName, "'", "''") & "'!A1", TextToDisplay:="≪ メインへ"
    With anchor
        .Font.Bold = False
        .Interior.Pattern = xlNone
        .HorizontalAlignment = xlRight
    End With
End Sub

Private Function 結果_設備毎の時間割_B2選択して窓枠固定() As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_RESULT_EQUIP_SCHEDULE)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("B2").Select
    ActiveWindow.FreezePanes = True
    結果_設備毎の時間割_B2選択して窓枠固定 = True
End Function

Private Function 結果_タスク一覧_F2選択して窓枠固定() As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_RESULT_TASK_LIST)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("F2").Select
    ActiveWindow.FreezePanes = True
    結果_タスク一覧_F2選択して窓枠固定 = True
End Function

Private Function 結果_カレンダー出勤簿_A2選択して窓枠固定() As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(SHEET_RESULT_CALENDAR_ATTEND)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
    結果_カレンダー出勤簿_A2選択して窓枠固定 = True
End Function

Private Sub 結果プレフィックスシートの表示倍率を設定(ByVal wb As Workbook, ByVal zoomPercent As Long)
    Dim ws As Worksheet
    Dim prevScr As Boolean
    prevScr = Application.ScreenUpdating
    On Error Resume Next
    Application.ScreenUpdating = False
    For Each ws In wb.Worksheets
        If Left$(ws.Name, 3) = "結果_" Then
            ws.Activate
            ActiveWindow.Zoom = zoomPercent
        End If
    Next ws
    On Error GoTo 0
    Application.ScreenUpdating = prevScr
End Sub

Private Sub 結果_設備ガント_表示倍率を設定(ByVal wb As Workbook, ByVal zoomPercent As Long)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(SHEET_RESULT_EQUIP_GANTT)
    If ws Is Nothing Then GoTo CleanZoom
    ws.Activate
    ActiveWindow.Zoom = zoomPercent
    ws.Range("A1").Activate
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
CleanZoom:
    On Error GoTo 0
End Sub

Private Sub 結果_設備ガント_列幅を設定(ByVal ws As Worksheet)
    Dim lastCol As Long
    Dim c As Long
    Dim wE As Double
    
    On Error Resume Next
    ws.Columns("A").ColumnWidth = 12   ' 日付（縦結合）
    ws.Columns("B").ColumnWidth = 16   ' 機械名（Python 側フォント拡大に合わせる）
    ws.Columns("C").ColumnWidth = 16   ' 工程名
    ws.Columns("D").ColumnWidth = 26   ' 担当者（主担当＋サブ列挙）
    ' E: タスク概要（依頼NO）… 列幅を約 38 ポイントにし折り返し
    ws.Columns("E").ColumnWidth = 8
    wE = ws.Columns("E").Width
    If wE > 0 Then
        ws.Columns("E").ColumnWidth = 38
    End If
    ws.Columns("E").WrapText = True
    ws.Columns("D").WrapText = True
    lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    On Error GoTo 0
    If lastCol < 6 Then Exit Sub
    For c = 6 To lastCol
        ws.Columns(c).ColumnWidth = 7.5   ' 時刻見出し 90° 回転・帯ラベル用に拡大
    Next c
    On Error Resume Next
    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("F4").Activate
    ActiveWindow.FreezePanes = True
    ws.Range("A1").Activate
    ActiveWindow.Zoom = 85
    ActiveWindow.ScrollColumn = 1
    ActiveWindow.ScrollRow = 1
    On Error GoTo 0
End Sub

Private Sub 結果_設備ガント_タイトルA1を左寄せに固定(ByVal ws As Worksheet)
    On Error Resume Next
    With ws.Range("A1")
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub 結果_設備ガント_行枠を通常に戻す(ByVal rng As Range)
    On Error Resume Next
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(102, 102, 102)
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(102, 102, 102)
    End With
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(102, 102, 102)
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(102, 102, 102)
    End With
    On Error GoTo 0
End Sub

Private Sub 結果_設備ガント_行枠を強調(ByVal rng As Range)
    ' xlThick 単線より視認性を上げるため二重線＋濃いオレンジ（Excel の Weight は xlThick が上限のため）
    Const hlR As Long = 204
    Const hlG As Long = 0
    Const hlB As Long = 0
    On Error Resume Next
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlDouble
        .Weight = xlThick
        .Color = RGB(hlR, hlG, hlB)
    End With
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlDouble
        .Weight = xlThick
        .Color = RGB(hlR, hlG, hlB)
    End With
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlDouble
        .Weight = xlThick
        .Color = RGB(hlR, hlG, hlB)
    End With
    With rng.Borders(xlEdgeRight)
        .LineStyle = xlDouble
        .Weight = xlThick
        .Color = RGB(hlR, hlG, hlB)
    End With
    On Error GoTo 0
End Sub

Private Function 結果_設備ガント_行は表頭行か(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim a As String
    Dim b As String
    On Error Resume Next
    a = Trim$(CStr(ws.Cells(r, 2).Value))
    b = Trim$(CStr(ws.Cells(r, 3).Value))
    On Error GoTo 0
    結果_設備ガント_行は表頭行か = (StrComp(a, "機械名", vbBinaryCompare) = 0 And StrComp(b, "工程名", vbBinaryCompare) = 0)
End Function

Private Function 結果_設備ガント_行は区切り行か(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim rh As Double
    On Error Resume Next
    rh = ws.Rows(r).RowHeight
    On Error GoTo 0
    If rh > 0# And rh <= 5.6 Then 結果_設備ガント_行は区切り行か = True
End Function

Private Function 結果_設備ガント_行はデータ行か(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    結果_設備ガント_行はデータ行か = False
    If r <= 2 Then Exit Function
    If 結果_設備ガント_行は表頭行か(ws, r) Then Exit Function
    If 結果_設備ガント_行は区切り行か(ws, r) Then Exit Function
    If r < 4 Then Exit Function
    結果_設備ガント_行はデータ行か = True
End Function

Private Sub 結果_設備ガント_行ハイライト_Clear(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim rng As Range
    
    If Len(mGanttHL_SheetName) = 0 Then Exit Sub
    If mGanttHL_Row < 1 Or mGanttHL_LastCol < 1 Then GoTo ResetState
    
    On Error Resume Next
    Set ws = wb.Worksheets(mGanttHL_SheetName)
    On Error GoTo 0
    If ws Is Nothing Then GoTo ResetState
    
    On Error Resume Next
    Set rng = ws.Range(ws.Cells(mGanttHL_Row, 1), ws.Cells(mGanttHL_Row, mGanttHL_LastCol))
    結果_設備ガント_行枠を通常に戻す rng
    On Error GoTo 0
    
ResetState:
    mGanttHL_SheetName = vbNullString
    mGanttHL_Row = 0
    mGanttHL_LastCol = 0
End Sub

Public Sub 結果_設備ガント_行ハイライト_OnSelection(ByVal sh As Object, ByVal Target As Range)
    Dim ws As Worksheet
    Dim wb As Workbook
    Dim r As Long
    Dim lastCol As Long
    Dim rng As Range
    
    On Error GoTo QuietExit
    If sh Is Nothing Then Exit Sub
    If Not TypeOf sh Is Worksheet Then Exit Sub
    Set ws = sh
    Set wb = ws.Parent
    
    結果_設備ガント_行ハイライト_Clear wb
    
    If StrComp(ws.Name, SHEET_RESULT_EQUIP_GANTT, vbBinaryCompare) <> 0 Then Exit Sub
    If Target Is Nothing Then Exit Sub
    
    r = Target.Cells(1, 1).Row
    If Not 結果_設備ガント_行はデータ行か(ws, r) Then Exit Sub
    
    On Error Resume Next
    lastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    On Error GoTo 0
    If lastCol < 4 Then lastCol = 4
    
    Set rng = ws.Range(ws.Cells(r, 1), ws.Cells(r, lastCol))
    結果_設備ガント_行枠を強調 rng
    
    mGanttHL_SheetName = ws.Name
    mGanttHL_Row = r
    mGanttHL_LastCol = lastCol
QuietExit:
End Sub

Public Sub 結果_設備ガント_保護を書式設定許可で更新()
    Dim ws As Worksheet
    On Error GoTo Quiet
    Set ws = ThisWorkbook.Worksheets(SHEET_RESULT_EQUIP_GANTT)
    If ws Is Nothing Then Exit Sub
    If ws.ProtectContents Then
        If Len(SHEET_FONT_UNPROTECT_PASSWORD) > 0 Then
            ws.Unprotect Password:=SHEET_FONT_UNPROTECT_PASSWORD
        End If
        If ws.ProtectContents Then ws.Unprotect
    End If
    If ws.ProtectContents Then Exit Sub
    If Len(SHEET_FONT_UNPROTECT_PASSWORD) > 0 Then
        ws.Protect Password:=SHEET_FONT_UNPROTECT_PASSWORD, DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFormattingCells:=True
    Else
        ws.Protect DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFormattingCells:=True
    End If
Quiet:
End Sub

Public Sub 結果_主要4結果シート_列オートフィット()
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ThisWorkbook
    On Error Resume Next
    
    ' 先に他シートの AutoFit を済ませ、最後に設備ガントの専用列幅を適用（他処理での幅変動を上書き）
    Err.Clear
    Set ws = wb.Worksheets("結果_カレンダー(出勤簿)")
    If Err.Number = 0 Then
        結果シート_列幅_AutoFit安定 ws
    Else
        Err.Clear
    End If
    
    Err.Clear
    Set ws = Nothing
    Set ws = wb.Worksheets("結果_メンバー別作業割合")
    If Err.Number = 0 Then
        結果シート_列幅_AutoFit安定 ws
    Else
        Err.Clear
    End If
    
    Err.Clear
    Set ws = Nothing
    Set ws = wb.Worksheets(SHEET_RESULT_TASK_LIST)
    If Err.Number = 0 Then
        結果シート_列幅_AutoFit非表示を維持 ws
    Else
        Err.Clear
    End If
    
    Err.Clear
    Set ws = Nothing
    Set ws = wb.Worksheets("結果_設備ガント")
    If Err.Number = 0 Then
        結果_設備ガント_列幅を設定 ws
        結果_設備ガント_タイトルA1を左寄せに固定 ws
    Else
        Err.Clear
    End If
    
    On Error GoTo 0
End Sub

Public Sub 個人シートを末尾へ並べ替え()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim arr() As String
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim temp As String
    
    Set wb = ThisWorkbook
    n = 0
    For Each ws In wb.Worksheets
        If Left$(ws.Name, 3) = "個人_" Then
            n = n + 1
            ReDim Preserve arr(1 To n)
            arr(n) = ws.Name
        End If
    Next ws
    
    If n > 0 Then
        For i = 1 To n - 1
            For j = i + 1 To n
                If arr(i) > arr(j) Then
                    temp = arr(i): arr(i) = arr(j): arr(j) = temp
                End If
            Next j
        Next i
    End If
    
    Application.ScreenUpdating = False
    
    ' 1) 個人_* を末尾へ（昇順）
    For i = 1 To n
        On Error Resume Next
        wb.Worksheets(arr(i)).Move After:=wb.Sheets(wb.Sheets.Count)
        On Error GoTo 0
    Next i
    
    ' 2) LOG を個人のさらに後ろ（ブック末尾）
    On Error Resume Next
    wb.Worksheets("LOG").Move After:=wb.Sheets(wb.Sheets.Count)
    On Error GoTo 0
    
    ' 3) 「設定」を最後尾（LOG のさらに後ろ）
    On Error Resume Next
    wb.Worksheets("設定").Move After:=wb.Sheets(wb.Sheets.Count)
    On Error GoTo 0
    
    Application.ScreenUpdating = True
End Sub

Private Sub 配台計画_タスク入力を前へ並べ替え()
    Const PLAN_SHEET As String = "配台計画_タスク入力"
    Dim wsPlan As Worksheet
    Dim wsMain As Worksheet
    Dim wsAfter As Worksheet
    
    MacroSplash_SetStep "段階1: 「配台計画_タスク入力」シートをメイン付近へ移動しています…"
    On Error Resume Next
    Set wsPlan = ThisWorkbook.Sheets(PLAN_SHEET)
    On Error GoTo 0
    If wsPlan Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set wsMain = GetMainWorksheet()
    On Error GoTo 0
    
    If wsMain Is Nothing Then
        Set wsAfter = ThisWorkbook.Sheets(1)
    Else
        Set wsAfter = wsMain
    End If
    
    If wsAfter Is Nothing Then Exit Sub
    If wsAfter.Name = wsPlan.Name Then Set wsAfter = ThisWorkbook.Sheets(1)
    
    If wsPlan.Index <> wsAfter.Index Then
        wsPlan.Move After:=wsAfter
    End If
End Sub

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

Private Function SettingsSheet_IsSplashExecutionLogWriteEnabled() As Boolean
    On Error GoTo DefaultTrue
    Dim ws As Worksheet
    Dim v As Variant
    Dim t As String
    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    v = ws.Range("D3").Value
    If IsError(v) Then GoTo DefaultTrue
    If VarType(v) = vbBoolean Then
        SettingsSheet_IsSplashExecutionLogWriteEnabled = CBool(v)
        Exit Function
    End If
    t = Trim$(CStr(v))
    If Len(t) = 0 Then GoTo DefaultTrue
    If StrComp(t, "false", vbTextCompare) = 0 Then
        SettingsSheet_IsSplashExecutionLogWriteEnabled = False
        Exit Function
    End If
    If StrComp(t, "true", vbTextCompare) = 0 Then
        SettingsSheet_IsSplashExecutionLogWriteEnabled = True
        Exit Function
    End If
DefaultTrue:
    SettingsSheet_IsSplashExecutionLogWriteEnabled = True
End Function

Private Function SettingsSheet_GetCompleteChimeTrack1to4() As Long
    On Error GoTo Def1
    Dim ws As Worksheet
    Dim v As Variant
    Dim n As Long
    Set ws = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    v = ws.Range("D4").Value
    If IsError(v) Then GoTo Def1
    If VarType(v) = vbString Then
        If Len(Trim$(CStr(v))) = 0 Then GoTo Def1
    End If
    If IsNumeric(v) Then
        n = CLng(CDbl(v))
    Else
        n = CLng(Val(CStr(v)))
    End If
    If n < 1 Or n > 4 Then GoTo Def1
    SettingsSheet_GetCompleteChimeTrack1to4 = n
    Exit Function
Def1:
    SettingsSheet_GetCompleteChimeTrack1to4 = 1
End Function

Private Function MacroCompleteChime_LocalWavPath() As String
    Dim folder As String
    folder = ThisWorkbook.path
    If Len(folder) = 0 Then Exit Function
    MacroCompleteChime_LocalWavPath = folder & "\" & MACRO_COMPLETE_CHIME_REL_DIR & "\" & MACRO_COMPLETE_CHIME_FILE_NAME
End Function

Private Function MacroCompleteChime_LocalMp3Path(ByVal track1to4 As Long) As String
    Dim folder As String
    Dim fn As String
    folder = ThisWorkbook.path
    If Len(folder) = 0 Then Exit Function
    Select Case track1to4
        Case 1: fn = MACRO_COMPLETE_MP3_1
        Case 2: fn = MACRO_COMPLETE_MP3_2
        Case 3: fn = MACRO_COMPLETE_MP3_3
        Case 4: fn = MACRO_COMPLETE_MP3_4
        Case Else: Exit Function
    End Select
    If Len(fn) = 0 Then Exit Function
    MacroCompleteChime_LocalMp3Path = folder & "\" & MACRO_COMPLETE_CHIME_REL_DIR & "\" & fn
End Function

Private Function MacroCompleteChime_MciPlayMp3(ByVal fullPath As String) As Boolean
    Dim a As String
    Dim cmdOpen As String
    Dim r As Long
    MacroCompleteChime_MciPlayMp3 = False
    a = ""
    On Error GoTo Fail
    Randomize
    ' Timer*1e6 は Long 上限を超えうるため、Rnd のみで 0?2147483646（CLng 安全域）
    a = "pm_ai_" & CStr(CLng(2147483646# * Rnd))
    r = mciSendStringW(StrPtr("close " & a), 0&, 0, 0&)
    Err.Clear
    cmdOpen = "open " & Chr$(34) & fullPath & Chr$(34) & " type mpegvideo alias " & a
    r = mciSendStringW(StrPtr(cmdOpen), 0&, 0, 0&)
    If r <> 0 Then GoTo Fail
    r = mciSendStringW(StrPtr("play " & a), 0&, 0, 0&)
    If r <> 0 Then GoTo Fail
    MacroCompleteChime_MciPlayMp3 = True
    Exit Function
Fail:
    On Error Resume Next
    If Len(a) > 0 Then r = mciSendStringW(StrPtr("close " & a), 0&, 0, 0&)
End Function

Private Function MacroCompleteChime_HttpDownloadBinary(ByVal url As String, ByVal destPath As String) As Boolean
    Dim xhr As Object
    Dim stm As Object
    On Error GoTo Fail
    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
    xhr.Open "GET", url, False
    xhr.setRequestHeader "User-Agent", "Excel-VBA-MacroCompleteChime/1"
    xhr.Send
    If xhr.Status < 200 Or xhr.Status >= 300 Then GoTo Fail
    If LenB(xhr.responseBody) = 0 Then GoTo Fail
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.Write xhr.responseBody
    stm.SaveToFile destPath, 2
    stm.Close
    MacroCompleteChime_HttpDownloadBinary = True
    Exit Function
Fail:
    On Error Resume Next
    If Not stm Is Nothing Then stm.Close
    MacroCompleteChime_HttpDownloadBinary = False
End Function

Public Sub SplashLog_AppendChunk(ByVal chunk As String)
    On Error Resume Next
    If Len(chunk) = 0 Then Exit Sub
    If Not SettingsSheet_IsSplashExecutionLogWriteEnabled() Then Exit Sub
    If Not m_macroSplashShown Then Exit Sub
    Dim tb As Object
    Set tb = frmMacroSplash.Controls("txtExecutionLog")
    If tb Is Nothing Then Exit Sub
    Dim n As Long
    n = Len(tb.text) + Len(chunk)
    If n > SPLASH_LOG_MAX_DISPLAY_CHARS Then
        tb.text = Right$(tb.text & chunk, SPLASH_LOG_MAX_DISPLAY_CHARS)
    Else
        tb.text = tb.text & chunk
    End If
    MacroSplash_TextBoxScrollToTail tb
End Sub

Private Sub アニメ付き_スプラッシュ付きで実行(ByVal splashMessage As String, ByVal procName As String, Optional ByVal arg1 As Variant, Optional ByVal arg2 As Variant, Optional ByVal lockExcelUI As Boolean = True, Optional ByVal allowMacroSound As Boolean = False)
    m_splashAllowMacroSound = allowMacroSound
    On Error GoTo EH
    MacroSplash_Show splashMessage, lockExcelUI
    If IsMissing(arg1) And IsMissing(arg2) Then
        Application.Run procName
    ElseIf Not IsMissing(arg1) And IsMissing(arg2) Then
        Application.Run procName, arg1
    Else
        Application.Run procName, arg1, arg2
    End If
    GoTo Finish
EH:
    On Error Resume Next
Finish:
    MacroStartBgm_FadeOutAndClose
    If m_animMacroSucceeded Then
        On Error Resume Next
        MacroCompleteChime
    End If
    MacroSplash_Hide
    m_splashAllowMacroSound = False
End Sub

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

Private Sub CreateCoolButtonWithPreset(btnText As String, macroName As String, posX As Single, posY As Single, ByVal presetId As Long)
    CreateCoolButton btnText, macroName, posX, posY, CoolButtonGradientTop(presetId), CoolButtonGradientBottom(presetId)
End Sub

Sub かっこいいボタンを作成()
    Dim y As Single
    Const gap As Single = 70
    
    y = 50
    CreateCoolButtonWithPreset "? シミュレーション実行", "アニメ付き_計画生成を実行", 50, y, 1
    y = y + gap
    CreateCoolButtonWithPreset "タスク抽出", "アニメ付き_タスク抽出を実行", 50, y, 3
    y = y + gap
    CreateCoolButtonWithPreset "段階1+2 連続", "アニメ付き_段階1と段階2を連続実行", 50, y, 5
    y = y + gap
    CreateCoolButtonWithPreset "環境構築 (初回のみ)", "アニメ付き_環境構築を実行", 50, y, 4
    y = y + gap
    CreateCoolButtonWithPreset "Gemini鍵を暗号化", "アニメ付き_Gemini認証を暗号化してB1に保存", 50, y, 6
    
    MsgBox "現在のシートにボタンを 5 つ作成しました！" & vbCrLf & _
           "グラデーションはプリセット 1/3/5/4 を使用しています（全 10 色はコード先頭のコメント参照）。" & vbCrLf & _
           "好きな場所にドラッグして配置してください。", vbInformation
End Sub

Sub かっこいいボタン_配色サンプル作成()
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
        CreateCoolButton "P" & CStr(i), "かっこいいボタンを作成", x, y, CoolButtonGradientTop(i), CoolButtonGradientBottom(i)
        On Error Resume Next
        ActiveSheet.Shapes(ActiveSheet.Shapes.Count).OnAction = ""
        On Error GoTo 0
    Next i
    MsgBox "配色プリセット P1～P10 の見本を配置しました。" & vbCrLf & _
           "クリックしてもマクロは動きません。不要なら図形を削除してください。", vbInformation
End Sub

Private Sub CreateCoolButton(btnText As String, macroName As String, posX As Single, posY As Single, colorTop As Long, colorBottom As Long)
    Dim shp As Shape
    
    Set shp = ActiveSheet.Shapes.AddShape(msoShapeRoundedRectangle, posX, posY, 220, 50)
    
    With shp
        With .TextFrame2.TextRange
            .text = btnText
            .Font.Name = "メイリオ"
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
        
        .line.Visible = msoFalse
        
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
        Randomize
        .Name = "CoolBtn_" & Format(Now, "yyyymmddhhnnss") & "_" & Format(Int(1000000 * Rnd), "000000")
        On Error GoTo 0
    End With
End Sub

Private Function Py3VersionOutput(wsh As Object) As String
    Dim execObj As Object
    Dim s As String
    Py3VersionOutput = ""
    On Error GoTo CleanExit
    Set execObj = wsh.Exec("cmd.exe /c py -3 --version")
    Do While execObj.Status = 0
        Sleep 50
    Loop
    s = execObj.StdOut.ReadAll()
    If Len(Trim$(s)) = 0 Then s = execObj.StdErr.ReadAll()
    Py3VersionOutput = s
CleanExit:
End Function

Private Function IsPython3Available(wsh As Object) As Boolean
    Dim s As String
    s = Py3VersionOutput(wsh)
    IsPython3Available = (InStr(1, s, "Python 3", vbTextCompare) > 0)
End Function

Private Function TryInstallPythonViaWinget(wsh As Object) As Boolean
    On Error GoTo Fail
    Dim wingetBat As String
    wingetBat = "@echo off" & vbCrLf & "winget install -e --id Python.Python.3.12 --silent --accept-package-agreements --accept-source-agreements" & vbCrLf & "exit /b %ERRORLEVEL%"
    RunTempCmdWithConsoleLayout wsh, wingetBat
    TryInstallPythonViaWinget = True
    Exit Function
Fail:
    TryInstallPythonViaWinget = False
End Function

Private Function TryInstallPythonViaOfficialInstaller(wsh As Object) As Boolean
    Dim psCmd As String
    Dim shellCmd As String
    Dim exitCode As Long
    On Error GoTo Fail
    psCmd = "powershell -NoProfile -ExecutionPolicy Bypass -Command """ & _
            "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; " & _
            "$url = '" & PY_OFFICIAL_INSTALLER_URL & "'; " & _
            "$out = Join-Path $env:TEMP 'python_official_installer.exe'; " & _
            "Invoke-WebRequest -Uri $url -OutFile $out -UseBasicParsing; " & _
            "if ((Get-Item $out).Length -lt 1MB) { throw 'ダウンロードに失敗した可能性があります' }; " & _
            "$p = Start-Process -FilePath $out -ArgumentList '/quiet InstallAllUsers=1 PrependPath=1 Include_test=0 Include_pip=1 Include_launcher=1' -Wait -PassThru; " & _
            "Remove-Item $out -ErrorAction SilentlyContinue; " & _
            "exit $p.ExitCode"""
    shellCmd = "cmd.exe /c " & psCmd
    exitCode = wsh.Run(shellCmd, 1, True)
    TryInstallPythonViaOfficialInstaller = (exitCode = 0)
    Exit Function
Fail:
    TryInstallPythonViaOfficialInstaller = False
End Function

Private Function RunPipInstallWithRefreshedPath(wsh As Object, ByVal workDir As String, ByVal setupRel As String) As Long
    Dim ps As String
    Dim shellCmd As String
    Dim wdEsc As String
    Dim setupEsc As String
    ' PATH を再合成したうえで、ブックフォルダで setup_environment.py を実行（pip + requirements + xlwings addin）
    wdEsc = Replace(workDir, "'", "''")
    setupEsc = Replace(setupRel, "'", "''")
    ps = "$env:Path = [System.Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [System.Environment]::GetEnvironmentVariable('Path','User'); " & _
         "$py = Get-Command py -ErrorAction SilentlyContinue; " & _
         "if (-not $py) { Write-Error 'py が見つかりません。Excel を一度終了してから再実行するか、PATH を確認してください。'; exit 91 }; " & _
         "Set-Location -LiteralPath '" & wdEsc & "'; " & _
         "& py -3 -u .\" & setupEsc & "; " & _
         "exit $LASTEXITCODE"
    shellCmd = "powershell -NoProfile -ExecutionPolicy Bypass -Command " & Chr(34) & ps & Chr(34)
    RunPipInstallWithRefreshedPath = wsh.Run(shellCmd, 1, True)
End Function

Private Sub DisableBackgroundDataRefreshAll()
    Dim wb As Workbook
    Dim cn As WorkbookConnection
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim pt As PivotTable
    Set wb = ThisWorkbook
    On Error Resume Next
    For Each cn In wb.Connections
        cn.OLEDBConnection.BackgroundQuery = False
        cn.ODBCConnection.BackgroundQuery = False
    Next cn
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            lo.QueryTable.BackgroundQuery = False
        Next lo
        For Each pt In ws.PivotTables
            pt.PivotCache.BackgroundQuery = False
        Next pt
    Next ws
    On Error GoTo 0
End Sub

Private Function PingHostOnceBeforeQueryRefresh(ByVal ipAddress As String, ByVal timeoutMs As Long) As Boolean
    Dim wsh As Object
    Dim cmd As String
    Dim rc As Long
    On Error GoTo EH
    If Len(ipAddress) = 0 Then PingHostOnceBeforeQueryRefresh = False: Exit Function
    Set wsh = CreateObject("WScript.Shell")
    cmd = "cmd /c ping -n 1 -w " & CStr(timeoutMs) & " " & ipAddress
    rc = wsh.Run(cmd, 0, True)
    PingHostOnceBeforeQueryRefresh = (rc = 0)
    Exit Function
EH:
    PingHostOnceBeforeQueryRefresh = False
End Function

Private Function TryRefreshWorkbookQueries() As Boolean
    Dim prevSU As Boolean
    Dim prevDA As Boolean
    On Error GoTo EH
    m_lastRefreshQueriesErrMsg = vbNullString
    prevSU = Application.ScreenUpdating
    prevDA = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    If SKIP_WORKBOOK_REFRESH_ALL Then
        Application.StatusBar = "（SKIP_WORKBOOK_REFRESH_ALL）接続の一括更新を省略しました"
        DoEvents
        Application.StatusBar = False
    ElseIf Not PingHostOnceBeforeQueryRefresh(PQ_REFRESH_PING_HOST, PQ_REFRESH_PING_TIMEOUT_MS) Then
        Application.StatusBar = "接続先 " & PQ_REFRESH_PING_HOST & " に ping 応答なし（" & CStr(PQ_REFRESH_PING_TIMEOUT_MS) & "ms）? Power Query 等の一括更新をスキップして処理を続行します"
        DoEvents
        Application.StatusBar = False
    Else
        Application.StatusBar = "データ接続を更新しています（完了までお待ちください）..."
        DoEvents
        Call DisableBackgroundDataRefreshAll
        ThisWorkbook.RefreshAll
        Application.CalculateUntilAsyncQueriesDone
        Application.StatusBar = False
    End If
    Application.DisplayAlerts = prevDA
    Application.ScreenUpdating = prevSU
    TryRefreshWorkbookQueries = True
    Exit Function
EH:
    Application.StatusBar = False
    On Error Resume Next
    Application.DisplayAlerts = prevDA
    Application.ScreenUpdating = prevSU
    On Error GoTo 0
    m_lastRefreshQueriesErrMsg = "データの更新（Power Query / 接続）: " & Err.Description
    TryRefreshWorkbookQueries = False
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

Private Function EscapeExcelFormulaText(ByVal s As String) As String
    If Len(s) > 0 Then
        If Left$(s, 1) = "=" Then
            EscapeExcelFormulaText = "'" & s
            Exit Function
        End If
    End If
    EscapeExcelFormulaText = s
End Function

Public Sub 設定_配台不要工程_シートを確保()
    Dim ws As Worksheet
    Dim sh As Worksheet
    Dim prevDA As Boolean

    On Error GoTo ErrHandler
    prevDA = Application.DisplayAlerts
    Application.DisplayAlerts = False

    Set ws = Nothing
    For Each sh In ThisWorkbook.Worksheets
        If StrComp(sh.Name, SHEET_EXCLUDE_ASSIGNMENT, vbBinaryCompare) = 0 Then
            Set ws = sh
            Exit For
        End If
    Next sh

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_EXCLUDE_ASSIGNMENT
    End If

    If StrComp(ws.Name, SHEET_EXCLUDE_ASSIGNMENT, vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 524, , "シート名を「" & SHEET_EXCLUDE_ASSIGNMENT & "」にできません（現在の名前: " & ws.Name & "）。同名シートや禁則文字を確認してください。"
    End If

    ' 非常に非表示のシートはタブに出ないため、確保時に必ず表示へ戻す
    ws.Visible = xlSheetVisible

    ' 1 行目は常に planning_core と同一見出し（空・不一致でも確実に揃える）
    ws.Cells(1, 1).Value = "工程名"
    ws.Cells(1, 2).Value = "機械名"
    ws.Cells(1, 3).Value = "配台不要"
    ws.Cells(1, 4).Value = "配台不能ロジック"
    ws.Cells(1, 5).Value = "ロジック式"

    Application.DisplayAlerts = prevDA
    Exit Sub
ErrHandler:
    Application.DisplayAlerts = prevDA
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function 設定_環境変数_1行目は見出し(ByVal ws As Worksheet) As Boolean
    Dim t As String
    t = LCase$(Trim$(CStr(ws.Cells(1, 1).Value)))
    If Len(t) = 0 Then
        設定_環境変数_1行目は見出し = False
        Exit Function
    End If
    設定_環境変数_1行目は見出し = (t = "変数名" Or t = "環境変数" Or t = "name" Or t = "key" Or t = "env")
End Function

Private Sub 設定_環境変数_欠損行を試し追記(ByVal dict As Object, ByVal ws As Worksheet, ByRef lastRow As Long, ByVal envKey As String, ByVal envVal As String, ByVal envDesc As String)
    Dim nk As String
    nk = LCase$(Trim$(envKey))
    If Len(nk) = 0 Then Exit Sub
    If dict.Exists(nk) Then Exit Sub
    lastRow = lastRow + 1
    ws.Cells(lastRow, 1).Value = envKey
    ws.Cells(lastRow, 2).Value = envVal
    ws.Cells(lastRow, 3).Value = envDesc
    dict.Add nk, True
End Sub

Public Sub 設定_環境変数_シートを確保()
    Dim ws As Worksheet
    Dim sh As Worksheet
    Dim prevDA As Boolean
    Dim dict As Object
    Dim r As Long
    Dim lastRow As Long
    Dim dataStart As Long
    Dim k As String
    Dim nk As String

    On Error GoTo ErrHandler
    prevDA = Application.DisplayAlerts
    Application.DisplayAlerts = False

    Set ws = Nothing
    For Each sh In ThisWorkbook.Worksheets
        If StrComp(sh.Name, SHEET_WORKBOOK_ENV, vbBinaryCompare) = 0 Then
            Set ws = sh
            Exit For
        End If
    Next sh

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_WORKBOOK_ENV
    End If

    If StrComp(ws.Name, SHEET_WORKBOOK_ENV, vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 525, , "シート名を「" & SHEET_WORKBOOK_ENV & "」にできません（現在: " & ws.Name & "）。"
    End If

    ws.Visible = xlSheetVisible

    If Len(Trim$(CStr(ws.Cells(1, 1).Value))) = 0 Then
        ws.Cells(1, 1).Value = "変数名"
        ws.Cells(1, 2).Value = "値"
        ws.Cells(1, 3).Value = "説明（任意）"
    ElseIf Not 設定_環境変数_1行目は見出し(ws) Then
        ' 1 行目がデータの場合は見出しを挿入しない（ユーザー構成を壊さない）
    Else
        ws.Cells(1, 1).Value = "変数名"
        ws.Cells(1, 2).Value = "値"
        ws.Cells(1, 3).Value = "説明（任意）"
    End If

    If 設定_環境変数_1行目は見出し(ws) Then
        dataStart = 2
    Else
        dataStart = 1
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1 ' vbTextCompare（Windows 環境変数は実質大文字小文字無視）

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < dataStart Then
        lastRow = dataStart - 1
    End If

    For r = dataStart To lastRow
        k = Trim$(CStr(ws.Cells(r, 1).Value))
        If Len(k) > 0 Then
            nk = LCase$(k)
            If Not dict.Exists(nk) Then dict.Add nk, True
        End If
    Next r

    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "TASK_PLAN_SHEET", "", "配台計画シート名（空なら既定 配台計画_タスク入力）")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "STAGE2_DISPATCH_FLOW_TRIAL_ORDER_FIRST", "1", "日内配台: 1=試行順優先マルチパス（既定） 0=従来ソート")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "STAGE2_SERIAL_DISPATCH_BY_TASK_ID", "0", "1=依頼NO直列")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "STAGE12_CMD_HIDE_WINDOW", "1", "段階1/2: cmd 1=非表示(既定) 0=画面上部にコンソール")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "STAGE1_SYNC_MASTER_SHEETS_TO_MACRO_BOOK", "0", "段階1: master から機械カレンダー・勤怠をマクロブックへコピー 1=する 0=しない（配台は master 直読み）")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "PLANNING_B1_INSPECTION_EXCLUSIVE_MACHINE", "1", "§B-2/§B-3 設備占有（0 で無効）")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "PLANNING_B2_EC_FOLLOWER_DISJOINT_TEAMS", "1", "B-2/3 ECと後続の担当者分離（0 で無効）")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "TEAM_ASSIGN_PRIORITIZE_SURPLUS_STAFF", "0", "1=人数最優先（従来）")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "TEAM_ASSIGN_START_SLACK_WAIT_MINUTES", "60", "スラック分（0 で開始のみ）")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "TEAM_ASSIGN_IGNORE_NEED_SURPLUS_ROW", "0", "1=need追加人数行を無視")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "TEAM_ASSIGN_USE_NEED_SURPLUS_IN_MAIN_PASS", "", "1=メインで req+追加上限まで探索")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "TEAM_ASSIGN_HEADCOUNT_FROM_NEED_ONLY", "1", "0=計画シート必要人数も参照")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "TEAM_ASSIGN_USE_MASTER_COMBO_SHEET", "1", "0=組合せ表プリセットを使わない")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "STAGE2_COPY_COLUMN_CONFIG_SHAPES_FROM_INPUT", "1", "段階2後の列設定図形コピー（0 で無効）")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "EXCLUDE_RULES_TRY_OPENPYXL_SAVE", "", "1=配台不要シートを openpyxl で保存試行")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "ASSIGN_EOD_DEFER_MAX_REMAINING_ROLLS", "5", "終業直前デファー対象の最大残ロール")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "ASSIGN_END_OF_DAY_DEFER_MINUTES", "45", "終業直前デファー分数（0で明示無効・既定45）")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "DEBUG_TASK_ID", "", "デバッグ用依頼NO")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "TRACE_TEAM_ASSIGN_TASK_ID", "", "チーム割付トレース対象")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "GEMINI_PRICE_USD_IN_PER_M", "0.075", "")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "GEMINI_PRICE_USD_OUT_PER_M", "0.30", "")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "GEMINI_JPY_PER_USD", "150", "")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "EXCLUDE_RULES_TEST_E1234", "", "テスト用（通常空）")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "EXCLUDE_RULES_TEST_E1234_ROW", "9", "")
    Call 設定_環境変数_欠損行を試し追記(dict, ws, lastRow, "#TASK_INPUT_WORKBOOK", "", "通常はVBAが設定（シートに書くと上書き）。先頭#行は Python 側でコメント扱い")

    ws.Columns(1).ColumnWidth = 28
    ws.Columns(2).ColumnWidth = 14
    ws.Columns(3).ColumnWidth = 52

    Application.DisplayAlerts = prevDA
    Exit Sub
ErrHandler:
    Application.DisplayAlerts = prevDA
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function 設定_シート表示_C列を表示状態に変換(ByVal s As String) As XlSheetVisibility
    Dim t As String
    t = LCase$(Trim$(s))
    If Len(t) = 0 Then
        設定_シート表示_C列を表示状態に変換 = xlSheetVisible
        Exit Function
    End If
    If t = "表示" Or t = "true" Or t = "1" Or t = "yes" Or t = "on" Or t = "y" Or t = "はい" Or t = "visible" Then
        設定_シート表示_C列を表示状態に変換 = xlSheetVisible
        Exit Function
    End If
    If t = "非表示" Or t = "hidden" Or t = "0" Or t = "false" Or t = "no" Or t = "off" Or t = "n" Or t = "いいえ" Then
        設定_シート表示_C列を表示状態に変換 = xlSheetHidden
        Exit Function
    End If
    If t = "完全非表示" Or t = "非常隠し" Or t = "veryhidden" Or t = "xlveryhidden" Or t = "2" Then
        設定_シート表示_C列を表示状態に変換 = xlSheetVeryHidden
        Exit Function
    End If
    設定_シート表示_C列を表示状態に変換 = xlSheetVisible
End Function

Private Function 設定_シート表示_表示状態の説明文字列(ByVal vis As XlSheetVisibility) As String
    Select Case vis
        Case xlSheetVisible
            設定_シート表示_表示状態の説明文字列 = "表示"
        Case xlSheetHidden
            設定_シート表示_表示状態の説明文字列 = "非表示"
        Case xlSheetVeryHidden
            設定_シート表示_表示状態の説明文字列 = "完全非表示"
        Case Else
            設定_シート表示_表示状態の説明文字列 = "表示"
    End Select
End Function

Private Sub 設定_シート表示_ドロップダウン候補セルを書く(ByVal ws As Worksheet)
    ws.Range("F1").Value = "（C列の候補・参照用）"
    ws.Range("F2").Value = "表示"
    ws.Range("F3").Value = "非表示"
    ws.Range("F4").Value = "完全非表示"
    ws.Columns(6).ColumnWidth = 14
End Sub

Private Sub 設定_シート表示_C列入力規則を付与(ByVal ws As Worksheet)
    Const NM_SHEET_VIS As String = "PM_AI_SheetVisList"
    Dim rng As Range
    Set rng = ws.Range("C2:C1000")
    On Error Resume Next
    rng.Validation.Delete
    Err.Clear
    rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="表示,非表示,完全非表示"
    If Err.Number = 0 Then GoTo ApplyVisFlags
    Err.Clear
    ThisWorkbook.names(NM_SHEET_VIS).Delete
    Err.Clear
    ThisWorkbook.names.Add Name:=NM_SHEET_VIS, RefersTo:=ws.Range("F2:F4")
    If Err.Number = 0 Then
        Err.Clear
        rng.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & NM_SHEET_VIS
    End If
ApplyVisFlags:
    On Error Resume Next
    rng.Validation.IgnoreBlank = True
    rng.Validation.InCellDropdown = True
    Err.Clear
    On Error GoTo 0
End Sub

Private Function 設定_シート表示_C列を正規化表示文字列(ByVal raw As String, ByVal fallbackVis As XlSheetVisibility) As String
    If Len(Trim$(raw)) = 0 Then
        設定_シート表示_C列を正規化表示文字列 = 設定_シート表示_表示状態の説明文字列(fallbackVis)
    Else
        設定_シート表示_C列を正規化表示文字列 = 設定_シート表示_表示状態の説明文字列(設定_シート表示_C列を表示状態に変換(raw))
    End If
End Function

Public Sub 設定_シート表示_シートを確保()
    Dim ws As Worksheet
    Dim sh As Worksheet
    Dim prevDA As Boolean

    On Error GoTo ErrHandler
    prevDA = Application.DisplayAlerts
    Application.DisplayAlerts = False

    Set ws = Nothing
    For Each sh In ThisWorkbook.Worksheets
        If StrComp(sh.Name, SHEET_SHEET_VISIBILITY, vbBinaryCompare) = 0 Then
            Set ws = sh
            Exit For
        End If
    Next sh

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_SHEET_VISIBILITY
    End If

    If StrComp(ws.Name, SHEET_SHEET_VISIBILITY, vbBinaryCompare) <> 0 Then
        Err.Raise vbObjectError + 531, , "シート名を「" & SHEET_SHEET_VISIBILITY & "」にできません（現在: " & ws.Name & "）。"
    End If

    ws.Visible = xlSheetVisible

    ws.Cells(1, 1).Value = "並び順"
    ws.Cells(1, 2).Value = "シート名"
    ws.Cells(1, 3).Value = "表示"
    ws.Cells(1, 4).Value = "（手順）一覧をブックから再取得 → 並び順(1始まり)・表示を編集 → ブックへ適用（適用後は一覧が自動更新）"

    ws.Columns(1).ColumnWidth = 10
    ws.Columns(2).ColumnWidth = 36
    ws.Columns(3).ColumnWidth = 14
    ws.Columns(4).ColumnWidth = 62

    Call 設定_シート表示_ドロップダウン候補セルを書く(ws)
    Call 設定_シート表示_C列入力規則を付与(ws)

    Application.DisplayAlerts = prevDA
    Exit Sub
ErrHandler:
    Application.DisplayAlerts = prevDA
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub 設定_シート表示_一覧をブックから再取得()
    Dim wb As Workbook
    Dim wsCfg As Worksheet
    Dim orderDict As Object
    Dim visDict As Object
    Dim lastRow As Long
    Dim r As Long
    Dim nm As String
    Dim ordVal As Double
    Dim maxOrder As Double
    Dim ws As Worksheet
    Dim prevSU As Boolean
    Dim prevDA As Boolean
    Dim n As Long
    Dim i As Long
    Dim j As Long
    Dim sortKey() As Double
    Dim sheetName() As String
    Dim visText() As String
    Dim origIdx() As Long
    Dim tmpK As Double
    Dim tmpN As String
    Dim tmpV As String
    Dim tmpO As Long

    Call 設定_シート表示_シートを確保
    Set wb = ThisWorkbook
    Set wsCfg = wb.Worksheets(SHEET_SHEET_VISIBILITY)

    Set orderDict = CreateObject("Scripting.Dictionary")
    orderDict.CompareMode = 1
    Set visDict = CreateObject("Scripting.Dictionary")
    visDict.CompareMode = 1

    lastRow = wsCfg.Cells(wsCfg.Rows.Count, 2).End(xlUp).Row
    If lastRow < 2 Then lastRow = 1
    maxOrder = 0
    For r = 2 To lastRow
        nm = Trim$(CStr(wsCfg.Cells(r, 2).Value))
        If Len(nm) > 0 Then
            If Not orderDict.Exists(nm) Then
                ordVal = 0
                Err.Clear
                On Error Resume Next
                ordVal = CDbl(wsCfg.Cells(r, 1).Value)
                If Err.Number <> 0 Then ordVal = 0
                On Error GoTo 0
                If ordVal <= 0 Then ordVal = CDbl(1000000# + r)
                orderDict.Add nm, ordVal
                If ordVal > maxOrder Then maxOrder = ordVal
                visDict.Add nm, Trim$(CStr(wsCfg.Cells(r, 3).Value))
            End If
        End If
    Next r

    prevSU = Application.ScreenUpdating
    prevDA = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    wsCfg.Range("A2:D" & wsCfg.Rows.Count).ClearContents

    n = wb.Worksheets.Count
    If n > 0 Then
        ReDim sortKey(1 To n)
        ReDim sheetName(1 To n)
        ReDim visText(1 To n)
        ReDim origIdx(1 To n)
    End If

    For i = 1 To n
        Set ws = wb.Worksheets(i)
        nm = ws.Name
        origIdx(i) = i
        sheetName(i) = nm
        If orderDict.Exists(nm) Then
            sortKey(i) = orderDict(nm)
        Else
            maxOrder = maxOrder + 1
            sortKey(i) = maxOrder
        End If
        If visDict.Exists(nm) And Len(Trim$(CStr(visDict(nm)))) > 0 Then
            visText(i) = 設定_シート表示_C列を正規化表示文字列(CStr(visDict(nm)), ws.Visible)
        Else
            visText(i) = 設定_シート表示_表示状態の説明文字列(ws.Visible)
        End If
    Next i

    For i = 1 To n - 1
        For j = i + 1 To n
            If sortKey(i) > sortKey(j) Or (sortKey(i) = sortKey(j) And origIdx(i) > origIdx(j)) Then
                tmpK = sortKey(i): sortKey(i) = sortKey(j): sortKey(j) = tmpK
                tmpN = sheetName(i): sheetName(i) = sheetName(j): sheetName(j) = tmpN
                tmpV = visText(i): visText(i) = visText(j): visText(j) = tmpV
                tmpO = origIdx(i): origIdx(i) = origIdx(j): origIdx(j) = tmpO
            End If
        Next j
    Next i

    For i = 1 To n
        wsCfg.Cells(i + 1, 1).Value = i
        wsCfg.Cells(i + 1, 2).Value = sheetName(i)
        wsCfg.Cells(i + 1, 3).Value = visText(i)
    Next i

    Call 設定_シート表示_ドロップダウン候補セルを書く(wsCfg)
    Call 設定_シート表示_C列入力規則を付与(wsCfg)

    Application.DisplayAlerts = prevDA
    Application.ScreenUpdating = prevSU
End Sub

Public Sub 設定_シート表示_ブックへ適用()
    Dim wb As Workbook
    Dim wsCfg As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim nm As String
    Dim ordVal As Double
    Dim listed As Object
    Dim orderList() As Double
    Dim nameList() As String
    Dim rowList() As Long
    Dim nListed As Long
    Dim i As Long
    Dim j As Long
    Dim tmpD As Double
    Dim tmpS As String
    Dim tmpR As Long
    Dim vis As XlSheetVisibility
    Dim cntVis As Long
    Dim nFull As Long
    Dim wi As Long
    Dim prevSU As Boolean
    Dim prevDA As Boolean
    Dim testWs As Worksheet

    On Error GoTo ErrHandler
    Call 設定_シート表示_シートを確保
    Set wb = ThisWorkbook
    Set wsCfg = wb.Worksheets(SHEET_SHEET_VISIBILITY)

    Set listed = CreateObject("Scripting.Dictionary")
    listed.CompareMode = 1

    lastRow = wsCfg.Cells(wsCfg.Rows.Count, 2).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "「" & SHEET_SHEET_VISIBILITY & "」にデータ行（2行目以降）がありません。先に「設定_シート表示_一覧をブックから再取得」を実行してください。", vbExclamation, "設定_シート表示"
        Exit Sub
    End If

    nListed = 0
    For r = 2 To lastRow
        nm = Trim$(CStr(wsCfg.Cells(r, 2).Value))
        If Len(nm) > 0 Then
            Set testWs = Nothing
            Err.Clear
            On Error Resume Next
            Set testWs = wb.Worksheets(nm)
            If Err.Number = 0 And Not testWs Is Nothing Then
                If Not listed.Exists(nm) Then
                    nListed = nListed + 1
                    ReDim Preserve orderList(1 To nListed)
                    ReDim Preserve nameList(1 To nListed)
                    ReDim Preserve rowList(1 To nListed)
                    ordVal = 0
                    On Error Resume Next
                    ordVal = CDbl(wsCfg.Cells(r, 1).Value)
                    If Err.Number <> 0 Then ordVal = 0
                    If ordVal <= 0 Then ordVal = CDbl(r + 10000)
                    orderList(nListed) = ordVal
                    nameList(nListed) = nm
                    rowList(nListed) = r
                    listed.Add nm, True
                End If
            End If
            Err.Clear
            On Error GoTo ErrHandler
        End If
    Next r

    If nListed = 0 Then
        MsgBox "有効なシート名の行がありません。", vbExclamation, "設定_シート表示"
        Exit Sub
    End If

    ' 同順位は元の行番号で安定ソート
    For i = 1 To nListed - 1
        For j = i + 1 To nListed
            If orderList(i) > orderList(j) Or (orderList(i) = orderList(j) And rowList(i) > rowList(j)) Then
                tmpD = orderList(i): orderList(i) = orderList(j): orderList(j) = tmpD
                tmpS = nameList(i): nameList(i) = nameList(j): nameList(j) = tmpS
                tmpR = rowList(i): rowList(i) = rowList(j): rowList(j) = tmpR
            End If
        Next j
    Next i

    cntVis = 0
    For i = 1 To nListed
        nm = nameList(i)
        r = rowList(i)
        vis = 設定_シート表示_C列を表示状態に変換(CStr(wsCfg.Cells(r, 3).Value))
        If StrComp(nm, SHEET_SHEET_VISIBILITY, vbBinaryCompare) = 0 Then
            vis = xlSheetVisible
        End If
        If vis = xlSheetVisible Then cntVis = cntVis + 1
    Next i

    For wi = 1 To wb.Worksheets.Count
        nm = wb.Worksheets(wi).Name
        If Not listed.Exists(nm) Then
            If wb.Worksheets(nm).Visible = xlSheetVisible Then cntVis = cntVis + 1
        End If
    Next wi

    If cntVis < 1 Then
        MsgBox "この内容では表示されるシートが 0 になります。Excel の制約のため中止しました。", vbCritical, "設定_シート表示"
        Exit Sub
    End If

    prevSU = Application.ScreenUpdating
    prevDA = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For i = 1 To nListed
        nm = nameList(i)
        r = rowList(i)
        vis = 設定_シート表示_C列を表示状態に変換(CStr(wsCfg.Cells(r, 3).Value))
        If StrComp(nm, SHEET_SHEET_VISIBILITY, vbBinaryCompare) = 0 Then
            vis = xlSheetVisible
        End If
        On Error Resume Next
        wb.Worksheets(nm).Visible = vis
        Err.Clear
        On Error GoTo ErrHandler
    Next i

    nFull = nListed
    For wi = 1 To wb.Worksheets.Count
        nm = wb.Worksheets(wi).Name
        If Not listed.Exists(nm) Then
            nFull = nFull + 1
            ReDim Preserve nameList(1 To nFull)
            nameList(nFull) = nm
        End If
    Next wi

    If nFull <> wb.Worksheets.Count Then
        Application.DisplayAlerts = prevDA
        Application.ScreenUpdating = prevSU
        Err.Raise vbObjectError + 532, , "内部エラー: シート数と並びリストが一致しません。"
    End If

    For i = 1 To nFull
        On Error Resume Next
        wb.Worksheets(nameList(i)).Move After:=wb.Sheets(wb.Sheets.Count)
        Err.Clear
        On Error GoTo ErrHandler
    Next i

    Application.DisplayAlerts = prevDA
    Application.ScreenUpdating = prevSU

    ' タブ順・表示を反映したあと、表の並びと A 列連番をブック現状に合わせる（失敗しても適用は維持）
    On Error Resume Next
    設定_シート表示_一覧をブックから再取得
    Err.Clear
    On Error GoTo 0

    Exit Sub
ErrHandler:
    Application.DisplayAlerts = prevDA
    Application.ScreenUpdating = prevSU
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function NormalizeWorkbookPathForCompare(ByVal p As String) As String
    NormalizeWorkbookPathForCompare = LCase$(Replace(Replace(Trim$(p), "/", "\"), vbTab, ""))
End Function

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

Private Function 段階1_マスタシートを本ブックへ置換コピー( _
    ByVal srcWb As Workbook, _
    ByVal sheetName As String) As Boolean
    Dim srcWs As Worksheet
    Dim wsOld As Worksheet
    Dim destWs As Worksheet
    Dim da As Boolean
    
    段階1_マスタシートを本ブックへ置換コピー = False
    On Error Resume Next
    Set srcWs = srcWb.Worksheets(sheetName)
    On Error GoTo 0
    If srcWs Is Nothing Then Exit Function
    
    da = Application.DisplayAlerts
    Application.DisplayAlerts = False
    On Error Resume Next
    Set wsOld = ThisWorkbook.Worksheets(sheetName)
    If Not wsOld Is Nothing Then
        If ThisWorkbook.Sheets.Count <= 1 Then
            Application.DisplayAlerts = da
            Exit Function
        End If
        wsOld.Unprotect
        If Not ThisWorkbook.ActiveSheet Is Nothing Then
            If StrComp(ThisWorkbook.ActiveSheet.Name, sheetName, vbBinaryCompare) = 0 Then
                ThisWorkbook.Worksheets(1).Activate
            End If
        End If
        wsOld.Delete
        Set wsOld = Nothing
    End If
    On Error GoTo CopySheetFailSt1
    
    srcWs.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    Set destWs = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
    On Error Resume Next
    destWs.Name = sheetName
    On Error GoTo 0
    
    ' 保護は段階1/2 マクロ終了時に 配台マクロ_対象シートを条件どおりに保護 でまとめて適用（処理中は全シート解除済み）
    
    Application.DisplayAlerts = da
    段階1_マスタシートを本ブックへ置換コピー = True
    Exit Function
CopySheetFailSt1:
    Application.DisplayAlerts = da
End Function

Private Sub 配台マクロ_全シート保護を試行解除()
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If ws.ProtectContents Then
            ws.Unprotect
            If ws.ProtectContents And Len(SHEET_FONT_UNPROTECT_PASSWORD) > 0 Then
                ws.Unprotect Password:=SHEET_FONT_UNPROTECT_PASSWORD
            End If
        End If
    Next ws
    On Error GoTo 0
End Sub

Private Sub 配台マクロ_対象シートを条件どおりに保護(Optional ByVal targetDir As String = "")
    Dim td As String
    Dim ws As Worksheet
    Dim nm As Variant
    Dim dict As Object
    
    On Error Resume Next
    td = targetDir
    If Len(td) = 0 Then td = ThisWorkbook.path
    
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, 3) = "結果_" Then
            If ws.ProtectContents Then ws.Unprotect
            If StrComp(ws.Name, SHEET_RESULT_EQUIP_GANTT, vbBinaryCompare) = 0 Then
                ws.Protect DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFormattingCells:=True
            Else
                ws.Protect DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True
            End If
        End If
    Next ws
    
    Set dict = CreateObject("Scripting.Dictionary")
    If Not dict.Exists(SHEET_MACHINE_CALENDAR) Then dict.Add SHEET_MACHINE_CALENDAR, True
    If Len(td) > 0 Then
        配台_マスタSkillsから勤怠シート名を辞書に追加 td, dict
    End If
    For Each nm In dict.keys
        Set ws = Nothing
        Set ws = ThisWorkbook.Worksheets(CStr(nm))
        If Not ws Is Nothing Then
            If ws.ProtectContents Then ws.Unprotect
            ws.Protect DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True
        End If
    Next nm
    
    On Error GoTo 0
End Sub

Public Sub RunPythonStage1()
    段階1_コア実行
    On Error Resume Next
    配台計画_タスク入力_A1を選択
    On Error GoTo 0
    If m_lastStage1ExitCode < 0 Then
        If Len(m_lastStage1ErrMsg) > 0 Then MsgBox m_lastStage1ErrMsg, vbExclamation, "段階1"
        Exit Sub
    End If
    If m_lastStage1ExitCode <> 0 Then
        Dim st1Block As String
        st1Block = Trim$(GeminiReadUtf8File(ThisWorkbook.path & "\log\stage2_blocking_message.txt"))
        If m_lastStage1ExitCode = 3 And Len(st1Block) > 0 Then
            MsgBox st1Block, vbCritical, "段階1"
        Else
            MsgBox "段階1の Python 終了コードが " & CStr(m_lastStage1ExitCode) & " です。" & vbCrLf & "LOG シート・log\execution_log.txt を確認してください。（検証中止時は log\stage2_blocking_message.txt も参照）", vbExclamation, "段階1"
        End If
        Exit Sub
    End If
    MacroSplash_SetStep "段階1が完了しました。配台計画シートを確認のうえ、必要なら段階2（計画生成）を実行してください。"
    m_animMacroSucceeded = True
End Sub

Public Sub RunPythonStage1ThenStage2()
    段階1_コア実行
    On Error Resume Next
    配台計画_タスク入力_A1を選択
    On Error GoTo 0
    If m_lastStage1ExitCode < 0 Then
        If Len(m_lastStage1ErrMsg) > 0 Then MsgBox m_lastStage1ErrMsg, vbExclamation, "段階1+2"
        Exit Sub
    End If
    If m_lastStage1ExitCode <> 0 Then
        Dim st1b2 As String
        st1b2 = Trim$(GeminiReadUtf8File(ThisWorkbook.path & "\log\stage2_blocking_message.txt"))
        If m_lastStage1ExitCode = 3 And Len(st1b2) > 0 Then
            MsgBox st1b2, vbCritical, "段階1+2"
        Else
            MsgBox "段階1の Python 終了コードが " & CStr(m_lastStage1ExitCode) & " のため、段階2は実行しません。" & vbCrLf & "LOG シート・log\execution_log.txt を確認してください。（検証中止時は log\stage2_blocking_message.txt も参照）", vbExclamation, "段階1+2"
        End If
        Exit Sub
    End If
    段階2_コア実行 True
    If m_lastStage2ExitCode <> 0 Or Len(m_lastStage2ErrMsg) > 0 Then
        If Len(m_lastStage2ErrMsg) > 0 Then
            MsgBox m_lastStage2ErrMsg, vbCritical, "段階1+2"
        Else
            MsgBox "段階2の Python 終了コードが " & CStr(m_lastStage2ExitCode) & " です。LOG シート・log\execution_log.txt を確認してください。", vbExclamation, "段階1+2"
        End If
        Exit Sub
    End If
    段階2_取り込み結果を報告
    If m_stage2PlanImported Or m_stage2MemberImported Then m_animMacroSucceeded = True
End Sub

Private Sub 配台計画_タスク入力_A1を選択()
    Const PLAN_SHEET As String = "配台計画_タスク入力"
    Dim ws As Worksheet
    On Error Resume Next
    ThisWorkbook.Activate
    Set ws = ThisWorkbook.Sheets(PLAN_SHEET)
    If ws Is Nothing Then Exit Sub
    If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    ws.Activate
    ws.Range("A1").Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    On Error GoTo 0
End Sub

Private Sub 配台計画_タスク入力_上書き列に入力色を付与(ByVal ws As Worksheet)
    Dim headers As Variant
    Dim i As Long
    Dim c As Long
    Dim lastRow As Long
    Dim rng As Range
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    headers = Array( _
        "配台不要", _
        "加工速度_上書き", _
        "原反投入日_上書き", _
        "担当OP_指定", _
        "特別指定_備考")
    lastRow = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If lastRow < 2 Then Exit Sub
    For i = LBound(headers) To UBound(headers)
        c = FindColHeader(ws, CStr(headers(i)))
        If c > 0 Then
            Set rng = ws.Range(ws.Cells(2, c), ws.Cells(lastRow, c))
            rng.Interior.Color = RGB(255, 242, 204)
        End If
    Next i
    On Error GoTo 0
End Sub

Private Function 取込ブック内のコピー先シートを取得(ByVal wb As Workbook, ByVal expectedSheetName As String) As Worksheet
    Dim sh As Worksheet
    Dim si As Long
    
    On Error Resume Next
    Set sh = wb.Sheets(expectedSheetName)
    On Error GoTo 0
    If Not sh Is Nothing Then
        Set 取込ブック内のコピー先シートを取得 = sh
        Exit Function
    End If
    
    On Error Resume Next
    Set sh = wb.Sheets(wb.Sheets.Count)
    On Error GoTo 0
    
    If Not sh Is Nothing Then
        If StrComp(sh.Name, SCRATCH_SHEET_FONT, vbBinaryCompare) <> 0 Then
            Set 取込ブック内のコピー先シートを取得 = sh
            Exit Function
        End If
    End If
    
    For si = 1 To wb.Sheets.Count
        If StrComp(wb.Sheets(si).Name, expectedSheetName, vbBinaryCompare) = 0 Then
            Set 取込ブック内のコピー先シートを取得 = wb.Sheets(si)
            Exit Function
        End If
    Next si
    
    Set 取込ブック内のコピー先シートを取得 = sh
End Function

Private Function シート名は計画取込の同源名またはExcel番号付き複製か(ByVal nm As String, ByVal baseName As String) As Boolean
    Dim i As Long
    Dim j As Long
    Dim ch As String
    
    If StrComp(nm, baseName, vbBinaryCompare) = 0 Then
        シート名は計画取込の同源名またはExcel番号付き複製か = True
        Exit Function
    End If
    If Len(nm) <= Len(baseName) Then Exit Function
    If StrComp(Left$(nm, Len(baseName)), baseName, vbBinaryCompare) <> 0 Then Exit Function
    
    i = Len(baseName) + 1
    Do While i <= Len(nm)
        ch = Mid$(nm, i, 1)
        If ch <> " " And ch <> ChrW(&H3000) Then Exit Do
        i = i + 1
    Loop
    If i > Len(nm) Then Exit Function
    
    ch = Mid$(nm, i, 1)
    If ch <> "(" And ch <> ChrW(&HFF08) Then Exit Function
    
    i = i + 1
    If i > Len(nm) Then Exit Function
    j = i
    Do While j <= Len(nm)
        ch = Mid$(nm, j, 1)
        If ch < "0" Or ch > "9" Then Exit Do
        j = j + 1
    Loop
    If j = i Then Exit Function
    
    If j > Len(nm) Then Exit Function
    ch = Mid$(nm, j, 1)
    If ch <> ")" And ch <> ChrW(&HFF09) Then Exit Function
    
    j = j + 1
    Do While j <= Len(nm)
        ch = Mid$(nm, j, 1)
        If ch <> " " And ch <> ChrW(&H3000) Then Exit Function
        j = j + 1
    Loop
    
    シート名は計画取込の同源名またはExcel番号付き複製か = True
End Function

Private Sub マクロブックから計画取込シート同源名シートを削除(ByVal wb As Workbook, ByVal sheetName As String)
    Dim i As Long
    Dim j As Long
    Dim nm As String
    Dim names() As String
    Dim n As Long
    
    n = 0
    ReDim names(1 To wb.Sheets.Count)
    For i = 1 To wb.Sheets.Count
        nm = wb.Sheets(i).Name
        If シート名は計画取込の同源名またはExcel番号付き複製か(nm, sheetName) Then
            n = n + 1
            names(n) = nm
        End If
    Next i
    
    For j = n To 1 Step -1
        On Error Resume Next
        wb.Sheets(names(j)).Delete
        Err.Clear
        On Error GoTo 0
    Next j
End Sub

Private Sub 段階2_取り込み結果を報告()
    Dim p As String
    p = ThisWorkbook.path
    
    If m_stage2PlanImported And m_stage2MemberImported Then
        MacroSplash_SetStep "計画生成が完了しました（結果シートと個人シートを取り込みました）。"
    ElseIf m_stage2PlanImported Then
        MacroSplash_SetStep "計画生成が完了しました（結果シートのみ。個人別 member_schedule は見つかりませんでした）。"
    ElseIf m_stage2MemberImported Then
        MacroSplash_SetStep "計画生成が完了しました（個人シートのみ。production_plan は見つかりませんでした）。"
    Else
        MsgBox "Pythonの実行は完了しましたが、output フォルダに計画・個人別のいずれの xlsx も見つかりませんでした。" & vbCrLf & vbCrLf & _
               "Python 終了コード: " & CStr(m_lastStage2ExitCode) & vbCrLf & _
               IIf(Len(p) > 0, "探索したフォルダ: " & p & "\output", "ブックが未保存のため output パスを表示できません。先に保存してください。") & vbCrLf & vbCrLf & _
               "LOG シートまたは " & IIf(Len(p) > 0, p & "\log\execution_log.txt", "log\execution_log.txt（ブックと同じフォルダ）") & " で「段階2を中断」「マスタファイル」「メンバーが0人」等を確認してください。" & vbCrLf & _
               "（テストコード直下に master.xlsm が無いとメンバー0で中断し、xlsx は出力されません。）", vbExclamation, "計画生成"
    End If
End Sub

Public Sub RunPython(Optional ByVal preserveStage1LogOnLogSheet As Boolean = False)
    段階2_コア実行 preserveStage1LogOnLogSheet
    If m_lastStage2ExitCode <> 0 Or Len(m_lastStage2ErrMsg) > 0 Then
        If Len(m_lastStage2ErrMsg) > 0 Then
            MsgBox m_lastStage2ErrMsg, vbCritical, "計画生成"
        Else
            MsgBox "Python の終了コードが " & CStr(m_lastStage2ExitCode) & " です。LOG シート・log\execution_log.txt を確認してください。", vbExclamation, "計画生成"
        End If
        Exit Sub
    End If
    段階2_取り込み結果を報告
    If m_stage2PlanImported Or m_stage2MemberImported Then m_animMacroSucceeded = True
End Sub

Private Function SafePersonalSheetName(ByVal baseName As String) As String
    Dim s As String
    s = "個人_" & Trim$(baseName)
    ' Excel シート名に使えない文字を除去
    s = Replace(s, "\", "")
    s = Replace(s, "/", "")
    s = Replace(s, "?", "")
    s = Replace(s, "*", "")
    s = Replace(s, "[", "")
    s = Replace(s, "]", "")
    s = Replace(s, ":", "")
    If Len(s) = 0 Then s = "個人_Sheet"
    If Len(s) > 31 Then s = Left$(s, 31)
    SafePersonalSheetName = s
End Function

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
        ws.Range("A1").Value = "（フォント選択用・削除しないでください）"
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

Public Sub 列設定_結果_タスク一覧_チェックボックスを配置()
    Dim ws As Worksheet
    Dim r As Long
    Dim lastR As Long
    Dim cb As CheckBox
    Dim rng As Range
    Dim linkAddr As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_COL_CONFIG_RESULT_TASK)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "シート「" & SHEET_COL_CONFIG_RESULT_TASK & "」が見つかりません。", vbExclamation, "列設定"
        Exit Sub
    End If

    On Error GoTo FailChk
    Do While ws.CheckBoxes.Count > 0
        ws.CheckBoxes(1).Delete
    Loop

    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR < 2 Then
        MsgBox "データ行がありません（1行目は見出し、2行目以降に列名を入れてください）。", vbInformation
        Exit Sub
    End If

    linkAddr = "'" & Replace(ws.Name, "'", "''") & "'!"

    For r = 2 To lastR
        If Len(Trim$(CStr(ws.Cells(r, 1).Value))) = 0 Then GoTo NextLoop

        Set rng = ws.Cells(r, 2)
        If Len(Trim$(CStr(rng.Value))) = 0 Then
            rng.Value = True
        End If

        Set cb = ws.CheckBoxes.Add(rng.Left + 2, rng.Top + 0.5, 50, 14)
        With cb
            .LinkedCell = linkAddr & rng.Address(True, True)
            .Caption = ""
        End With
NextLoop:
    Next r

    MsgBox "チェックボックスを配置しました。" & vbCrLf & _
        "表示列(B)の TRUE/FALSE と連動します。", vbInformation
    Exit Sub

FailChk:
    MsgBox "チェックボックス配置でエラー: " & Err.Description, vbCritical
End Sub

Public Sub 列設定_結果_タスク一覧_列順表示をPython適用()
    Dim wsh As Object
    Dim runBat As String
    Dim targetDir As String
    Dim exitCode As Long
    Dim wsRes As Worksheet
    Dim wsCfg As Worksheet
    Dim prevScreen As Boolean

    targetDir = ThisWorkbook.path
    If Len(targetDir) = 0 Then
        MsgBox "先にこの Excel ファイルを保存してください。", vbExclamation, "列設定の適用"
        Exit Sub
    End If

    On Error Resume Next
    Set wsRes = ThisWorkbook.Worksheets(SHEET_RESULT_TASK_LIST)
    Set wsCfg = ThisWorkbook.Worksheets(SHEET_COL_CONFIG_RESULT_TASK)
    On Error GoTo 0
    If wsRes Is Nothing Then
        MsgBox "シート「" & SHEET_RESULT_TASK_LIST & "」がありません。", vbExclamation, "列設定の適用"
        Exit Sub
    End If
    If wsCfg Is Nothing Then
        MsgBox "シート「" & SHEET_COL_CONFIG_RESULT_TASK & "」がありません。", vbExclamation, "列設定の適用"
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    Set wsh = CreateObject("WScript.Shell")
    wsh.Environment("Process")("TASK_INPUT_WORKBOOK") = ThisWorkbook.FullName

    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    MacroSplash_SetStep "列設定: Python で結果タスク一覧の列順・表示を適用しています…"
    runBat = "@echo off" & vbCrLf & "pushd """ & targetDir & """" & vbCrLf & "chcp 65001>nul" & vbCrLf & _
             "py -3 -u python\apply_result_task_column_layout.py" & vbCrLf & _
             "echo." & vbCrLf & _
             "echo [column-layout] ERRORLEVEL=%ERRORLEVEL%" & vbCrLf & _
             "exit /b %ERRORLEVEL%"
    exitCode = RunTempCmdWithConsoleLayout(wsh, runBat)
    Application.ScreenUpdating = prevScreen

    On Error Resume Next
    Set wsRes = ThisWorkbook.Worksheets(SHEET_RESULT_TASK_LIST)
    If Not wsRes Is Nothing Then
        結果シート_列幅_AutoFit非表示を維持 wsRes
    End If
    On Error GoTo 0

    If exitCode <> 0 Then
        MsgBox "Python の終了コードが " & CStr(exitCode) & " です。" & vbCrLf _
            & "log\execution_log.txt を確認してください。", vbExclamation, "列設定の適用"
    Else
        MacroSplash_SetStep "「" & SHEET_RESULT_TASK_LIST & "」の列順・表示を「" & SHEET_COL_CONFIG_RESULT_TASK & "」に合わせました。"
        m_animMacroSucceeded = True
    End If
End Sub

Public Sub 列設定_結果_タスク一覧_重複列名を整理()
    Dim wsh As Object
    Dim runBat As String
    Dim targetDir As String
    Dim exitCode As Long
    Dim wsCfg As Worksheet
    Dim prevScreen As Boolean

    targetDir = ThisWorkbook.path
    If Len(targetDir) = 0 Then
        MsgBox "先にこの Excel ファイルを保存してください。", vbExclamation, "列設定の整理"
        Exit Sub
    End If

    On Error Resume Next
    Set wsCfg = ThisWorkbook.Worksheets(SHEET_COL_CONFIG_RESULT_TASK)
    On Error GoTo 0
    If wsCfg Is Nothing Then
        MsgBox "シート「" & SHEET_COL_CONFIG_RESULT_TASK & "」がありません。", vbExclamation, "列設定の整理"
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    Set wsh = CreateObject("WScript.Shell")
    wsh.Environment("Process")("TASK_INPUT_WORKBOOK") = ThisWorkbook.FullName

    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    MacroSplash_SetStep "列設定: Python で重複列名を整理しています…"
    runBat = "@echo off" & vbCrLf & "pushd """ & targetDir & """" & vbCrLf & "chcp 65001>nul" & vbCrLf & _
             "py -3 -u python\dedupe_result_task_column_config_sheet.py" & vbCrLf & _
             "echo." & vbCrLf & _
             "echo [dedupe-column-config] ERRORLEVEL=%ERRORLEVEL%" & vbCrLf & _
             "exit /b %ERRORLEVEL%"
    exitCode = RunTempCmdWithConsoleLayout(wsh, runBat)
    Application.ScreenUpdating = prevScreen

    If exitCode <> 0 Then
        MsgBox "Python の終了コードが " & CStr(exitCode) & " です。" & vbCrLf _
            & "log\execution_log.txt を確認してください。", vbExclamation, "列設定の整理"
    Else
        MacroSplash_SetStep "「" & SHEET_COL_CONFIG_RESULT_TASK & "」の重複列名を除き A:B を更新しました。（チェックボックス利用時は配置マクロの再実行を推奨）"
        m_animMacroSucceeded = True
    End If
End Sub

Public Sub COM操作テスト_全シートをログに出す()
    Const LOG_SHEET As String = "COM操作テストログ"
    Const TEST_A99_ADDR As String = "A99"
    Const TEST_A99_TEXT As String = "A666"
    Dim wsLog As Worksheet
    Dim s As Object
    Dim ws As Worksheet
    Dim r As Long
    Dim detail As String
    Dim oldA99 As Variant
    Dim backA99 As Variant
    
    Application.ScreenUpdating = False
    On Error Resume Next
    Set wsLog = ThisWorkbook.Worksheets(LOG_SHEET)
    If Not wsLog Is Nothing Then
        Application.DisplayAlerts = False
        wsLog.Delete
        Application.DisplayAlerts = True
    End If
    On Error GoTo 0
    
    Set wsLog = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    On Error Resume Next
    wsLog.Name = LOG_SHEET
    On Error GoTo 0
    
    wsLog.Cells(1, 1).Value = "シート名"
    wsLog.Cells(1, 2).Value = "TypeName"
    wsLog.Cells(1, 3).Value = "表示状態"
    wsLog.Cells(1, 4).Value = "セル保護"
    wsLog.Cells(1, 5).Value = "読取 A1"
    wsLog.Cells(1, 6).Value = "UsedRange"
    wsLog.Cells(1, 7).Value = "ZZ1000 書込"
    wsLog.Cells(1, 8).Value = "A99へA666"
    wsLog.Cells(1, 9).Value = "Activate"
    wsLog.Cells(1, 10).Value = "メモ"
    
    r = 2
    For Each s In ThisWorkbook.Sheets
        If StrComp(s.Name, LOG_SHEET, vbBinaryCompare) = 0 Then GoTo NextSheetIter
        
        detail = ""
        wsLog.Cells(r, 1).Value = s.Name
        wsLog.Cells(r, 2).Value = TypeName(s)
        
        Select Case s.Visible
            Case xlSheetVisible
                wsLog.Cells(r, 3).Value = "表示"
            Case xlSheetHidden
                wsLog.Cells(r, 3).Value = "非表示"
            Case xlSheetVeryHidden
                wsLog.Cells(r, 3).Value = "VeryHidden"
            Case Else
                wsLog.Cells(r, 3).Value = CStr(s.Visible)
        End Select
        
        If TypeName(s) = "Worksheet" Then
            Set ws = s
            On Error Resume Next
            If ws.ProtectContents Then
                wsLog.Cells(r, 4).Value = "保護中"
            Else
                wsLog.Cells(r, 4).Value = "なし"
            End If
            If Err.Number <> 0 Then
                wsLog.Cells(r, 4).Value = "確認不可: " & Err.Description
                detail = detail & "ProtectContents " & Err.Description & "; "
            End If
            Err.Clear
            
            Dim dummy As Variant
            dummy = ws.Range("A1").Value
            If Err.Number <> 0 Then
                wsLog.Cells(r, 5).Value = "NG"
                detail = detail & "読取 " & Err.Description & "; "
            Else
                wsLog.Cells(r, 5).Value = "OK"
            End If
            Err.Clear
            
            Dim urAdr As String
            urAdr = ws.UsedRange.Address
            If Err.Number <> 0 Then
                wsLog.Cells(r, 6).Value = "NG"
                detail = detail & "UsedRange " & Err.Description & "; "
            Else
                wsLog.Cells(r, 6).Value = "OK (" & urAdr & ")"
            End If
            Err.Clear
            
            ws.Range("ZZ1000").Value = "__COM_TEST__"
            If Err.Number <> 0 Then
                wsLog.Cells(r, 7).Value = "NG"
                detail = detail & "ZZ書込 " & Err.Description & "; "
            Else
                wsLog.Cells(r, 7).Value = "OK"
                ws.Range("ZZ1000").ClearContents
            End If
            Err.Clear
            
            ' A99 に文字列 A666 を書き、読み戻して一致したら OK（元の値に復元）
            oldA99 = ws.Range(TEST_A99_ADDR).Value
            Err.Clear
            ws.Range(TEST_A99_ADDR).Value = TEST_A99_TEXT
            If Err.Number <> 0 Then
                wsLog.Cells(r, 8).Value = "NG(書込)"
                detail = detail & "A99書込Err " & Err.Description & "; "
                Err.Clear
            Else
                backA99 = ws.Range(TEST_A99_ADDR).Value
                If Err.Number <> 0 Then
                    wsLog.Cells(r, 8).Value = "NG(読取)"
                    detail = detail & "A99読取Err " & Err.Description & "; "
                    Err.Clear
                ElseIf CStr(backA99) <> TEST_A99_TEXT Then
                    wsLog.Cells(r, 8).Value = "不一致"
                    detail = detail & "A99期待=" & TEST_A99_TEXT & " 実際=" & CStr(backA99) & "; "
                Else
                    wsLog.Cells(r, 8).Value = "OK"
                End If
                ws.Range(TEST_A99_ADDR).Value = oldA99
                If Err.Number <> 0 Then
                    detail = detail & "A99復元Err " & Err.Description & "; "
                    Err.Clear
                End If
            End If
            
            ws.Activate
            If Err.Number <> 0 Then
                wsLog.Cells(r, 9).Value = "NG"
                detail = detail & "Activate " & Err.Description & "; "
            Else
                wsLog.Cells(r, 9).Value = "OK"
            End If
            Err.Clear
            On Error GoTo 0
        Else
            wsLog.Cells(r, 4).Value = "?"
            wsLog.Cells(r, 5).Value = "?"
            wsLog.Cells(r, 6).Value = "?"
            wsLog.Cells(r, 7).Value = "?"
            wsLog.Cells(r, 8).Value = "?"
            On Error Resume Next
            s.Activate
            If Err.Number <> 0 Then
                wsLog.Cells(r, 9).Value = "NG"
                detail = detail & "Activate " & Err.Description
            Else
                wsLog.Cells(r, 9).Value = "OK"
            End If
            Err.Clear
            On Error GoTo 0
            detail = detail & "（Worksheet 以外はセル系テスト対象外）"
        End If
        
        wsLog.Cells(r, 10).Value = detail
        r = r + 1
NextSheetIter:
    Next s
    
    wsLog.Columns("A:J").AutoFit
    Application.ScreenUpdating = True
    wsLog.Activate
    wsLog.Range("A1").Select
    
    MsgBox "シート「" & LOG_SHEET & "」に結果を出しました。" & vbCrLf & vbCrLf & _
        "列の意味:" & vbCrLf & _
        "・A99 列: 文字列「A666」を A99 に書き、読み戻して一致→OK、元の値に復元。" & vbCrLf & _
        "・読取/UsedRange/書込/Activate の NG は、その操作で Err が出たシートです。" & vbCrLf & _
        "・保護中で書込 NG は正常なことが多いです。" & vbCrLf & _
        "・VBA からの試験です。Python 等の別プロセス COM は環境により異なります。", _
        vbInformation, "COM 操作テスト"
End Sub

Sub Auto_Open()
    ShortcutMainSheet_OnKeyRegister
End Sub

