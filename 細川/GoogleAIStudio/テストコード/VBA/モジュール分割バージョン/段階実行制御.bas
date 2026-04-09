<<<<<<< HEAD
Private Function 段階12_CMDウィンドウ非表示_実効値() As Boolean
=======
Option Explicit

Public Function Stage12CmdHideWindowEffective() As Boolean
>>>>>>> main4
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim cellKey As String
    Dim v As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_WORKBOOK_ENV)
    On Error GoTo 0
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lastRow >= 2 Then
            For r = 2 To lastRow
                cellKey = Trim$(CStr(ws.Cells(r, 1).Value))
                If Len(cellKey) > 0 And Left$(cellKey, 1) <> "#" Then
                    If StrComp(cellKey, "STAGE12_CMD_HIDE_WINDOW", vbTextCompare) = 0 Then
                        v = Trim$(CStr(ws.Cells(r, 2).Value))
                        If Len(v) > 0 Then
                            段階12_CMDウィンドウ非表示_実効値 = 段階12_CMD非表示フラグを真偽に変換(v, STAGE12_CMD_HIDE_WINDOW)
                            Exit Function
                        End If
                        Exit For
                    End If
                End If
            Next r
        End If
    End If
    v = Trim$(Environ("STAGE12_CMD_HIDE_WINDOW"))
    If Len(v) > 0 Then
        段階12_CMDウィンドウ非表示_実効値 = 段階12_CMD非表示フラグを真偽に変換(v, STAGE12_CMD_HIDE_WINDOW)
        Exit Function
    End If
    段階12_CMDウィンドウ非表示_実効値 = STAGE12_CMD_HIDE_WINDOW
End Function

' 段階1: master.xlsm から機械カレンダー・メンバー勤怠をマクロブックへコピーするか。
' Python（段階1/2）は常に master.xlsm を直接読むため、配台ロジック上このコピーは不要。既定 False（スキップ）。
' 1 / true / yes … 従来どおりコピー（マクロブック内でマスタのスナップショットを見たい場合）。
<<<<<<< HEAD
Private Function 段階1_マスタ同期マクロブック_実効値() As Boolean
=======
Public Function Stage1SyncMasterSheetsToMacroBookEffective() As Boolean
>>>>>>> main4
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim cellKey As String
    Dim v As String
    段階1_マスタ同期マクロブック_実効値 = False
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_WORKBOOK_ENV)
    On Error GoTo 0
    If Not ws Is Nothing Then
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lastRow >= 2 Then
            For r = 2 To lastRow
                cellKey = Trim$(CStr(ws.Cells(r, 1).Value))
                If Len(cellKey) > 0 And Left$(cellKey, 1) <> "#" Then
                    If StrComp(cellKey, "STAGE1_SYNC_MASTER_SHEETS_TO_MACRO_BOOK", vbTextCompare) = 0 Then
                        v = Trim$(CStr(ws.Cells(r, 2).Value))
                        If Len(v) > 0 Then
                            段階1_マスタ同期マクロブック_実効値 = 段階12_CMD非表示フラグを真偽に変換(v, False)
                            Exit Function
                        End If
                        Exit For
                    End If
                End If
            Next r
        End If
    End If
    v = Trim$(Environ("STAGE1_SYNC_MASTER_SHEETS_TO_MACRO_BOOK"))
    If Len(v) > 0 Then
        段階1_マスタ同期マクロブック_実効値 = 段階12_CMD非表示フラグを真偽に変換(v, False)
    End If
End Function

Public Function 一時CMDファイルに書き出し(ByVal body As String) As String
    Dim p As String
    Dim fh As Integer
    Dim lines() As String
    Dim j As Long
    Dim txt As String
    txt = Replace(body, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)
    Randomize
    p = Environ("TEMP") & "\pm_ai_run_" & Format(Now, "yyyymmddhhnnss") & "_" & CStr(Int(1000000 * Rnd)) & ".cmd"
    fh = FreeFile
    Open p For Output As #fh
    For j = LBound(lines) To UBound(lines)
        Print #fh, lines(j)
    Next j
    Close #fh
    一時CMDファイルに書き出し = p
End Function

' 段階1/2 用: 同一 .cmd 内で最初に title してから本処理（FindWindow はこのキャプション1つだけになる）
Public Function CMD本文へコンソールタイトルを付与(ByVal body As String, ByVal titleText As String) As String
    Const echoOffCrLf As String = "@echo off" & vbCrLf
    If Len(body) >= Len(echoOffCrLf) And LCase$(Left$(body, 9)) = "@echo off" Then
        If Mid$(body, 10, 2) = vbCrLf Then
            CMD本文へコンソールタイトルを付与 = echoOffCrLf & "title " & titleText & vbCrLf & Mid$(body, 12)
            Exit Function
        End If
    End If
    CMD本文へコンソールタイトルを付与 = echoOffCrLf & "title " & titleText & vbCrLf & body
End Function

' 非表示時: py 行へ 1>nul 2>&1（標準出力・標準エラーを捨てる。詰まり防止）。本番ログは planning_core の execution_log（UserForm で表示）
<<<<<<< HEAD
Private Function 段階バッチ_Python行に標準出力破棄を付与(ByVal body As String) As String
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
                段階バッチ_Python行に標準出力破棄を付与 = Join(lines, vbCrLf)
                Exit Function
            End If
        End If
    Next i
    段階バッチ_Python行に標準出力破棄を付与 = body
End Function

Private Function 段階VBA終了コードファイルのパス() As String
=======
Public Function StageVbaExitCodeFilePath() As String
>>>>>>> main4
    Dim k As Long
    段階VBA終了コードファイルのパス = ""
    If Len(m_splashExecutionLogPath) > 0 Then
        k = InStrRev(m_splashExecutionLogPath, "\")
        If k <= 0 Then Exit Function
        段階VBA終了コードファイルのパス = Left$(m_splashExecutionLogPath, k) & "stage_vba_exitcode.txt"
    ElseIf Len(m_stageVbaExitCodeLogDir) > 0 Then
        段階VBA終了コードファイルのパス = m_stageVbaExitCodeLogDir & "\stage_vba_exitcode.txt"
    End If
End Function

<<<<<<< HEAD
Private Function 段階VBA終了コードをファイルから読取(ByVal fullPath As String) As Long
=======
Public Function ReadStageVbaExitCodeFromFile(ByVal fullPath As String) As Long
>>>>>>> main4
    Dim s As String
    On Error GoTo Fail
    段階VBA終了コードをファイルから読取 = &H7FFFFFFF
    If Len(Dir(fullPath)) = 0 Then Exit Function
    s = Gemini_UTF8ファイルを読込(fullPath)
    s = Trim$(Replace(Replace(Replace(s, vbCrLf, ""), vbLf, ""), vbCr, ""))
    If Len(s) = 0 Then Exit Function
    段階VBA終了コードをファイルから読取 = CLng(Val(s))
    Exit Function
Fail:
    段階VBA終了コードをファイルから読取 = &H7FFFFFFF
End Function

<<<<<<< HEAD
' xlwings ダイアログ付き_段階2を実行: runpy.run_path で python\xlwings_console_runner.py を実行
Private Sub Xlwings_コンソールランナー実行(ByVal entryPoint As String)
=======
' xlwings RunPython: runpy.run_path で python\xlwings_console_runner.py を実行
Public Sub XwRunConsoleRunner(ByVal entryPoint As String)
>>>>>>> main4
    On Error GoTo EH
    xlwings.RunPython "import os, runpy, xlwings as xw; wb=xw.Book.caller(); p=os.path.join(os.path.dirname(str(wb.fullname)), 'python', 'xlwings_console_runner.py'); ns=runpy.run_path(p); ns['" & entryPoint & "']()"
    Exit Sub
EH:
    Err.Raise Err.Number, "Xlwings_コンソールランナー実行", "ダイアログ付き_段階2を実行: " & Err.Description
End Sub

' hideConsoleWindow 用: Windows Terminal を挟まずヘッドレス起動（conhost 無ければ通常 cmd）
<<<<<<< HEAD
Private Function 段階実行_コマンドラインを構築(ByVal cmdFilePath As String, ByVal hideConsoleWindow As Boolean) As String
=======
Public Function BuildStageExecCommandLine(ByVal cmdFilePath As String, ByVal hideConsoleWindow As Boolean) As String
>>>>>>> main4
    Dim conhostExe As String
    Dim comSpec As String
    If Not hideConsoleWindow Then
        段階実行_コマンドラインを構築 = "cmd.exe /c """ & cmdFilePath & """"
        Exit Function
    End If
    conhostExe = Environ("SystemRoot") & "\System32\conhost.exe"
    comSpec = Environ("ComSpec")
    If Len(comSpec) = 0 Then comSpec = Environ("SystemRoot") & "\System32\cmd.exe"
    If Len(Dir(conhostExe)) > 0 Then
        段階実行_コマンドラインを構築 = """" & conhostExe & """ --headless """ & comSpec & """ /c """ & cmdFilePath & """"
    Else
        段階実行_コマンドラインを構築 = """" & comSpec & """ /c """ & cmdFilePath & """"
    End If
End Function

' D3=false オーバーレイ専用: 既定端末が Windows Terminal のとき cmd 直起動だと FindWindow が CASCADIA_HOSTING を返し中身が空振りしうるため、conhost で古典コンソールを強制
<<<<<<< HEAD
Private Function 段階表示用クラシックコンソールCMDを構築(ByVal cmdFilePath As String) As String
=======
Public Function BuildStageVisibleClassicConhostCmd(ByVal cmdFilePath As String) As String
>>>>>>> main4
    Dim conhostExe As String
    Dim cmdExe As String
    conhostExe = Environ("SystemRoot") & "\System32\conhost.exe"
    cmdExe = Environ("SystemRoot") & "\System32\cmd.exe"
    If Len(Dir(conhostExe)) = 0 Or Len(Dir(cmdExe)) = 0 Then
        段階表示用クラシックコンソールCMDを構築 = "cmd.exe /c """ & cmdFilePath & """"
    Else
        段階表示用クラシックコンソールCMDを構築 = """" & conhostExe & """ """ & cmdExe & """ /c """ & cmdFilePath & """"
    End If
End Function

' D3=false: txtExecutionLog の画面ピクセル矩形（フォームのクライアント原点＋ポイント→DPI 換算）
#If VBA7 Then
<<<<<<< HEAD
Private Sub コンソール枠なし化を必要なら適用(ByVal hwnd As LongPtr)
=======
Public Sub ConsoleApplyBorderlessIfNeeded(ByVal hwnd As LongPtr)
>>>>>>> main4
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
#Else
<<<<<<< HEAD
Private Sub コンソール枠なし化を必要なら適用(ByVal hwnd As Long)
=======
Public Sub ConsoleApplyBorderlessIfNeeded(ByVal hwnd As Long)
>>>>>>> main4
End Sub
#End If

' 段階1/2: 設定 D3=true … Exec＋待機ループで execution_log をポーリング。D3=false … スプラッシュ時は Exec＋ログ枠へ cmd 重ね（useSplashLogRectConsole）。hide 時は conhost --headless
Public Function CMDファイルをExecしポーリングして実行(ByVal wsh As Object, ByVal cmdFilePath As String, ByVal consoleTitle As String, ByVal applyQuarterLayout As Boolean, ByVal hideConsoleWindow As Boolean, Optional ByVal useSplashLogRectConsole As Boolean = False) As Long
    Dim execObj As Object
    Dim cmdLine As String
#If VBA7 Then
    Dim hwnd As LongPtr
#Else
    Dim hwnd As Long
#End If
    Dim positioned As Boolean
    Dim cx As Long
    Dim cyQuarter As Long
    Dim probe As Long
    Dim waitStyle As Long
    Dim exitFromFile As Long
    Dim exitPath As String
    Dim unlockInteractiveForPoll As Boolean
    Dim splashPollIter As Long
    Dim haveRect As Boolean
    Dim ox As Long
    Dim oy As Long
    Dim ow As Long
    Dim oh As Long
    cmdLine = 段階実行_コマンドラインを構築(cmdFilePath, hideConsoleWindow)
    If hideConsoleWindow Then waitStyle = 0 Else waitStyle = 1
    unlockInteractiveForPoll = False
    On Error Resume Next
    ' D3=false: スプラッシュ表示中かつ useSplashLogRectConsole … Exec＋title 検出でログ枠へ SetWindowPos。それ以外は同期 Run
    If Not 設定シート_スプラッシュログ書込み有効か() Then
        If m_macroSplashShown And m_macroSplashLockedExcel And Not Application.Interactive Then
            Application.Interactive = True
            unlockInteractiveForPoll = True
        End If
        ' hideConsoleWindow を維持（従来はここで False に上書きし STAGE12_CMD_HIDE_WINDOW が無効化されていた）
        If useSplashLogRectConsole And m_macroSplashShown And Len(consoleTitle) > 0 And Not hideConsoleWindow Then
            Dim overlayCmdLine As String
            overlayCmdLine = 段階表示用クラシックコンソールCMDを構築(cmdFilePath)
            Set execObj = wsh.Exec(overlayCmdLine)
            If Err.Number = 0 And Not execObj Is Nothing Then
                On Error GoTo 0
                If m_macroSplashShown And m_macroSplashLockedExcel And Not Application.Interactive Then
                    Application.Interactive = True
                    unlockInteractiveForPoll = True
                End If
                positioned = False
                probe = 0
                splashPollIter = 0
                haveRect = スプラッシュ_実行ログ領域の画面ピクセル矩形を取得(ox, oy, ow, oh)
                If haveRect Then
                    If ow > STAGE12_CMD_OVERLAY_RECT_INSET_PX * 2 + 80 And oh > STAGE12_CMD_OVERLAY_RECT_INSET_PX * 2 + 80 Then
                        ox = ox + STAGE12_CMD_OVERLAY_RECT_INSET_PX
                        oy = oy + STAGE12_CMD_OVERLAY_RECT_INSET_PX
                        ow = ow - 2 * STAGE12_CMD_OVERLAY_RECT_INSET_PX
                        oh = oh - 2 * STAGE12_CMD_OVERLAY_RECT_INSET_PX
                    End If
                End If
                cx = GetSystemMetrics(SM_CXSCREEN)
                cyQuarter = GetSystemMetrics(SM_CYSCREEN) \ 4
                Do While execObj.Status = 0
                    splashPollIter = splashPollIter + 1
                    If Len(consoleTitle) > 0 And Not positioned Then
                        hwnd = FindWindow(0&, consoleTitle)
                        If hwnd <> 0 Then
                            スプラッシュ_コンソールオーバーレイ開始
                            コンソール枠なし化を必要なら適用 hwnd
                            If haveRect Then
                                SetWindowPos hwnd, 0&, ox, oy, ow, oh, SWP_SHOWWINDOW Or SWP_NOACTIVATE
                            Else
                                SetWindowPos hwnd, 0&, 0&, 0&, cx, cyQuarter, SWP_NOZORDER Or SWP_SHOWWINDOW Or SWP_NOACTIVATE
                            End If
                            positioned = True
                        Else
                            probe = probe + 1
                            If probe > 200 Then positioned = True
                        End If
                    End If
                    Sleep STAGE12_CMD_OVERLAY_POLL_MS
                    DoEvents
                Loop
                スプラッシュ_コンソールオーバーレイ終了
                On Error Resume Next
                exitPath = 段階VBA終了コードファイルのパス()
                If Len(Dir(exitPath)) > 0 Then
                    exitFromFile = 段階VBA終了コードをファイルから読取(exitPath)
                    If exitFromFile <> &H7FFFFFFF Then
                        CMDファイルをExecしポーリングして実行 = exitFromFile
                    Else
                        CMDファイルをExecしポーリングして実行 = CLng(execObj.exitCode)
                    End If
                Else
                    CMDファイルをExecしポーリングして実行 = CLng(execObj.exitCode)
                End If
                If Err.Number <> 0 Then CMDファイルをExecしポーリングして実行 = -1
                On Error GoTo 0
                GoTo RestoreInteractiveAfterStagePoll
            End If
            Err.Clear
            On Error GoTo 0
        End If
        CMDファイルをExecしポーリングして実行 = wsh.Run(cmdLine, waitStyle, True)
        If Err.Number <> 0 Then CMDファイルをExecしポーリングして実行 = -1
        exitPath = 段階VBA終了コードファイルのパス()
        If Len(Dir(exitPath)) > 0 Then
            exitFromFile = 段階VBA終了コードをファイルから読取(exitPath)
            If exitFromFile <> &H7FFFFFFF Then CMDファイルをExecしポーリングして実行 = exitFromFile
        End If
        GoTo RestoreInteractiveAfterStagePoll
    End If
    Set execObj = wsh.Exec(cmdLine)
    If Err.Number <> 0 Or execObj Is Nothing Then
        Err.Clear
        ' 同期 Run は VBA を占有しスプラッシュは更新されない。念のため Interactive を戻しても Run 中は DoEvents 不能。
        If m_macroSplashShown And m_macroSplashLockedExcel And Not Application.Interactive Then
            Application.Interactive = True
            unlockInteractiveForPoll = True
        End If
        CMDファイルをExecしポーリングして実行 = wsh.Run(cmdLine, waitStyle, True)
        exitPath = 段階VBA終了コードファイルのパス()
        If Len(Dir(exitPath)) > 0 Then
            exitFromFile = 段階VBA終了コードをファイルから読取(exitPath)
            If exitFromFile <> &H7FFFFFFF Then CMDファイルをExecしポーリングして実行 = exitFromFile
        End If
        GoTo RestoreInteractiveAfterStagePoll
    End If
    On Error GoTo 0
    ' Application.Interactive=False だと UserForm がほぼ再描画されない。ポーリング中だけ True に戻す（終了後は再ロック）
    If m_macroSplashShown And m_macroSplashLockedExcel And Not Application.Interactive Then
        Application.Interactive = True
        unlockInteractiveForPoll = True
    End If
    positioned = False
    probe = 0
    splashPollIter = 0
    cx = GetSystemMetrics(SM_CXSCREEN)
    cyQuarter = GetSystemMetrics(SM_CYSCREEN) \ 4
    Do While execObj.Status = 0
        スプラッシュ_実行ログ枠を更新
        splashPollIter = splashPollIter + 1
        ' 長時間 COM（xlwings 等）で背後に回ると固まったように見えることがある。約20回ポーリングごとに前面化（1200ms×20?24秒）
        If splashPollIter Mod 20 = 0 Then スプラッシュ_フォームを最前面へ
        ' ポーリング間隔 SPLASH_LOG_POLL_INTERVAL_MS: 短すぎると xlwings COM と競合しやすい（既定 1200ms）。FileLen 不変時は全文読みスキップ
        ' ヘッドレス時はコンソール HWND が無い。従来表示時のみ位置調整
        If Not hideConsoleWindow Then
            If Len(consoleTitle) > 0 And Not positioned Then
                hwnd = FindWindow(0&, consoleTitle)
                If hwnd <> 0 Then
                    If applyQuarterLayout Then
                        SetWindowPos hwnd, 0&, 0&, 0&, cx, cyQuarter, SWP_NOZORDER Or SWP_SHOWWINDOW
                    End If
                    positioned = True
                Else
                    probe = probe + 1
                    If probe > 120 Then positioned = True
                End If
            End If
        End If
        Sleep SPLASH_LOG_POLL_INTERVAL_MS
        DoEvents
    Loop
    スプラッシュ_実行ログ枠を更新
    On Error Resume Next
    exitPath = 段階VBA終了コードファイルのパス()
    If Len(Dir(exitPath)) > 0 Then
        exitFromFile = 段階VBA終了コードをファイルから読取(exitPath)
        If exitFromFile <> &H7FFFFFFF Then
            CMDファイルをExecしポーリングして実行 = exitFromFile
        Else
            CMDファイルをExecしポーリングして実行 = CLng(execObj.exitCode)
        End If
    Else
        CMDファイルをExecしポーリングして実行 = CLng(execObj.exitCode)
    End If
    If Err.Number <> 0 Then CMDファイルをExecしポーリングして実行 = -1
    On Error GoTo 0
RestoreInteractiveAfterStagePoll:
    If unlockInteractiveForPoll Then
        If m_macroSplashLockedExcel And m_macroSplashShown Then Application.Interactive = False
    End If
End Function

Public Function CMDファイルをコンソールレイアウトで実行(ByVal wsh As Object, ByVal cmdFilePath As String) As Long
    CMDファイルをコンソールレイアウトで実行 = wsh.Run("cmd.exe /c """ & cmdFilePath & """", 1, True)
End Function

Public Sub ダイアログ付き_段階1を実行()
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
        st1Block = Trim$(Gemini_UTF8ファイルを読込(ThisWorkbook.path & "\log\stage2_blocking_message.txt"))
        If m_lastStage1ExitCode = 3 And Len(st1Block) > 0 Then
            MsgBox st1Block, vbCritical, "段階1"
        Else
            MsgBox "段階1の Python 終了コードが " & CStr(m_lastStage1ExitCode) & " です。" & vbCrLf & "LOG シート・log\execution_log.txt を確認してください。（検証中止時は log\stage2_blocking_message.txt も参照）", vbExclamation, "段階1"
        End If
        Exit Sub
    End If
    スプラッシュ_手順文を設定 "段階1が完了しました。配台計画シートを確認のうえ、必要なら段階2（計画生成）を実行してください。"
    m_animMacroSucceeded = True
End Sub

' 互換: 段階1→段階2（完了通知はスプラッシュ＋チャイム。エラー時のみ MsgBox）
Public Sub ダイアログ付き_段階1と2を連続実行()
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
        st1b2 = Trim$(Gemini_UTF8ファイルを読込(ThisWorkbook.path & "\log\stage2_blocking_message.txt"))
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

' 段階1取り込み前: 既存「配台計画_タスク入力」のフォント（名・サイズ）を退避（Clear で消えるため）
Public Sub 配台計画_タスク入力_既存シートの基準フォントを取得( _
    ByVal ws As Worksheet, _
    ByRef fontName As String, _
    ByRef fontSize As Double, _
    ByRef haveFont As Boolean)
    Dim r As Long, c As Long
    Dim ur As Range
    Dim r0 As Long, c0 As Long, rMax As Long, cMax As Long
    
    fontName = "": fontSize = 0: haveFont = False
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo 0
    If ur Is Nothing Then Exit Sub
    
    r0 = ur.Row
    c0 = ur.Column
    rMax = r0 + ur.Rows.Count - 1
    cMax = c0 + ur.Columns.Count - 1
    
    ' 先頭行を見出しとみなし、その次行以降で最初の非空セルのフォントを採用
    For r = r0 + 1 To rMax
        For c = c0 To cMax
            If Len(Trim$(CStr(ws.Cells(r, c).Value))) > 0 Then
                fontName = ws.Cells(r, c).Font.Name
                fontSize = ws.Cells(r, c).Font.Size
                If Len(fontName) > 0 And fontSize > 0 Then
                    haveFont = True
                    Exit Sub
                End If
            End If
        Next c
    Next r
    
    On Error Resume Next
    fontName = ws.Cells(r0, c0).Font.Name
    fontSize = ws.Cells(r0, c0).Font.Size
    On Error GoTo 0
    If Len(fontName) > 0 And fontSize > 0 Then haveFont = True
End Sub

' 取り込み直後: xlsx 側の既定フォントを上書きし、退避した体裁に戻す
Public Sub 配台計画_タスク入力_UsedRangeにフォント名とサイズを適用( _
    ByVal ws As Worksheet, _
    ByVal fontName As String, _
    ByVal fontSize As Double)
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    If Len(fontName) = 0 Or fontSize <= 0 Then Exit Sub
    With ws.UsedRange.Font
        .Name = fontName
        .Size = fontSize
    End With
    On Error GoTo 0
End Sub

Public Sub ダイアログ付き_段階2を実行(Optional ByVal preserveStage1LogOnLogSheet As Boolean = False)
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

' =========================================================
' 【補助】個人別スケジュール用シート名（個人_プレフィックス・禁則文字除去・31文字以内）
' =========================================================
