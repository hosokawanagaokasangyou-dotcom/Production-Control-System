Attribute VB_Name = "段階実行制御"
Option Explicit

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

Private Function Stage12CmdHideWindowEffective() As Boolean
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
                            Stage12CmdHideWindowEffective = ParseStage12CmdHideWindowBool(v, STAGE12_CMD_HIDE_WINDOW)
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
        Stage12CmdHideWindowEffective = ParseStage12CmdHideWindowBool(v, STAGE12_CMD_HIDE_WINDOW)
        Exit Function
    End If
    Stage12CmdHideWindowEffective = STAGE12_CMD_HIDE_WINDOW
End Function

Private Function Stage1SyncMasterSheetsToMacroBookEffective() As Boolean
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long
    Dim cellKey As String
    Dim v As String
    Stage1SyncMasterSheetsToMacroBookEffective = False
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
                            Stage1SyncMasterSheetsToMacroBookEffective = ParseStage12CmdHideWindowBool(v, False)
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
        Stage1SyncMasterSheetsToMacroBookEffective = ParseStage12CmdHideWindowBool(v, False)
    End If
End Function

Private Function WriteTempCmdFile(ByVal body As String) As String
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
    WriteTempCmdFile = p
End Function

Private Function StageVbaExitCodeFilePath() As String
    Dim k As Long
    StageVbaExitCodeFilePath = ""
    If Len(m_splashExecutionLogPath) > 0 Then
        k = InStrRev(m_splashExecutionLogPath, "\")
        If k <= 0 Then Exit Function
        StageVbaExitCodeFilePath = Left$(m_splashExecutionLogPath, k) & "stage_vba_exitcode.txt"
    ElseIf Len(m_stageVbaExitCodeLogDir) > 0 Then
        StageVbaExitCodeFilePath = m_stageVbaExitCodeLogDir & "\stage_vba_exitcode.txt"
    End If
End Function

Private Function ReadStageVbaExitCodeFromFile(ByVal fullPath As String) As Long
    Dim s As String
    On Error GoTo Fail
    ReadStageVbaExitCodeFromFile = &H7FFFFFFF
    If Len(Dir(fullPath)) = 0 Then Exit Function
    s = GeminiReadUtf8File(fullPath)
    s = Trim$(Replace(Replace(Replace(s, vbCrLf, ""), vbLf, ""), vbCr, ""))
    If Len(s) = 0 Then Exit Function
    ReadStageVbaExitCodeFromFile = CLng(Val(s))
    Exit Function
Fail:
    ReadStageVbaExitCodeFromFile = &H7FFFFFFF
End Function

Private Sub XwRunConsoleRunner(ByVal entryPoint As String)
    On Error GoTo EH
    xlwings.RunPython "import os, runpy, xlwings as xw; wb=xw.Book.caller(); p=os.path.join(os.path.dirname(str(wb.fullname)), 'python', 'xlwings_console_runner.py'); ns=runpy.run_path(p); ns['" & entryPoint & "']()"
    Exit Sub
EH:
    Err.Raise Err.Number, "XwRunConsoleRunner", "RunPython: " & Err.Description
End Sub

Private Function BuildStageExecCommandLine(ByVal cmdFilePath As String, ByVal hideConsoleWindow As Boolean) As String
    Dim conhostExe As String
    Dim comSpec As String
    If Not hideConsoleWindow Then
        BuildStageExecCommandLine = "cmd.exe /c """ & cmdFilePath & """"
        Exit Function
    End If
    conhostExe = Environ("SystemRoot") & "\System32\conhost.exe"
    comSpec = Environ("ComSpec")
    If Len(comSpec) = 0 Then comSpec = Environ("SystemRoot") & "\System32\cmd.exe"
    If Len(Dir(conhostExe)) > 0 Then
        BuildStageExecCommandLine = """" & conhostExe & """ --headless """ & comSpec & """ /c """ & cmdFilePath & """"
    Else
        BuildStageExecCommandLine = """" & comSpec & """ /c """ & cmdFilePath & """"
    End If
End Function

Private Function BuildStageVisibleClassicConhostCmd(ByVal cmdFilePath As String) As String
    Dim conhostExe As String
    Dim cmdExe As String
    conhostExe = Environ("SystemRoot") & "\System32\conhost.exe"
    cmdExe = Environ("SystemRoot") & "\System32\cmd.exe"
    If Len(Dir(conhostExe)) = 0 Or Len(Dir(cmdExe)) = 0 Then
        BuildStageVisibleClassicConhostCmd = "cmd.exe /c """ & cmdFilePath & """"
    Else
        BuildStageVisibleClassicConhostCmd = """" & conhostExe & """ """ & cmdExe & """ /c """ & cmdFilePath & """"
    End If
End Function

Private Function RunCmdFileStageExecAndPoll(ByVal wsh As Object, ByVal cmdFilePath As String, ByVal consoleTitle As String, ByVal applyQuarterLayout As Boolean, ByVal hideConsoleWindow As Boolean, Optional ByVal useSplashLogRectConsole As Boolean = False) As Long
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
    cmdLine = BuildStageExecCommandLine(cmdFilePath, hideConsoleWindow)
    If hideConsoleWindow Then waitStyle = 0 Else waitStyle = 1
    unlockInteractiveForPoll = False
    On Error Resume Next
    ' D3=false: スプラッシュ表示中かつ useSplashLogRectConsole … Exec＋title 検出でログ枠へ SetWindowPos。それ以外は同期 Run
    If Not SettingsSheet_IsSplashExecutionLogWriteEnabled() Then
        If m_macroSplashShown And m_macroSplashLockedExcel And Not Application.Interactive Then
            Application.Interactive = True
            unlockInteractiveForPoll = True
        End If
        ' hideConsoleWindow を維持（従来はここで False に上書きし STAGE12_CMD_HIDE_WINDOW が無効化されていた）
        If useSplashLogRectConsole And m_macroSplashShown And Len(consoleTitle) > 0 And Not hideConsoleWindow Then
            Dim overlayCmdLine As String
            overlayCmdLine = BuildStageVisibleClassicConhostCmd(cmdFilePath)
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
                haveRect = MacroSplash_GetTxtExecutionLogScreenRectPixels(ox, oy, ow, oh)
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
                            MacroSplash_BeginConsoleOverlay
                            ConsoleApplyBorderlessIfNeeded hwnd
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
                MacroSplash_EndConsoleOverlay
                On Error Resume Next
                exitPath = StageVbaExitCodeFilePath()
                If Len(Dir(exitPath)) > 0 Then
                    exitFromFile = ReadStageVbaExitCodeFromFile(exitPath)
                    If exitFromFile <> &H7FFFFFFF Then
                        RunCmdFileStageExecAndPoll = exitFromFile
                    Else
                        RunCmdFileStageExecAndPoll = CLng(execObj.exitCode)
                    End If
                Else
                    RunCmdFileStageExecAndPoll = CLng(execObj.exitCode)
                End If
                If Err.Number <> 0 Then RunCmdFileStageExecAndPoll = -1
                On Error GoTo 0
                GoTo RestoreInteractiveAfterStagePoll
            End If
            Err.Clear
            On Error GoTo 0
        End If
        RunCmdFileStageExecAndPoll = wsh.Run(cmdLine, waitStyle, True)
        If Err.Number <> 0 Then RunCmdFileStageExecAndPoll = -1
        exitPath = StageVbaExitCodeFilePath()
        If Len(Dir(exitPath)) > 0 Then
            exitFromFile = ReadStageVbaExitCodeFromFile(exitPath)
            If exitFromFile <> &H7FFFFFFF Then RunCmdFileStageExecAndPoll = exitFromFile
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
        RunCmdFileStageExecAndPoll = wsh.Run(cmdLine, waitStyle, True)
        exitPath = StageVbaExitCodeFilePath()
        If Len(Dir(exitPath)) > 0 Then
            exitFromFile = ReadStageVbaExitCodeFromFile(exitPath)
            If exitFromFile <> &H7FFFFFFF Then RunCmdFileStageExecAndPoll = exitFromFile
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
        MacroSplash_RefreshExecutionLogPane
        splashPollIter = splashPollIter + 1
        ' 長時間 COM（xlwings 等）で背後に回ると固まったように見えることがある。約20回ポーリングごとに前面化（1200ms×20?24秒）
        If splashPollIter Mod 20 = 0 Then MacroSplash_BringFormToFront
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
    MacroSplash_RefreshExecutionLogPane
    On Error Resume Next
    exitPath = StageVbaExitCodeFilePath()
    If Len(Dir(exitPath)) > 0 Then
        exitFromFile = ReadStageVbaExitCodeFromFile(exitPath)
        If exitFromFile <> &H7FFFFFFF Then
            RunCmdFileStageExecAndPoll = exitFromFile
        Else
            RunCmdFileStageExecAndPoll = CLng(execObj.exitCode)
        End If
    Else
        RunCmdFileStageExecAndPoll = CLng(execObj.exitCode)
    End If
    If Err.Number <> 0 Then RunCmdFileStageExecAndPoll = -1
    On Error GoTo 0
RestoreInteractiveAfterStagePoll:
    If unlockInteractiveForPoll Then
        If m_macroSplashLockedExcel And m_macroSplashShown Then Application.Interactive = False
    End If
End Function

Private Function RunCmdFileWithConsoleLayout(ByVal wsh As Object, ByVal cmdFilePath As String) As Long
    RunCmdFileWithConsoleLayout = wsh.Run("cmd.exe /c """ & cmdFilePath & """", 1, True)
End Function

Private Function RunTempCmdWithConsoleLayout(ByVal wsh As Object, ByVal body As String, Optional ByVal applyTopQuarterFullWidthConsole As Boolean = False, Optional ByVal hideCmdWindow As Boolean = False) As Long
    Dim p As String
    Dim uniq As String
    Dim batText As String
    ' D3=false: STAGE12_D3FALSE_SPLASH_CONSOLE_LAYOUT かつスプラッシュ時のみオーバーレイ用 Exec。それ以外は同期 Run（ウィンドウレイアウトは OS 任せ）
    If Not SettingsSheet_IsSplashExecutionLogWriteEnabled() Then
        ' ログ枠オーバーレイは「見えるコンソール」前提。非表示指定時は D3=true 経路と同様に headless へ。
        If m_macroSplashShown And STAGE12_D3FALSE_SPLASH_CONSOLE_LAYOUT And Not hideCmdWindow Then
            Randomize
            uniq = "PM_AI_CMD_" & Format(Now, "yyyymmddhhnnss") & "_" & CStr(Int(1000000 * Rnd))
            batText = AugmentCmdBodyWithConsoleTitle(body, uniq)
            p = WriteTempCmdFile(batText)
            RunTempCmdWithConsoleLayout = RunCmdFileStageExecAndPoll(wsh, p, uniq, False, False, True)
        ElseIf hideCmdWindow Or applyTopQuarterFullWidthConsole Then
            batText = body
            If hideCmdWindow Then batText = EnsureStageBatchStdoutRedirect(batText)
            Randomize
            uniq = "PM_AI_" & Format(Now, "yyyymmddhhnnss") & "_" & CStr(Int(1000000 * Rnd))
            batText = AugmentCmdBodyWithConsoleTitle(batText, uniq)
            p = WriteTempCmdFile(batText)
            RunTempCmdWithConsoleLayout = RunCmdFileStageExecAndPoll(wsh, p, uniq, applyTopQuarterFullWidthConsole And Not hideCmdWindow, hideCmdWindow, False)
        Else
            p = WriteTempCmdFile(body)
            RunTempCmdWithConsoleLayout = RunCmdFileStageExecAndPoll(wsh, p, "", False, False, False)
        End If
        GoTo RunTempCmdWithConsoleLayoutCleanup
    End If
    If hideCmdWindow Or applyTopQuarterFullWidthConsole Then
        batText = body
        If hideCmdWindow Then batText = EnsureStageBatchStdoutRedirect(batText)
        Randomize
        uniq = "PM_AI_" & Format(Now, "yyyymmddhhnnss") & "_" & CStr(Int(1000000 * Rnd))
        batText = AugmentCmdBodyWithConsoleTitle(batText, uniq)
        p = WriteTempCmdFile(batText)
        RunTempCmdWithConsoleLayout = RunCmdFileStageExecAndPoll(wsh, p, uniq, applyTopQuarterFullWidthConsole And Not hideCmdWindow, hideCmdWindow)
    Else
        p = WriteTempCmdFile(body)
        RunTempCmdWithConsoleLayout = RunCmdFileWithConsoleLayout(wsh, p)
    End If
RunTempCmdWithConsoleLayoutCleanup:
    On Error Resume Next
    Kill p
    On Error GoTo 0
End Function

