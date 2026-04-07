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

' ?i?K1: master.xlsm ????@?B?J?????_?[?E?????o?[?Éü???}?N???u?b?N??R?s?[?????B
' Python?i?i?K1/2?j???? master.xlsm ????????A?z???W?b?N?ÉP??R?s?[??s?v?B???? False?i?X?L?b?v?j?B
' 1 / true / yes ?c ?]???????R?s?[?i?}?N???u?b?N????}?X?^??X?i?b?v?V???b?g???????????j?B
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

' ?i?K1/2 ?p: ???? .cmd ???????? title ???????{?????iFindWindow ?????L???v?V????1?????????j
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

' ??\????: py ?s?? 1>nul 2>&1?i?W???o??E?W???G???[??????B?l???h?~?j?B?{????O?? planning_core ?? execution_log?iUserForm ??\???j
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

' xlwings RunPython: runpy.run_path ?? python\xlwings_console_runner.py ?????s
Private Sub XwRunConsoleRunner(ByVal entryPoint As String)
    On Error GoTo EH
    xlwings.RunPython "import os, runpy, xlwings as xw; wb=xw.Book.caller(); p=os.path.join(os.path.dirname(str(wb.fullname)), 'python', 'xlwings_console_runner.py'); ns=runpy.run_path(p); ns['" & entryPoint & "']()"
    Exit Sub
EH:
    Err.Raise Err.Number, "XwRunConsoleRunner", "RunPython: " & Err.Description
End Sub

' hideConsoleWindow ?p: Windows Terminal ????????w?b?h???X?N???iconhost ??????É†?? cmd?j
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

' D3=false ?I?[?o?[???C??p: ????[???? Windows Terminal ???? cmd ???N?????? FindWindow ?? CASCADIA_HOSTING ???????g????U????????Aconhost ???T?R???\?[????????
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

' D3=false: txtExecutionLog ????s?N?Z????`?i?t?H?[????N???C?A???g???_?{?|?C???g??DPI ???Z?j
#If VBA7 Then
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
#Else
Private Sub ConsoleApplyBorderlessIfNeeded(ByVal hwnd As Long)
End Sub
#End If

' ?i?K1/2: ??? D3=true ?c Exec?{??@???[?v?? execution_log ???|?[?????O?BD3=false ?c ?X?v???b?V?????? Exec?{???O?g?? cmd ?d??iuseSplashLogRectConsole?j?Bhide ???? conhost --headless
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
    ' D3=false: ?X?v???b?V???\???????? useSplashLogRectConsole ?c Exec?{title ???o????O?g?? SetWindowPos?B?????O????? Run
    If Not SettingsSheet_IsSplashExecutionLogWriteEnabled() Then
        If m_macroSplashShown And m_macroSplashLockedExcel And Not Application.Interactive Then
            Application.Interactive = True
            unlockInteractiveForPoll = True
        End If
        ' hideConsoleWindow ??????i?]????????? False ??????? STAGE12_CMD_HIDE_WINDOW ?????????????????j
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
        ' ???? Run ?? VBA ???L???X?v???b?V????X?V???????B?O????? Interactive ??????? Run ???? DoEvents ?s?\?B
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
    ' Application.Interactive=False ???? UserForm ??????`?ú{?????B?|?[?????O?????? True ?????i?I????????b?N?j
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
        ' ?????? COM?ixlwings ???j??w???????????????????Á∑???????B??20??|?[?????O?????O????i1200ms?~20?24?b?j
        If splashPollIter Mod 20 = 0 Then MacroSplash_BringFormToFront
        ' ?|?[?????O??u SPLASH_LOG_POLL_INTERVAL_MS: ?Z??????? xlwings COM ????????????i???? 1200ms?j?BFileLen ?s?????S?????X?L?b?v
        ' ?w?b?h???X????R???\?[?? HWND ???????B?]???\????????u????
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

Public Sub RunPythonStage1()
    ?i?K1_?R?A???s
    On Error Resume Next
    ?z??v??_?^?X?N????_A1??I??
    On Error GoTo 0
    If m_lastStage1ExitCode < 0 Then
        If Len(m_lastStage1ErrMsg) > 0 Then MsgBox m_lastStage1ErrMsg, vbExclamation, "?i?K1"
        Exit Sub
    End If
    If m_lastStage1ExitCode <> 0 Then
        Dim st1Block As String
        st1Block = Trim$(GeminiReadUtf8File(ThisWorkbook.path & "\log\stage2_blocking_message.txt"))
        If m_lastStage1ExitCode = 3 And Len(st1Block) > 0 Then
            MsgBox st1Block, vbCritical, "?i?K1"
        Else
            MsgBox "?i?K1?? Python ?I???R?[?h?? " & CStr(m_lastStage1ExitCode) & " ????B" & vbCrLf & "LOG ?V?[?g?Elog\execution_log.txt ???m?F????????????B?i??????~???? log\stage2_blocking_message.txt ???Q??j", vbExclamation, "?i?K1"
        End If
        Exit Sub
    End If
    MacroSplash_SetStep "?i?K1??????????????B?z??v??V?[?g???m?F??????A?K?v???i?K2?i?v?ùè???j?????s????????????B"
    m_animMacroSucceeded = True
End Sub

' ???: ?i?K1???i?K2?i??????m??X?v???b?V???{?`???C???B?G???[????? MsgBox?j
Public Sub RunPythonStage1ThenStage2()
    ?i?K1_?R?A???s
    On Error Resume Next
    ?z??v??_?^?X?N????_A1??I??
    On Error GoTo 0
    If m_lastStage1ExitCode < 0 Then
        If Len(m_lastStage1ErrMsg) > 0 Then MsgBox m_lastStage1ErrMsg, vbExclamation, "?i?K1+2"
        Exit Sub
    End If
    If m_lastStage1ExitCode <> 0 Then
        Dim st1b2 As String
        st1b2 = Trim$(GeminiReadUtf8File(ThisWorkbook.path & "\log\stage2_blocking_message.txt"))
        If m_lastStage1ExitCode = 3 And Len(st1b2) > 0 Then
            MsgBox st1b2, vbCritical, "?i?K1+2"
        Else
            MsgBox "?i?K1?? Python ?I???R?[?h?? " & CStr(m_lastStage1ExitCode) & " ?????A?i?K2????s???????B" & vbCrLf & "LOG ?V?[?g?Elog\execution_log.txt ???m?F????????????B?i??????~???? log\stage2_blocking_message.txt ???Q??j", vbExclamation, "?i?K1+2"
        End If
        Exit Sub
    End If
    ?i?K2_?R?A???s True
    If m_lastStage2ExitCode <> 0 Or Len(m_lastStage2ErrMsg) > 0 Then
        If Len(m_lastStage2ErrMsg) > 0 Then
            MsgBox m_lastStage2ErrMsg, vbCritical, "?i?K1+2"
        Else
            MsgBox "?i?K2?? Python ?I???R?[?h?? " & CStr(m_lastStage2ExitCode) & " ????BLOG ?V?[?g?Elog\execution_log.txt ???m?F????????????B", vbExclamation, "?i?K1+2"
        End If
        Exit Sub
    End If
    ?i?K2_??????????
    If m_stage2PlanImported Or m_stage2MemberImported Then m_animMacroSucceeded = True
End Sub

' ?i?K1?????O: ?????u?z??v??_?^?X?N????v??t?H???g?i???E?T?C?Y?j?????iClear ?????????j
Private Sub ?z??v??_?^?X?N????_?????V?[?g????t?H???g???èÔ( _
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
    
    ' ???s?????o????????A??????s??~????????Z????t?H???g????p
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

' ????????: xlsx ???????t?H???g?????????A?????????????
Private Sub ?z??v??_?^?X?N????_UsedRange??t?H???g????T?C?Y??K?p( _
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

Public Sub RunPython(Optional ByVal preserveStage1LogOnLogSheet As Boolean = False)
    ?i?K2_?R?A???s preserveStage1LogOnLogSheet
    If m_lastStage2ExitCode <> 0 Or Len(m_lastStage2ErrMsg) > 0 Then
        If Len(m_lastStage2ErrMsg) > 0 Then
            MsgBox m_lastStage2ErrMsg, vbCritical, "?v?ùè??"
        Else
            MsgBox "Python ??I???R?[?h?? " & CStr(m_lastStage2ExitCode) & " ????BLOG ?V?[?g?Elog\execution_log.txt ???m?F????????????B", vbExclamation, "?v?ùè??"
        End If
        Exit Sub
    End If
    ?i?K2_??????????
    If m_stage2PlanImported Or m_stage2MemberImported Then m_animMacroSucceeded = True
End Sub

' =========================================================
' ?y???z??l??X?P?W???[???p?V?[?g???i??l_?v???t?B?b?N?X?E????????????E31????????j
' =========================================================
