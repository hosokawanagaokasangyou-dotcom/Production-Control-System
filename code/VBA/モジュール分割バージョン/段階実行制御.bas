Option Explicit

' 段階1/2 cmd 非表示などと同様の真偽値解釈（設定_環境変数・OS 環境変数の B 列用）
Public Function ParseStage12CmdHideWindowBool(ByVal v As String, ByVal defaultVal As Boolean) As Boolean
    Dim t As String
    t = LCase$(Trim$(v))
    If Len(t) = 0 Then
        ParseStage12CmdHideWindowBool = defaultVal
        Exit Function
    End If
    If t = "1" Or t = "true" Or t = "yes" Or t = "on" Or t = "y" Or t = "はい" Or t = "有効" Or t = "○" Or t = "〇" Then
        ParseStage12CmdHideWindowBool = True
        Exit Function
    End If
    If t = "0" Or t = "false" Or t = "no" Or t = "off" Or t = "n" Or t = "いいえ" Or t = "無効" Or t = "×" Then
        ParseStage12CmdHideWindowBool = False
        Exit Function
    End If
    ParseStage12CmdHideWindowBool = defaultVal
End Function

' planning_core の XLWINGS_SUSPEND_AUTO_CALCULATION と同名。VBA の一括セル操作前に自動計算を手動にするか。
' シート「設定_環境変数」A 列一致かつ B 非空 → その値。未設定なら Environ。どちらも空なら既定 True（手動化する）。
Public Function XlWingsSuspendAutoCalculationEffective() As Boolean
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
                    If StrComp(cellKey, "XLWINGS_SUSPEND_AUTO_CALCULATION", vbTextCompare) = 0 Then
                        v = Trim$(CStr(ws.Cells(r, 2).Value))
                        If Len(v) > 0 Then
                            XlWingsSuspendAutoCalculationEffective = ParseStage12CmdHideWindowBool(v, True)
                            Exit Function
                        End If
                        Exit For
                    End If
                End If
            Next r
        End If
    End If
    v = Trim$(Environ("XLWINGS_SUSPEND_AUTO_CALCULATION"))
    If Len(v) > 0 Then
        XlWingsSuspendAutoCalculationEffective = ParseStage12CmdHideWindowBool(v, True)
        Exit Function
    End If
    XlWingsSuspendAutoCalculationEffective = True
End Function

Public Function Stage12CmdHideWindowEffective() As Boolean
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

' 段階1: master.xlsm から機械カレンダー・メンバー勤怠をマクロブックへコピーするか。
' Python（段階1/2）は常に master.xlsm を直接読むため、配台ロジック上このコピーは不要。既定 False（スキップ）。
' 1 / true / yes … 従来どおりコピー（マクロブック内でマスタのスナップショットを見たい場合）。
Public Function Stage1SyncMasterSheetsToMacroBookEffective() As Boolean
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

Public Function WriteTempCmdFile(ByVal body As String) As String
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

' 段階1/2 用: 同一 .cmd 内で最初に title してから本処理（FindWindow はこのキャプション1つだけになる）
Public Function AugmentCmdBodyWithConsoleTitle(ByVal body As String, ByVal titleText As String) As String
    Const echoOffCrLf As String = "@echo off" & vbCrLf
    If Len(body) >= Len(echoOffCrLf) And LCase$(Left$(body, 9)) = "@echo off" Then
        If Mid$(body, 10, 2) = vbCrLf Then
            AugmentCmdBodyWithConsoleTitle = echoOffCrLf & "title " & titleText & vbCrLf & Mid$(body, 12)
            Exit Function
        End If
    End If
    AugmentCmdBodyWithConsoleTitle = echoOffCrLf & "title " & titleText & vbCrLf & body
End Function

' 非表示時: py 行へ 1>nul 2>&1（標準出力・標準エラーを捨てる。詰まり防止）。本番ログは planning_core の execution_log（UserForm で表示）
Public Function StageVbaExitCodeFilePath() As String
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

Public Function ReadStageVbaExitCodeFromFile(ByVal fullPath As String) As Long
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

' xlwings RunPython: runpy.run_path で python\xlwings_console_runner.py を実行
Public Sub XwRunConsoleRunner(ByVal entryPoint As String)
    On Error GoTo EH
    xlwings.RunPython "import os, runpy, xlwings as xw; wb=xw.Book.caller(); p=os.path.join(os.path.dirname(str(wb.fullname)), 'python', 'xlwings_console_runner.py'); ns=runpy.run_path(p); ns['" & entryPoint & "']()"
    Exit Sub
EH:
    Err.Raise Err.Number, "XwRunConsoleRunner", "RunPython: " & Err.Description
End Sub

' hideConsoleWindow 用: Windows Terminal を挟まずヘッドレス起動（conhost 無ければ通常 cmd）
Public Function BuildStageExecCommandLine(ByVal cmdFilePath As String, ByVal hideConsoleWindow As Boolean) As String
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

' D3=false オーバーレイ専用: 既定端末が Windows Terminal のとき cmd 直起動だと FindWindow が CASCADIA_HOSTING を返し中身が空振りしうるため、conhost で古典コンソールを強制
Public Function BuildStageVisibleClassicConhostCmd(ByVal cmdFilePath As String) As String
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

' D3=false: txtExecutionLog の画面ピクセル矩形（フォームのクライアント原点＋ポイント→DPI 換算）
#If VBA7 Then
Public Sub ConsoleApplyBorderlessIfNeeded(ByVal hwnd As LongPtr)
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
Public Sub ConsoleApplyBorderlessIfNeeded(ByVal hwnd As Long)
End Sub
#End If

' 段階1/2: 設定 D3=true … Exec＋待機ループで execution_log をポーリング。D3=false … スプラッシュ時は Exec＋ログ枠へ cmd 重ね（useSplashLogRectConsole）。hide 時は conhost --headless
Public Function RunCmdFileStageExecAndPoll(ByVal wsh As Object, ByVal cmdFilePath As String, ByVal consoleTitle As String, ByVal applyQuarterLayout As Boolean, ByVal hideConsoleWindow As Boolean, Optional ByVal useSplashLogRectConsole As Boolean = False) As Long
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
        ' 長時間 COM（xlwings 等）で背後に回ると固まったように見えることがある。約20回ポーリングごとに前面化（1000ms×20?20秒）
        If splashPollIter Mod 20 = 0 Then MacroSplash_BringFormToFront
        ' ポーリング間隔 SPLASH_LOG_POLL_INTERVAL_MS: 短すぎると xlwings COM と競合しやすい（既定 1000ms）。FileLen 不変時は全文読みスキップ
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

Public Function RunCmdFileWithConsoleLayout(ByVal wsh As Object, ByVal cmdFilePath As String) As Long
    RunCmdFileWithConsoleLayout = wsh.Run("cmd.exe /c """ & cmdFilePath & """", 1, True)
End Function

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
    
    ' 段階1: Python 実行後の後処理（メイン反映）を実行するか
    Dim skipPost As Boolean
    Dim t0 As Double
    skipPost = Stage1SkipMainSheetPostProcessEffective()
    If skipPost Then
        Stage1AppendExecutionLogLine "INFO", "段階1: 後処理（シートの表示/非表示・並べ替え等）はスキップします（手動実行）。"
    Else
        t0 = Timer
        Stage1AppendExecutionLogLine "INFO", "段階1: 後処理（シートの表示/非表示・並べ替え等）を開始。"
        On Error Resume Next
        メインシート_段階1実行後_リンク更新
        On Error GoTo 0
        Stage1AppendExecutionLogLine "INFO", "段階1: 後処理（シートの表示/非表示・並べ替え等）が完了。sec=" & Format$(Timer - t0, "0.000")
    End If
    MacroSplash_SetStep "段階1が完了しました。配台計画シートを確認のうえ、必要なら段階2（計画生成）を実行してください。"
    ' メイン反映の直後はメインがアクティブになるため、タスク抽出完了時は配台計画シートへ戻す
    On Error Resume Next
    配台計画_タスク入力_A1を選択
    On Error GoTo 0
    m_animMacroSucceeded = True
End Sub

' 段階1: 終了後の後処理（シートの表示/非表示・並べ替え等）を省略するか。
' 既定は True（ユーザー要望: 手動で実行するため段階1では省略）。
'
' 環境変数:
' - STAGE1_SKIP_MAIN_POST_PROCESS: 1/true/yes/on でスキップ、0/false/no/off で実行
' - 互換: STAGE1_SKIP_MAIN_LINK_UPDATE（旧名）
Public Function Stage1SkipMainSheetPostProcessEffective() As Boolean
    Dim v As String
    v = Trim$(Environ$("STAGE1_SKIP_MAIN_POST_PROCESS"))
    If Len(v) = 0 Then v = Trim$(Environ$("STAGE1_SKIP_MAIN_LINK_UPDATE"))
    If Len(v) > 0 Then
        v = LCase$(v)
        Stage1SkipMainSheetPostProcessEffective = (v = "1" Or v = "true" Or v = "yes" Or v = "on")
        Exit Function
    End If
    Stage1SkipMainSheetPostProcessEffective = True
End Function

' 段階1: log\execution_log.txt へ追記（失敗しても落とさない）
Public Sub Stage1AppendExecutionLogLine(ByVal level As String, ByVal msg As String)
    On Error Resume Next
    Dim p As String
    Dim fh As Integer
    Dim ts As String
    Dim line As String
    p = ThisWorkbook.path & "\log\execution_log.txt"
    ts = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    line = ts & " - " & level & " - " & msg & vbCrLf
    fh = FreeFile
    Open p For Append As #fh
    Print #fh, Replace(Replace(line, vbCrLf, vbLf), vbLf, vbCrLf);
    Close #fh
    On Error GoTo 0
End Sub

' 互換: 段階1→段階2（完了通知はスプラッシュ＋チャイム。エラー時のみ MsgBox）
Public Sub RunPythonStage1ThenStage2()
    Dim t0 As Double
    t0 = Timer
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
    t0 = Timer
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

Public Sub RunPython(Optional ByVal preserveStage1LogOnLogSheet As Boolean = False)
    Dim t0 As Double
    t0 = Timer
    段階2_コア実行 preserveStage1LogOnLogSheet
    If m_lastStage2ExitCode <> 0 Or Len(m_lastStage2ErrMsg) > 0 Then
        If Len(m_lastStage2ErrMsg) > 0 Then
            MsgBox m_lastStage2ErrMsg, vbCritical, "計画生成"
        Else
            MsgBox "Python の終了コードが " & CStr(m_lastStage2ExitCode) & " です。LOG シート・log\execution_log.txt を確認してください。", vbExclamation, "計画生成"
        End If
        Exit Sub
    End If
    t0 = Timer
    段階2_取り込み結果を報告
    If m_stage2PlanImported Or m_stage2MemberImported Then m_animMacroSucceeded = True
End Sub

' =========================================================
' 【補助】個人別スケジュール用シート名（個人_プレフィックス・禁則文字除去・31文字以内）
' =========================================================

 
