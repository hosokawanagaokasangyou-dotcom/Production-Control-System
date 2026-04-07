Private Function MacroSplash_GetTxtExecutionLogScreenRectPixels(ByRef outL As Long, ByRef outT As Long, ByRef outW As Long, ByRef outH As Long) As Boolean
#If VBA7 Then
    Dim hwndSplash As LongPtr
    Dim hdc As LongPtr
#Else
    Dim hwndSplash As Long
    Dim hdc As Long
#End If
    Dim pt As POINTAPI
    Dim dpiX As Long
    Dim dpiY As Long
    Dim tb As Object
    On Error GoTo Fail
    MacroSplash_GetTxtExecutionLogScreenRectPixels = False
    If Not m_macroSplashShown Then Exit Function
    hwndSplash = FindWindow(0&, SPLASH_FORM_WINDOW_TITLE)
    If hwndSplash = 0 Then Exit Function
    Set tb = frmMacroSplash.Controls("txtExecutionLog")
    If tb Is Nothing Then Exit Function
    dpiX = 96
    dpiY = 96
    hdc = GetDC(hwndSplash)
    If hdc <> 0 Then
        dpiX = GetDeviceCaps(hdc, LOGPIXELSX)
        dpiY = GetDeviceCaps(hdc, LOGPIXELSY)
        Call ReleaseDC(hwndSplash, hdc)
    End If
    pt.x = 0
    pt.y = 0
    Call ClientToScreen(hwndSplash, pt)
    outL = pt.x + CLng((CDbl(tb.Left) * dpiX) / 72#)
    outT = pt.y + CLng((CDbl(tb.Top) * dpiY) / 72#)
    outW = CLng((CDbl(tb.Width) * dpiX) / 72#)
    outH = CLng((CDbl(tb.Height) * dpiY) / 72#)
    If outW < 20 Or outH < 20 Then Exit Function
    MacroSplash_GetTxtExecutionLogScreenRectPixels = True
    Exit Function
Fail:
    MacroSplash_GetTxtExecutionLogScreenRectPixels = False
End Function

' cmd ?????O?g??d??????A??d?\??????????? TextBox ??????\??
Private Sub MacroSplash_BeginConsoleOverlay()
    On Error Resume Next
    If Not m_macroSplashShown Then Exit Sub
    If m_splashConsoleOverlayActive Then Exit Sub
    frmMacroSplash.Controls("txtExecutionLog").Visible = False
    m_splashConsoleOverlayActive = True
    frmMacroSplash.Repaint
    On Error GoTo 0
End Sub

Private Sub MacroSplash_EndConsoleOverlay()
    On Error Resume Next
    If Not m_splashConsoleOverlayActive Then Exit Sub
    m_splashConsoleOverlayActive = False
    If m_macroSplashShown Then
        frmMacroSplash.Controls("txtExecutionLog").Visible = True
        frmMacroSplash.Repaint
    End If
    On Error GoTo 0
End Sub

' ?R???\?[???g????????iWin32/Win64?BWT ?z?X?g HWND ????????????? ?\ ?I?[?o?[???C?? conhost ????????p?j
#If VBA7 Then
Public Sub MacroSplash_SetStep(ByVal stepMessage As String)
    Dim prevSU As Boolean
    On Error Resume Next
    If Not m_macroSplashShown Then Exit Sub
    prevSU = Application.ScreenUpdating
    If Not prevSU Then Application.ScreenUpdating = True
    frmMacroSplash.lblMessage.Caption = stepMessage
    frmMacroSplash.Repaint
    DoEvents
    If Not prevSU Then Application.ScreenUpdating = False
End Sub

Private Sub MacroSplash_ClearExecutionLogPane()
    Dim tb As Object
    On Error Resume Next
    m_splashReadErrShown = False
    m_splashLastLogSnapshot = ""
    m_splashPollHaveCachedFileLen = False
    m_splashPollLastFileLen = 0
    If Not m_macroSplashShown Then Exit Sub
    If Not SettingsSheet_IsSplashExecutionLogWriteEnabled() Then Exit Sub
    Set tb = frmMacroSplash.Controls("txtExecutionLog")
    If Not tb Is Nothing Then tb.text = ""
End Sub

' ???O?????????V?B?L?????b?g??????u?? txtExecutionLog ??t?H?[?J?X?iUserForm ??? SetFocus ??????j
Private Sub MacroSplash_TextBoxScrollToTail(ByVal tb As Object)
    On Error Resume Next
    tb.HideSelection = False
    tb.SelStart = Len(tb.text)
    tb.SelLength = 0
    If Application.Interactive Then
        tb.SetFocus
    End If
    frmMacroSplash.Repaint
    DoEvents
End Sub

' m_splashExecutionLogPath ?? UTF-8 ???O?? txtExecutionLog ??i???????????????j
Private Sub MacroSplash_RefreshExecutionLogPane()
    Dim tb As Object
    Dim s As String
    Dim n As Long
    Dim prevSU As Boolean
    Dim errBanner As String
    Dim flen As Long
    Dim flenAtStart As Long
    On Error Resume Next
    If Not SettingsSheet_IsSplashExecutionLogWriteEnabled() Then Exit Sub
    If Not m_macroSplashShown Then Exit Sub
    If Len(m_splashExecutionLogPath) = 0 Then Exit Sub
    Set tb = frmMacroSplash.Controls("txtExecutionLog")
    If tb Is Nothing Then Exit Sub
    flenAtStart = -1
    If Len(Dir(m_splashExecutionLogPath)) > 0 Then
        flenAtStart = FileLen(m_splashExecutionLogPath)
        If Err.Number <> 0 Then Err.Clear: flenAtStart = -1
    End If
    If m_splashPollHaveCachedFileLen And flenAtStart >= 0 And flenAtStart = m_splashPollLastFileLen And Not m_splashReadErrShown Then
        Exit Sub
    End If
    s = GeminiReadUtf8File(m_splashExecutionLogPath)
    If Len(s) = 0 And Len(Dir(m_splashExecutionLogPath)) > 0 Then
        s = GeminiReadUtf8FileViaTempCopy(m_splashExecutionLogPath)
    End If
    n = Len(s)
    If n > 0 Then
        m_splashReadErrShown = False
        If n > SPLASH_LOG_MAX_DISPLAY_CHARS Then
            s = "?c?i?`????????B??????\???j?c" & vbCrLf & Right$(s, SPLASH_LOG_MAX_DISPLAY_CHARS)
        End If
        If StrComp(s, m_splashLastLogSnapshot, vbBinaryCompare) = 0 Then
            m_splashPollLastFileLen = flenAtStart
            m_splashPollHaveCachedFileLen = True
            Exit Sub
        End If
        m_splashLastLogSnapshot = s
        m_splashPollLastFileLen = flenAtStart
        m_splashPollHaveCachedFileLen = True
        prevSU = Application.ScreenUpdating
        If Not prevSU Then Application.ScreenUpdating = True
        tb.text = s
        MacroSplash_TextBoxScrollToTail tb
        If Not prevSU Then Application.ScreenUpdating = False
        Exit Sub
    End If
    If Len(Dir(m_splashExecutionLogPath)) = 0 Then Exit Sub
    flen = 0
    flen = FileLen(m_splashExecutionLogPath)
    If Err.Number <> 0 Then Err.Clear
    If flen = 0 And Len(tb.text) = 0 Then Exit Sub
    If m_splashReadErrShown Then Exit Sub
    m_splashPollHaveCachedFileLen = False
    errBanner = "?y???O?\???G???[?zexecution_log.txt ?? VBA ???????????????iPython ???t?@?C?????J????????j?BLOG ?V?[?g?????G?f?B?^?? " & m_splashExecutionLogPath & " ???J????m?F????????????B" & vbCrLf & vbCrLf
    prevSU = Application.ScreenUpdating
    If Not prevSU Then Application.ScreenUpdating = True
    tb.text = errBanner & tb.text
    m_splashLastLogSnapshot = tb.text
    tb.SelStart = 1
    tb.SelLength = 0
    frmMacroSplash.lblMessage.Caption = "?c?i???s???O??\??????s ? ???L??y???O?\???G???[?z???Q??j"
    frmMacroSplash.Repaint
    DoEvents
    If Not prevSU Then Application.ScreenUpdating = False
    m_splashReadErrShown = True
End Sub

' RunPython ?I????????A???? Python ??|?[?????O??????????????? execution_log ?????\???iInteractive ?? True ??`??j
Private Sub MacroSplash_LoadExecutionLogFromPath(ByVal fullPath As String)
    Dim tb As Object
    Dim s As String
    Dim n As Long
    Dim prevInt As Boolean
    Dim errBanner As String
    Dim flen As Long
    On Error Resume Next
    If Not SettingsSheet_IsSplashExecutionLogWriteEnabled() Then Exit Sub
    If Not m_macroSplashShown Then Exit Sub
    If Len(Dir(fullPath)) = 0 Then Exit Sub
    Set tb = frmMacroSplash.Controls("txtExecutionLog")
    If tb Is Nothing Then Exit Sub
    s = GeminiReadUtf8File(fullPath)
    If Len(s) = 0 Then s = GeminiReadUtf8FileViaTempCopy(fullPath)
    n = Len(s)
    If n = 0 Then
        flen = FileLen(fullPath)
        If Err.Number <> 0 Then Err.Clear
        If flen > 0 Or Len(tb.text) > 0 Then
            errBanner = "?y???O?\???G???[?zexecution_log.txt ????????????????B?p?X: " & fullPath & vbCrLf & vbCrLf
            prevInt = Application.Interactive
            Application.Interactive = True
            tb.text = errBanner & tb.text
            m_splashLastLogSnapshot = tb.text
            frmMacroSplash.lblMessage.Caption = "?c?i???s???O????\??????s ? ???L???Q??j"
            frmMacroSplash.Repaint
            DoEvents
            If m_macroSplashLockedExcel Then Application.Interactive = False Else Application.Interactive = prevInt
        End If
        Exit Sub
    End If
    m_splashReadErrShown = False
    If n > SPLASH_LOG_MAX_DISPLAY_CHARS Then
        s = "?c?i?`????????B??????\???j?c" & vbCrLf & Right$(s, SPLASH_LOG_MAX_DISPLAY_CHARS)
    End If
    prevInt = Application.Interactive
    Application.Interactive = True
    tb.text = s
    m_splashLastLogSnapshot = s
    MacroSplash_TextBoxScrollToTail tb
    If m_macroSplashLockedExcel Then
        Application.Interactive = False
    Else
        Application.Interactive = prevInt
    End If
End Sub

Private Sub MacroSplash_PositionDockExcelBottomCenter()
#If VBA7 Then
    Dim xlHwnd As LongPtr
    Dim splashHwnd As LongPtr
#Else
    Dim xlHwnd As Long
    Dim splashHwnd As Long
#End If
    Dim rcX As RECT
    Dim rcS As RECT
    Dim xlW As Long
    Dim sw As Long
    Dim sh As Long
    Dim newL As Long
    Dim newT As Long
    On Error GoTo SplashDockDone
    If Not m_macroSplashShown Then GoTo SplashDockDone
#If VBA7 Then
    xlHwnd = Application.hwnd
#Else
    xlHwnd = Application.hwnd
#End If
    If xlHwnd = 0 Then GoTo SplashDockDone
    splashHwnd = FindWindow(0&, SPLASH_FORM_WINDOW_TITLE)
    If splashHwnd = 0 Then GoTo SplashDockDone
    If GetWindowRect(xlHwnd, rcX) = 0 Then GoTo SplashDockDone
    If GetWindowRect(splashHwnd, rcS) = 0 Then GoTo SplashDockDone
    xlW = rcX.Right - rcX.Left
    sw = rcS.Right - rcS.Left
    sh = rcS.Bottom - rcS.Top
    If xlW < 80 Or sw < 40 Or sh < 40 Then GoTo SplashDockDone
    newL = rcX.Left + (xlW - sw) \ 2
    newT = rcX.Bottom - sh - SPLASH_EXCEL_BOTTOM_GAP_PX
    Call SetWindowPos(splashHwnd, 0&, newL, newT, sw, sh, SWP_NOZORDER Or SWP_SHOWWINDOW)
SplashDockDone:
    On Error GoTo 0
End Sub

' ???[?h???X UserForm ???w???c????`??E???O?X?V???~?????????????????????BShow ?????O???i???[?U?[???N???b?N????????????|?j?B
Private Sub MacroSplash_BringFormToFront()
#If VBA7 Then
    Dim hwnd As LongPtr
#Else
    Dim hwnd As Long
#End If
    On Error Resume Next
    If Not m_macroSplashShown Then Exit Sub
    MacroSplash_PositionDockExcelBottomCenter
    hwnd = FindWindow(0&, SPLASH_FORM_WINDOW_TITLE)
    If hwnd = 0 Then Exit Sub
    BringWindowToTop hwnd
    SetForegroundWindow hwnd
    On Error GoTo 0
End Sub

Private Sub MacroSplash_Show(Optional ByVal message As String, Optional ByVal lockExcelUI As Boolean = True)
    On Error GoTo CleanupFail
    If m_macroSplashShown Then MacroSplash_Hide
    m_animMacroSucceeded = False
    If Len(Trim$(message)) = 0 Then
        message = "??????????B?????Y??????????????B"
    End If
    frmMacroSplash.Caption = SPLASH_FORM_WINDOW_TITLE
    frmMacroSplash.lblMessage.Caption = message
    frmMacroSplash.StartUpPosition = 2  ' ???????B????? MacroSplash_PositionDockExcelBottomCenter ?? Excel ???[??????
    m_macroSplashLockedExcel = False
    If lockExcelUI Then
        Application.Interactive = False
        m_macroSplashLockedExcel = True
    End If
    frmMacroSplash.Show vbModeless
    m_macroSplashShown = True
    On Error Resume Next
    frmMacroSplash.Controls("txtExecutionLog").HideSelection = False
    MacroSplash_BringFormToFront
    DoEvents
    MacroStartBgm_StartIfAvailable
    Exit Sub
CleanupFail:
    On Error Resume Next
    If m_macroSplashLockedExcel Then Application.Interactive = True
    m_macroSplashLockedExcel = False
    m_macroSplashShown = False
End Sub

Private Sub MacroSplash_Hide()
    On Error Resume Next
    MacroStartBgm_FadeOutAndClose
    m_splashConsoleOverlayActive = False
    If m_macroSplashShown Then
        Unload frmMacroSplash
    End If
    m_macroSplashShown = False
    If m_macroSplashLockedExcel Then
        Application.Interactive = True
        m_macroSplashLockedExcel = False
    End If
End Sub

' Python?ixlwings?j??????BPM_AI_SPLASH_XLWINGS=1 ???????s?????z??B?}?N??????????????? PM_AI_XLWINGS_SPLASH_MACRO=?W?????W???[????.SplashLog_AppendChunk
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

' ?A?j???t??_* ???????o???F?X?v???b?V???\?? ?? ?}?N?????s?i????????2???? Application.Run ?????j
' lockExcelUI?FFalse = InputBox?^?t?H???g?_?C?A???O??? Excel ???b???K?v??}?N??????
' allowMacroSound?FTrue = ?i?K1?^?i?K2????l?? BGM?E???????`???C????????i???? False?j
Private Sub ?A?j???t??_?X?v???b?V???t??????s(ByVal splashMessage As String, ByVal procName As String, Optional ByVal arg1 As Variant, Optional ByVal arg2 As Variant, Optional ByVal lockExcelUI As Boolean = True, Optional ByVal allowMacroSound As Boolean = False)
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

' =========================================================
' ???????????{?^????????????????}?N??
' =========================================================
' ?O???f?[?V?????z?F?v???Z?b?g?iCreateCoolButtonWithPreset ?? presetId?j
' 1=???C?????u???[ 2=?e?B?[?? 3=?I?????W 4=?t?H???X?g?O???[?? 5=?p?[?v??
' 6=?C???f?B?S 7=?X???[?g 8=?R?[???? 9=?A???o?[ 10=?}?[???^

#End If
