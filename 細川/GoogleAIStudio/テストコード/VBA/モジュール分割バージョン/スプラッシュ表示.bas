Option Explicit

' UserForm「frmMacroSplash」の既定グローバルインスタンス（VB_PredeclaredId）に依存しない。
' 手作業で追加した UserForm は事前宣言 ID が無効になり「frmMacroSplash が未定義」になることがあるため、New で確保する。
Private m_frmMacroSplash As frmMacroSplash

Private Function MacroSplash_Form() As frmMacroSplash
    If m_frmMacroSplash Is Nothing Then
        Set m_frmMacroSplash = New frmMacroSplash
    End If
    Set MacroSplash_Form = m_frmMacroSplash
End Function

Public Function MacroSplash_GetTxtExecutionLogScreenRectPixels(ByRef outL As Long, ByRef outT As Long, ByRef outW As Long, ByRef outH As Long) As Boolean
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
    Set tb = MacroSplash_Form.Controls("txtExecutionLog")
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

' cmd をログ枠に重ねるとき、二重表示を避けるため TextBox を一時非表示
Public Sub MacroSplash_BeginConsoleOverlay()
    On Error Resume Next
    If Not m_macroSplashShown Then Exit Sub
    If m_splashConsoleOverlayActive Then Exit Sub
    MacroSplash_Form.Controls("txtExecutionLog").Visible = False
    m_splashConsoleOverlayActive = True
    MacroSplash_Form.Repaint
    On Error GoTo 0
End Sub

Public Sub MacroSplash_EndConsoleOverlay()
    On Error Resume Next
    If Not m_splashConsoleOverlayActive Then Exit Sub
    m_splashConsoleOverlayActive = False
    If m_macroSplashShown Then
        MacroSplash_Form.Controls("txtExecutionLog").Visible = True
        MacroSplash_Form.Repaint
    End If
    On Error GoTo 0
End Sub

' コンソール枠の簡易除去（Win32/Win64。WT ホスト HWND には無効な場合あり ― オーバーレイは conhost 強制と併用）
Public Sub MacroSplash_SetStep(ByVal stepMessage As String)
    Dim prevSU As Boolean
    On Error Resume Next
    If Not m_macroSplashShown Then Exit Sub
    prevSU = Application.ScreenUpdating
    If Not prevSU Then Application.ScreenUpdating = True
    MacroSplash_Form.lblMessage.Caption = stepMessage
    MacroSplash_Form.Repaint
    DoEvents
    If Not prevSU Then Application.ScreenUpdating = False
End Sub

Public Sub MacroSplash_ClearExecutionLogPane()
    Dim tb As Object
    On Error Resume Next
    m_splashReadErrShown = False
    m_splashLastLogSnapshot = ""
    m_splashPollHaveCachedFileLen = False
    m_splashPollLastFileLen = 0
    If Not m_macroSplashShown Then Exit Sub
    If Not SettingsSheet_IsSplashExecutionLogWriteEnabled() Then Exit Sub
    Set tb = MacroSplash_Form.Controls("txtExecutionLog")
    If Not tb Is Nothing Then tb.text = ""
End Sub

' ログは末尾が最新。キャレットを最後に置き txtExecutionLog にフォーカス（UserForm には SetFocus がない）
Public Sub MacroSplash_TextBoxScrollToTail(ByVal tb As Object)
    On Error Resume Next
    tb.HideSelection = False
    tb.SelStart = Len(tb.text)
    tb.SelLength = 0
    If Application.Interactive Then
        tb.SetFocus
    End If
    MacroSplash_Form.Repaint
    DoEvents
End Sub

' m_splashExecutionLogPath の UTF-8 ログを txtExecutionLog へ（長いときは末尾のみ）
Public Sub MacroSplash_RefreshExecutionLogPane()
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
    Set tb = MacroSplash_Form.Controls("txtExecutionLog")
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
            s = "…（冒頭を省略。直近のみ表示）…" & vbCrLf & Right$(s, SPLASH_LOG_MAX_DISPLAY_CHARS)
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
    errBanner = "【ログ表示エラー】execution_log.txt を VBA から読めませんでした（Python がファイルを開いている等）。LOG シートまたはエディタで " & m_splashExecutionLogPath & " を直接開いて確認してください。" & vbCrLf & vbCrLf
    prevSU = Application.ScreenUpdating
    If Not prevSU Then Application.ScreenUpdating = True
    tb.text = errBanner & tb.text
    m_splashLastLogSnapshot = tb.text
    tb.SelStart = 1
    tb.SelLength = 0
    MacroSplash_Form.lblMessage.Caption = "…（実行ログの表示に失敗 ? 下記の【ログ表示エラー】を参照）"
    MacroSplash_Form.Repaint
    DoEvents
    If Not prevSU Then Application.ScreenUpdating = False
    m_splashReadErrShown = True
End Sub

' RunPython 終了直後など、同期 Python でポーリングできなかったあとに execution_log を一括表示（Interactive 一時 True で描画）
Public Sub MacroSplash_LoadExecutionLogFromPath(ByVal fullPath As String)
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
    Set tb = MacroSplash_Form.Controls("txtExecutionLog")
    If tb Is Nothing Then Exit Sub
    s = GeminiReadUtf8File(fullPath)
    If Len(s) = 0 Then s = GeminiReadUtf8FileViaTempCopy(fullPath)
    n = Len(s)
    If n = 0 Then
        flen = FileLen(fullPath)
        If Err.Number <> 0 Then Err.Clear
        If flen > 0 Or Len(tb.text) > 0 Then
            errBanner = "【ログ表示エラー】execution_log.txt を読み込めませんでした。パス: " & fullPath & vbCrLf & vbCrLf
            prevInt = Application.Interactive
            Application.Interactive = True
            tb.text = errBanner & tb.text
            m_splashLastLogSnapshot = tb.text
            MacroSplash_Form.lblMessage.Caption = "…（実行ログの一括表示に失敗 ? 下記を参照）"
            MacroSplash_Form.Repaint
            DoEvents
            If m_macroSplashLockedExcel Then Application.Interactive = False Else Application.Interactive = prevInt
        End If
        Exit Sub
    End If
    m_splashReadErrShown = False
    If n > SPLASH_LOG_MAX_DISPLAY_CHARS Then
        s = "…（冒頭を省略。直近のみ表示）…" & vbCrLf & Right$(s, SPLASH_LOG_MAX_DISPLAY_CHARS)
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

Public Sub MacroSplash_PositionDockExcelBottomCenter()
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

' モードレス UserForm が背後に残ると再描画・ログ更新が止まったように見えることがある。Show 直後に前面へ（ユーザーがクリックしたときと同趣旨）。
Public Sub MacroSplash_BringFormToFront()
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

Public Sub MacroSplash_Show(Optional ByVal message As String, Optional ByVal lockExcelUI As Boolean = True)
    On Error GoTo CleanupFail
    If m_macroSplashShown Then MacroSplash_Hide
    m_animMacroSucceeded = False
    If Len(Trim$(message)) = 0 Then
        message = "処理中です。しばらくお待ちください。"
    End If
    MacroSplash_Form.Caption = SPLASH_FORM_WINDOW_TITLE
    MacroSplash_Form.lblMessage.Caption = message
    MacroSplash_Form.StartUpPosition = 2  ' 初期のみ。直後に MacroSplash_PositionDockExcelBottomCenter で Excel 下端中央へ
    m_macroSplashLockedExcel = False
    If lockExcelUI Then
        Application.Interactive = False
        m_macroSplashLockedExcel = True
    End If
    ' Show の Modal=False（または省略）でモードレス。vbModeless は参照設定によって未定義になることがある
    MacroSplash_Form.Show False
    m_macroSplashShown = True
    On Error Resume Next
    MacroSplash_Form.Controls("txtExecutionLog").HideSelection = False
    MacroSplash_BringFormToFront
    DoEvents
    MacroStartBgm_StartIfAvailable
    Exit Sub
CleanupFail:
    On Error Resume Next
    If m_macroSplashLockedExcel Then Application.Interactive = True
    m_macroSplashLockedExcel = False
    m_macroSplashShown = False
    If Not m_frmMacroSplash Is Nothing Then
        Unload m_frmMacroSplash
    End If
    Set m_frmMacroSplash = Nothing
End Sub

Public Sub MacroSplash_Hide()
    On Error Resume Next
    MacroStartBgm_FadeOutAndClose
    m_splashConsoleOverlayActive = False
    If m_macroSplashShown Then
        If Not m_frmMacroSplash Is Nothing Then
            Unload m_frmMacroSplash
        End If
        Set m_frmMacroSplash = Nothing
    End If
    m_macroSplashShown = False
    If m_macroSplashLockedExcel Then
        Application.Interactive = True
        m_macroSplashLockedExcel = False
    End If
End Sub

' Python（xlwings）から呼ぶ。PM_AI_SPLASH_XLWINGS=1 時のみ実行される想定。マクロ名衝突時は環境変数 PM_AI_XLWINGS_SPLASH_MACRO=標準モジュール名.SplashLog_AppendChunk
Public Sub SplashLog_AppendChunk(ByVal chunk As String)
    On Error Resume Next
    If Len(chunk) = 0 Then Exit Sub
    If Not SettingsSheet_IsSplashExecutionLogWriteEnabled() Then Exit Sub
    If Not m_macroSplashShown Then Exit Sub
    Dim tb As Object
    Set tb = MacroSplash_Form.Controls("txtExecutionLog")
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

' アニメ付き_* から呼び出し：スプラッシュ表示 → マクロ実行（引数は最大2つまで Application.Run に委譲）
' lockExcelUI：False = InputBox／フォントダイアログなど Excel 対話が必要なマクロ向け
' allowMacroSound：True = 段階1／段階2と同様に BGM・成功時チャイムを許可（既定 False）
