Private Function スプラッシュ_実行ログ領域の画面ピクセル矩形を取得(ByRef outL As Long, ByRef outT As Long, ByRef outW As Long, ByRef outH As Long) As Boolean
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
    スプラッシュ_実行ログ領域の画面ピクセル矩形を取得 = False
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
    スプラッシュ_実行ログ領域の画面ピクセル矩形を取得 = True
    Exit Function
Fail:
    スプラッシュ_実行ログ領域の画面ピクセル矩形を取得 = False
End Function

' cmd をログ枠に重ねるとき、二重描画を避けるため TextBox を一時非表示
Private Sub スプラッシュ_コンソールオーバーレイ開始()
    On Error Resume Next
    If Not m_macroSplashShown Then Exit Sub
    If m_splashConsoleOverlayActive Then Exit Sub
    frmMacroSplash.Controls("txtExecutionLog").Visible = False
    m_splashConsoleOverlayActive = True
    frmMacroSplash.Repaint
    On Error GoTo 0
End Sub

Private Sub スプラッシュ_コンソールオーバーレイ終了()
    On Error Resume Next
    If Not m_splashConsoleOverlayActive Then Exit Sub
    m_splashConsoleOverlayActive = False
    If m_macroSplashShown Then
        frmMacroSplash.Controls("txtExecutionLog").Visible = True
        frmMacroSplash.Repaint
    End If
    On Error GoTo 0
End Sub

' コンソール制御（Win32/Win64。WT ホスト HWND に効かない場合があるためオーバーレイ時は conhost 直起動を併用）
#If VBA7 Then
Public Sub スプラッシュ_手順文を設定(ByVal stepMessage As String)
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

Private Sub スプラッシュ_実行ログ枠をクリア()
    Dim tb As Object
    On Error Resume Next
    m_splashReadErrShown = False
    m_splashLastLogSnapshot = ""
    m_splashPollHaveCachedFileLen = False
    m_splashPollLastFileLen = 0
    If Not m_macroSplashShown Then Exit Sub
    If Not 設定シート_スプラッシュログ書込み有効か() Then Exit Sub
    Set tb = frmMacroSplash.Controls("txtExecutionLog")
    If Not tb Is Nothing Then tb.text = ""
End Sub

' ログは常に最新。キャレットを末尾に移動し txtExecutionLog へフォーカス（UserForm 自体に SetFocus は不可）
Private Sub スプラッシュ_テキストボックス末尾へスクロール(ByVal tb As Object)
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

' m_splashExecutionLogPath の UTF-8 ログを txtExecutionLog に反映（変更がない場合は再描画しない）
Private Sub スプラッシュ_実行ログ枠を更新()
    Dim tb As Object
    Dim s As String
    Dim n As Long
    Dim prevSU As Boolean
    Dim errBanner As String
    Dim flen As Long
    Dim flenAtStart As Long
    On Error Resume Next
    If Not 設定シート_スプラッシュログ書込み有効か() Then Exit Sub
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
    s = Gemini_UTF8ファイルを読込(m_splashExecutionLogPath)
    If Len(s) = 0 And Len(Dir(m_splashExecutionLogPath)) > 0 Then
        s = Gemini_UTF8ファイルを一時コピーで読込(m_splashExecutionLogPath)
    End If
    n = Len(s)
    If n > 0 Then
        m_splashReadErrShown = False
        If n > SPLASH_LOG_MAX_DISPLAY_CHARS Then
            s = "（長文ログのため末尾のみ表示）" & vbCrLf & Right$(s, SPLASH_LOG_MAX_DISPLAY_CHARS)
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
        スプラッシュ_テキストボックス末尾へスクロール tb
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
    errBanner = "【ログ表示エラー】execution_log.txt を VBA で読めませんでした（Python がファイルを開いている可能性）。" & vbCrLf & _
        "LOG シートまたはエディタで次のパスを直接開いて確認してください: " & m_splashExecutionLogPath & vbCrLf & vbCrLf
    prevSU = Application.ScreenUpdating
    If Not prevSU Then Application.ScreenUpdating = True
    tb.text = errBanner & tb.text
    m_splashLastLogSnapshot = tb.text
    tb.SelStart = 1
    tb.SelLength = 0
    frmMacroSplash.lblMessage.Caption = "（実行ログの表示に失敗 - 上記【ログ表示エラー】を参照）"
    frmMacroSplash.Repaint
    DoEvents
    If Not prevSU Then Application.ScreenUpdating = False
    m_splashReadErrShown = True
End Sub

' ダイアログ付き_段階2を実行 終了後など、途中ポーリングできなかった場合に execution_log を一括表示（Interactive 一時 True で描画）
Private Sub スプラッシュ_実行ログをパスから読込(ByVal fullPath As String)
    Dim tb As Object
    Dim s As String
    Dim n As Long
    Dim prevInt As Boolean
    Dim errBanner As String
    Dim flen As Long
    On Error Resume Next
    If Not 設定シート_スプラッシュログ書込み有効か() Then Exit Sub
    If Not m_macroSplashShown Then Exit Sub
    If Len(Dir(fullPath)) = 0 Then Exit Sub
    Set tb = frmMacroSplash.Controls("txtExecutionLog")
    If tb Is Nothing Then Exit Sub
    s = Gemini_UTF8ファイルを読込(fullPath)
    If Len(s) = 0 Then s = Gemini_UTF8ファイルを一時コピーで読込(fullPath)
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
            frmMacroSplash.lblMessage.Caption = "（実行ログの一括表示に失敗 - 上記参照）"
            frmMacroSplash.Repaint
            DoEvents
            If m_macroSplashLockedExcel Then Application.Interactive = False Else Application.Interactive = prevInt
        End If
        Exit Sub
    End If
    m_splashReadErrShown = False
    If n > SPLASH_LOG_MAX_DISPLAY_CHARS Then
        s = "（長文ログのため末尾のみ表示）" & vbCrLf & Right$(s, SPLASH_LOG_MAX_DISPLAY_CHARS)
    End If
    prevInt = Application.Interactive
    Application.Interactive = True
    tb.text = s
    m_splashLastLogSnapshot = s
    スプラッシュ_テキストボックス末尾へスクロール tb
    If m_macroSplashLockedExcel Then
        Application.Interactive = False
    Else
        Application.Interactive = prevInt
    End If
End Sub

Private Sub スプラッシュ_Excel下端中央にドッキング()
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

' モードレス UserForm を前面に維持し、再描画・ログ更新が止まったように見えることを防ぐ。Show 直後も前面化する。
Private Sub スプラッシュ_フォームを最前面へ()
#If VBA7 Then
    Dim hwnd As LongPtr
#Else
    Dim hwnd As Long
#End If
    On Error Resume Next
    If Not m_macroSplashShown Then Exit Sub
    スプラッシュ_Excel下端中央にドッキング
    hwnd = FindWindow(0&, SPLASH_FORM_WINDOW_TITLE)
    If hwnd = 0 Then Exit Sub
    BringWindowToTop hwnd
    SetForegroundWindow hwnd
    On Error GoTo 0
End Sub

Private Sub スプラッシュ_表示(Optional ByVal message As String, Optional ByVal lockExcelUI As Boolean = True)
    On Error GoTo CleanupFail
    If m_macroSplashShown Then スプラッシュ_非表示
    m_animMacroSucceeded = False
    If Len(Trim$(message)) = 0 Then
        message = "処理中です。しばらくお待ちください。"
    End If
    frmMacroSplash.Caption = SPLASH_FORM_WINDOW_TITLE
    frmMacroSplash.lblMessage.Caption = message
    frmMacroSplash.StartUpPosition = 2  ' 初期表示のみ。直後に スプラッシュ_Excel下端中央にドッキング で Excel 下端へ移動
    m_macroSplashLockedExcel = False
    If lockExcelUI Then
        Application.Interactive = False
        m_macroSplashLockedExcel = True
    End If
    frmMacroSplash.Show vbModeless
    m_macroSplashShown = True
    On Error Resume Next
    frmMacroSplash.Controls("txtExecutionLog").HideSelection = False
    スプラッシュ_フォームを最前面へ
    DoEvents
    起動BGM_利用可能なら開始
    Exit Sub
CleanupFail:
    On Error Resume Next
    If m_macroSplashLockedExcel Then Application.Interactive = True
    m_macroSplashLockedExcel = False
    m_macroSplashShown = False
End Sub

Private Sub スプラッシュ_非表示()
    On Error Resume Next
    起動BGM_フェードアウトして閉じる
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

' Python（xlwings）から呼ぶ。PM_AI_SPLASH_XLWINGS=1 のときのみ実行される想定。
' マクロ名を環境変数 PM_AI_XLWINGS_SPLASH_MACRO=標準モジュール名.スプラッシュログ_チャンクを追記 で指定可能。
Public Sub スプラッシュログ_チャンクを追記(ByVal chunk As String)
    On Error Resume Next
    If Len(chunk) = 0 Then Exit Sub
    If Not 設定シート_スプラッシュログ書込み有効か() Then Exit Sub
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
    スプラッシュ_テキストボックス末尾へスクロール tb
End Sub

' アニメ付き_* から呼ぶ：スプラッシュ表示中にマクロ実行（引数は最大2つまで Application.Run へ委譲）
' lockExcelUI=False は InputBox / フォントダイアログ等で Excel 操作が必要なマクロ向け
' allowMacroSound=True で 段階1/段階2 と同様に BGM・完了チャイムを許可（通常は False）
Private Sub アニメ付き_スプラッシュ付きで実行(ByVal splashMessage As String, ByVal procName As String, Optional ByVal arg1 As Variant, Optional ByVal arg2 As Variant, Optional ByVal lockExcelUI As Boolean = True, Optional ByVal allowMacroSound As Boolean = False)
    m_splashAllowMacroSound = allowMacroSound
    On Error GoTo EH
    スプラッシュ_表示 splashMessage, lockExcelUI
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
    起動BGM_フェードアウトして閉じる
    If m_animMacroSucceeded Then
        On Error Resume Next
        完了チャイムを再生処理
    End If
    スプラッシュ_非表示
    m_splashAllowMacroSound = False
End Sub

' =========================================================
' かっこいいボタンを自動生成するマクロ
' =========================================================
' グラデーション配色プリセット（クールボタンをプリセットで作成 の presetId）
' 1=ロイヤルブルー 2=ティール 3=オレンジ 4=フォレストグリーン 5=パープル
' 6=インディゴ 7=スレート 8=コーラル 9=アンバー 10=マゼンタ

#End If
