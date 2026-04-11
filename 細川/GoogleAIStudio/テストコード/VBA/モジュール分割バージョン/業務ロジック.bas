<<<<<<< HEAD
Option Explicit

Public Function EnsureStageBatchStdoutRedirect(ByVal body As String) As String
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

Public Function RunTempCmdWithConsoleLayout(ByVal wsh As Object, ByVal body As String, Optional ByVal applyTopQuarterFullWidthConsole As Boolean = False, Optional ByVal hideCmdWindow As Boolean = False) As Long
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
=======
Attribute VB_Name = "業務ロジック"
Option Explicit

Private Function AugmentCmdBodyWithConsoleTitle(ByVal body As String, ByVal titleText As String) As String
    Const echoOffCrLf As String = "@echo off" & vbCrLf
    If Len(body) >= Len(echoOffCrLf) And LCase$(Left$(body, 9)) = "@echo off" Then
        If Mid$(body, 10, 2) = vbCrLf Then
            AugmentCmdBodyWithConsoleTitle = echoOffCrLf & "title " & titleText & vbCrLf & Mid$(body, 12)
            Exit Function
>>>>>>> hosokawa/main2
        End If
    End If
    AugmentCmdBodyWithConsoleTitle = echoOffCrLf & "title " & titleText & vbCrLf & body
End Function

<<<<<<< HEAD
' =========================================================
' ★ 図形に登録するためのアニメーション付き起動マクロ ★
' 処理本体は 段階1_コア実行 / 段階2_コア実行、ダイアログ付きの公開入口は RunPythonStage1 / RunPython / RunPythonStage1ThenStage2
' 段階1・段階2のコアが成功で終わった直後、配台_全シートフォントBIZ_UDP_自動適用 で全シートを BIZ UDPゴシックに統一し、結果_主要4結果シート_列オートフィット で主要4結果シートの列幅を調整（完了の vbInformation MsgBox は使わずスプラッシュ＋システム音）
' 段階2 Finish: 取り込み成功時は「結果_」で始まる全シートの表示倍率を 100% にし、その後 結果_設備ガント のみ 85% に戻す。結果_設備毎の時間割(B2)・結果_タスク一覧(F2)・結果_カレンダー(出勤簿)(A2) で窓枠固定を付与したうえで、最後にメインシート A1 をアクティブにして終了する
' =========================================================
=======
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

>>>>>>> hosokawa/main2
Sub アニメ付き_計画生成を実行()
    Call AnimateButtonPush
    アニメ付き_スプラッシュ付きで実行 "シミュレーション（計画生成）を実行しています…", "RunPython", False, , True, True
End Sub

<<<<<<< HEAD
' 段階1: 加工計画DATA からタスク抽出 → output に xlsx 出力し「配台計画_タスク入力」へ取り込み
=======
>>>>>>> hosokawa/main2
Sub アニメ付き_タスク抽出を実行()
    Call AnimateButtonPush
    アニメ付き_スプラッシュ付きで実行 "タスク抽出（段階1）を実行しています…", "RunPythonStage1", , , True, True
End Sub

<<<<<<< HEAD
' 段階1→保存反映→段階2を続けて実行（配台計画シートの手編集を挟まない一括実行）
=======
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' 図形ボタン用：Caller が取れるのは「この Sub が OnAction のとき」だけ。本体を直接割り当てるとアニメは動かない。
Sub アニメ付き_全シートフォントをリストから選択して統一()
    Call AnimateButtonPush
    ' xlDialogFormatFont 表示のためグリッド操作ブロックは使わない
    アニメ付き_スプラッシュ付きで実行 "全シートのフォントを一覧から選んで統一しています…", "全シートフォントをリストから選択して統一", , , False
End Sub

Sub アニメ付き_全シートフォントを手入力で統一()
    Call AnimateButtonPush
    ' Application.InputBox 用にグリッド操作ブロックは使わない
    アニメ付き_スプラッシュ付きで実行 "全シートのフォントを手入力の名前で統一しています…", "全シートフォントを手入力で統一", , , False
End Sub

Sub アニメ付き_全シートフォント_BIZ_UDPゴシックに統一()
    Call AnimateButtonPush
    アニメ付き_スプラッシュ付きで実行 "全シートのフォントを BIZ UDP ゴシックに統一しています…", "全シートフォント_BIZ_UDPゴシックに統一"
End Sub

' =========================================================
' Gemini API キーを暗号化 JSON にし「設定」B1 にパスを書く（押下アニメ付きはアニメ付き_* を図形に割当）
' 暗号化パスフレーズは InputBox で入力し --passphrase-file 経由で Python に渡す。B2 にはパスフレーズを書かない。
' Python: python\encrypt_gemini_credentials.py（要 cryptography）。起動は py -3 を推奨。
' =========================================================
Sub アニメ付き_Gemini認証を暗号化してB1に保存()
    Call AnimateButtonPush
    ' InputBox 等があるためグリッド操作ブロックは使わない（スプラッシュのみ）
    アニメ付き_スプラッシュ付きで実行 "Gemini 認証を暗号化して保存しています…", "設定_Gemini認証を暗号化してB1に保存", , , False
End Sub

' 列設定シートの内容を「結果_タスク一覧」へ反映（Python）。図形の OnAction には本マクロを指定（本体を直指定するとアニメは動かない）。
=======
>>>>>>> hosokawa/main2
Sub アニメ付き_列設定_結果_タスク一覧_列順表示をPython適用()
    Call AnimateButtonPush
    アニメ付き_スプラッシュ付きで実行 "列設定を結果タスク一覧に反映しています…", "列設定_結果_タスク一覧_列順表示をPython適用"
End Sub

<<<<<<< HEAD
' 列設定シート A:B のみ重複列名を削除（結果シートは触らない）。図形には「アニメ付き_列設定_結果_タスク一覧_重複列名を整理」。
=======
>>>>>>> hosokawa/main2
Sub アニメ付き_列設定_結果_タスク一覧_重複列名を整理()
    Call AnimateButtonPush
    アニメ付き_スプラッシュ付きで実行 "列設定シートの重複列名を整理しています…", "列設定_結果_タスク一覧_重複列名を整理"
End Sub

<<<<<<< HEAD
' 配台計画_タスク入力: 「配台不要」を手動でクリアしたあと等に試行順を付け直す。図形の OnAction は本マクロ（本体直指定だと AnimateButtonPush が動かない）。
Sub アニメ付き_配台計画_タスク入力_配台試行順番を再計算()
=======
Sub メインシート_メンバー一覧と出勤表示_手動()
    メインシート_メンバー一覧と出勤表示 False
End Sub

Sub アニメ付き_メインシート_masterブックを開く()
>>>>>>> hosokawa/main2
    Call AnimateButtonPush
    アニメ付き_スプラッシュ付きで実行 "配台試行順番を再計算しています…", "配台計画_タスク入力_配台試行順番をPythonで再計算"
End Sub

<<<<<<< HEAD
' 配台試行順番列を小数キーとして昇順に並べ替え 1..n（マスタ・上書き連携なし）。図形 OnAction は本マクロ。
Sub アニメ付き_配台計画_タスク入力_試行順を小数キーで並べ替え()
    Call AnimateButtonPush
    アニメ付き_スプラッシュ付きで実行 "配台試行順番をキー順に並べ替えています…", "配台計画_タスク入力_試行順を小数キーでPython並べ替え"
End Sub

' 上記と同じ図形をシートに自動配置（初回・位置調整用）。本体は フォント管理 の 配台計画_タスク入力_配台試行順再計算ボタンを配置。
Sub アニメ付き_配台計画_タスク入力_配台試行順再計算ボタンを配置()
    Call AnimateButtonPush
    配台計画_タスク入力_配台試行順再計算ボタンを配置
End Sub

' 小数キー並べ替えボタンを配置（グラデーション図形）。本体は フォント管理。
Sub アニメ付き_配台計画_タスク入力_試行順小数キー並べ替えボタンを配置()
    Call AnimateButtonPush
    配台計画_タスク入力_試行順小数キー並べ替えボタンを配置
End Sub

' 小数キー並べ替えボタンを配置（かっこいいボタン版）。
Sub アニメ付き_配台計画_タスク入力_試行順小数キー並べ替えクールボタンを配置()
    Call AnimateButtonPush
    配台計画_タスク入力_試行順小数キー並べ替え_クールボタンを配置
End Sub

Public Function GetMainWorksheet() As Worksheet
=======
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
>>>>>>> hosokawa/main2
    ' 配台ブックのメイン UI はシート名「メイン_」固定（旧「メイン」「Main」や部分一致は使わない）
    On Error Resume Next
    Set GetMainWorksheet = ThisWorkbook.Worksheets("メイン_")
    On Error GoTo 0
End Function

<<<<<<< HEAD
' planning_core が log\gemini_usage_summary_for_main.txt（UTF-8）に出力した Gemini 利用サマリを、
' メインシート P16 以降に反映する。openpyxl が保存できない（ブック開きっぱなし）とき用。段階1/2 コアから呼ぶ。
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

' メインシートを表示し A1 をアクティブにする（シート名「メイン_」＝GetMainWorksheet と同じ）
Public Sub メインシートA1を選択()
=======
Private Sub メインシートA1を選択()
>>>>>>> hosokawa/main2
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetMainWorksheet()
    If ws Is Nothing Then Exit Sub
    ws.Activate
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

<<<<<<< HEAD
' Ctrl+Shift+テンキー - 用（手続き名は従来互換で CtrlShift0 のまま。Application.OnKey の Procedure は ASCII 名が無難）
Public Function FindColHeader(ws As Worksheet, ByVal headerText As String) As Long
=======
Private Function FindColHeader(ws As Worksheet, ByVal headerText As String) As Long
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' ハイパーリンク再付与後に既定の青リンク体へ戻らないよう、クリア前のフォントを記憶・復元する
Public Sub メインシート_フォント属性を取得( _
    ByVal src As Range, _
    ByRef fn As String, ByRef fs As Double, ByRef fc As Variant, _
    ByRef fBold As Boolean, ByRef fItalic As Boolean, ByRef fUl As Long)
    On Error Resume Next
    fn = "": fs = 0: fc = Empty: fBold = False: fItalic = False: fUl = xlUnderlineStyleNone
    If src Is Nothing Then Exit Sub
    With src.Font
        fn = .Name
        fs = .Size
        fc = .Color
        fBold = .Bold
        fItalic = .Italic
        fUl = .Underline
    End With
    On Error GoTo 0
End Sub

Public Sub メインシート_フォント属性を適用( _
    ByVal tgt As Range, _
    ByVal fn As String, ByVal fs As Double, ByVal fc As Variant, _
    ByVal fBold As Boolean, ByVal fItalic As Boolean, ByVal fUl As Long)
    On Error Resume Next
    If tgt Is Nothing Then Exit Sub
    With tgt.Font
        If Len(fn) > 0 Then .Name = fn
        If fs > 0 Then .Size = fs
        If Not IsEmpty(fc) Then .Color = fc
        .Bold = fBold
        .Italic = fItalic
        .Underline = fUl
    End With
    On Error GoTo 0
End Sub

' 範囲内の各セルに付いたハイパーリンクだけ削除（書式は維持）
Public Sub メインシート_指定範囲のハイパーリンクを削除(ByVal wsMain As Worksheet, ByVal Target As Range)
=======
Private Sub メインシート_指定範囲のハイパーリンクを削除(ByVal wsMain As Worksheet, ByVal Target As Range)
>>>>>>> hosokawa/main2
    Dim c As Range
    If wsMain Is Nothing Or Target Is Nothing Then Exit Sub
    On Error Resume Next
    For Each c In Target.Cells
        If c.Hyperlinks.Count > 0 Then c.Hyperlinks.Delete
    Next c
    On Error GoTo 0
End Sub

<<<<<<< HEAD
' メインシート A1～：ブック内で名前が「結果_」で始まるシートへのジャンプリンクを並べる
Public Sub メインシート_結果シートリンクを更新(ByVal wsMain As Worksheet)
=======
Private Sub メインシート_結果シートリンクを更新(ByVal wsMain As Worksheet)
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' 結果_*（設備ガント以外）・個人_*: 実験コードと同じ手順で列オートフィット
' ・呼び出し元が ScreenUpdating=False のとき、Select 前に True に戻さないと AutoFit が効かないことがある
' ・元の ScreenUpdating は必ず復帰
' ・引数名は targetWs（RunPython 等の呼び出し側にも「ws」があり、ウォッチで親フレームの ws と混同しやすいため）
Public Sub 結果シート_列幅_AutoFit安定(ByVal targetWs As Worksheet)
=======
Private Sub 結果シート_列幅_AutoFit安定(ByVal targetWs As Worksheet)
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' 結果_タスク一覧 専用: 非表示列に EntireColumn.AutoFit をかけると列が再表示されるため、表示列のみ AutoFit する。
Public Sub 結果シート_列幅_AutoFit非表示を維持(ByVal targetWs As Worksheet)
=======
Private Sub 結果シート_列幅_AutoFit非表示を維持(ByVal targetWs As Worksheet)
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' 結果_タスク一覧: 列「配完_回答指定16時まで」（旧名「配完_基準16時まで」）が「いいえ」のセルを赤背景・白文字・太字にする。
' 段階2の xlsx 取り込み直後に呼ぶ（openpyxl 側の書式に加え、列幅調整後の見た目を確実にする）。
Public Sub 結果_タスク一覧_配完回答指定16時_いいえを強調(ByVal ws As Worksheet)
=======
Private Sub 結果_タスク一覧_配完回答指定16時_いいえを強調(ByVal ws As Worksheet)
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' 過去の取り込み不具合で残った「列設定_結果_タスク一覧 (2)」等を削除（本体シートは残す）
' ※呼び出し元が DisplayAlerts=False のとき、終了時に True に戻さない（シート削除確認が出るのを防ぐ）
Public Sub 列設定結果タスク一覧_番号付き重複シートを削除(ByVal wb As Workbook)
=======
Private Sub 列設定結果タスク一覧_番号付き重複シートを削除(ByVal wb As Workbook)
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' 段階2 取り込み後: 「設定」タブの直前に列設定シートを置く（個人_*・LOG・設定の並べ替えの後に呼ぶ）
Public Sub 列設定_結果_タスク一覧を設定の直前へ移動(ByVal wb As Workbook)
=======
Private Sub 列設定_結果_タスク一覧を設定の直前へ移動(ByVal wb As Workbook)
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' 結果_* シートの1行目右側に、メインシートへ戻る内部リンクを1つ置く（取り込み直後に呼ぶ）
Public Sub 結果シート_メインへ戻るリンクを付与(ByVal ws As Worksheet)
=======
Private Sub 結果シート_メインへ戻るリンクを付与(ByVal ws As Worksheet)
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' 段階2 完了時: 結果_設備毎の時間割 を表示し B2 を選択、1行目と A 列を窓枠固定。シートが無ければ False。
Public Function 結果_設備毎の時間割_B2選択して窓枠固定() As Boolean
=======
Private Function 結果_設備毎の時間割_B2選択して窓枠固定() As Boolean
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' 段階2 完了時: 結果_タスク一覧 を表示し F2 を選択、1 行目と A～E 列を窓枠固定。シートが無ければ False。
Public Function 結果_タスク一覧_F2選択して窓枠固定() As Boolean
=======
Private Function 結果_タスク一覧_F2選択して窓枠固定() As Boolean
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' 段階2 完了時: 結果_カレンダー(出勤簿) を表示し A2 を選択、1 行目を窓枠固定。シートが無ければ False。
Public Function 結果_カレンダー出勤簿_A2選択して窓枠固定() As Boolean
=======
Private Function 結果_カレンダー出勤簿_A2選択して窓枠固定() As Boolean
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' 段階2 完了間際: 名前が「結果_」で始まるシートの表示倍率を指定％にする（各シートを一度アクティブにして ActiveWindow.Zoom を設定）
Public Sub 結果プレフィックスシートの表示倍率を設定(ByVal wb As Workbook, ByVal zoomPercent As Long)
=======
Private Sub 結果プレフィックスシートの表示倍率を設定(ByVal wb As Workbook, ByVal zoomPercent As Long)
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' 結果_設備ガントのみ表示倍率を設定（シートをアクティブにして ActiveWindow.Zoom）
Public Sub 結果_設備ガント_表示倍率を設定(ByVal wb As Workbook, ByVal zoomPercent As Long)
=======
Private Sub 結果_設備ガント_表示倍率を設定(ByVal wb As Workbook, ByVal zoomPercent As Long)
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' 結果_設備ガント：取り込み直後に列幅を設定（Python 本体では列幅を書かない）
Public Sub 結果_設備ガント_列幅を設定(ByVal ws As Worksheet)
=======
Private Sub 結果_設備ガント_列幅を設定(ByVal ws As Worksheet)
>>>>>>> hosokawa/main2
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
        ws.Columns(c).ColumnWidth = 3   ' 時刻グリッド（F 列?）
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

<<<<<<< HEAD
' 結果_設備ガント：タイトルは結合セル先頭 A1。取り込み後の 1 行目一括書式のあと左寄せが崩れることがあるため固定する。
Public Sub 結果_設備ガント_タイトルA1を左寄せに固定(ByVal ws As Worksheet)
=======
Private Sub 結果_設備ガント_タイトルA1を左寄せに固定(ByVal ws As Worksheet)
>>>>>>> hosokawa/main2
    On Error Resume Next
    With ws.Range("A1")
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
End Sub

<<<<<<< HEAD
' planning_core のガント罫線 thin color 666666 に合わせる（ハイライト解除時）
Public Sub 結果_設備ガント_行枠を通常に戻す(ByVal rng As Range)
=======
Private Sub 結果_設備ガント_行枠を通常に戻す(ByVal rng As Range)
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
Public Sub 結果_設備ガント_行枠を強調(ByVal rng As Range)
=======
Private Sub 結果_設備ガント_行枠を強調(ByVal rng As Range)
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
Public Function 結果_設備ガント_行は表頭行か(ByVal ws As Worksheet, ByVal r As Long) As Boolean
=======
Private Function 結果_設備ガント_行は表頭行か(ByVal ws As Worksheet, ByVal r As Long) As Boolean
>>>>>>> hosokawa/main2
    Dim a As String
    Dim b As String
    On Error Resume Next
    a = Trim$(CStr(ws.Cells(r, 2).Value))
    b = Trim$(CStr(ws.Cells(r, 3).Value))
    On Error GoTo 0
    結果_設備ガント_行は表頭行か = (StrComp(a, "機械名", vbBinaryCompare) = 0 And StrComp(b, "工程名", vbBinaryCompare) = 0)
End Function

<<<<<<< HEAD
' Python が挿入する日付ブロック間の黒帯（行高さ約 5pt）
Public Function 結果_設備ガント_行は区切り行か(ByVal ws As Worksheet, ByVal r As Long) As Boolean
=======
Private Function 結果_設備ガント_行は区切り行か(ByVal ws As Worksheet, ByVal r As Long) As Boolean
>>>>>>> hosokawa/main2
    Dim rh As Double
    On Error Resume Next
    rh = ws.Rows(r).RowHeight
    On Error GoTo 0
    If rh > 0# And rh <= 5.6 Then 結果_設備ガント_行は区切り行か = True
End Function

<<<<<<< HEAD
' タイトル・メタ（1?2 行）・表頭・区切り行以外＝計画行・実績行（日付は A 列縦結合のため列 A にも ? が現れる）
Public Function 結果_設備ガント_行はデータ行か(ByVal ws As Worksheet, ByVal r As Long) As Boolean
=======
Private Function 結果_設備ガント_行はデータ行か(ByVal ws As Worksheet, ByVal r As Long) As Boolean
>>>>>>> hosokawa/main2
    結果_設備ガント_行はデータ行か = False
    If r <= 2 Then Exit Function
    If 結果_設備ガント_行は表頭行か(ws, r) Then Exit Function
    If 結果_設備ガント_行は区切り行か(ws, r) Then Exit Function
    If r < 4 Then Exit Function
    結果_設備ガント_行はデータ行か = True
End Function

<<<<<<< HEAD
Public Sub 結果_設備ガント_行ハイライト_Clear(ByVal wb As Workbook)
=======
Private Sub 結果_設備ガント_行ハイライト_Clear(ByVal wb As Workbook)
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' ThisWorkbook.SheetSelectionChange から呼ぶ（標準モジュールは Public 必須）
=======
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' 既存ブックで設備ガントが「セルの書式設定」無効の保護のとき、ハイライト罫線が効かない。パスワードは SHEET_FONT_UNPROTECT_PASSWORD のみ対応（手動パスワードはユーザーが一度解除してから実行）
=======
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' 段階1/2 終盤・全シートフォント適用後: 結果の主要4シートの列オートフィット
' ・結果_タスク一覧 は非表示列を開かない（結果シート_列幅_AutoFit非表示を維持）
' ・結果_設備ガント は専用列幅（時刻列を潰さない）＋タイトル A1 左寄せ再固定
=======
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' =========================================================
' シート並び：個人_*（名前昇順）→ その後ろに LOG → 最後に「設定」
' （シート名は正確に LOG / 設定。無い場合はスキップ）
' =========================================================
=======
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' =========================================================
' シート並び：配台計画_タスク入力を前へ
' （step1完了時点で「個人_*」「LOG」「設定」の前の方へ配置）
' =========================================================
Public Sub 配台計画_タスク入力を前へ並べ替え()
=======
Private Sub 配台計画_タスク入力を前へ並べ替え()
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' =========================================================
' 共通：ボタンを押し込むアニメーション処理
' ※ActiveSheet.Shapes(名前) だけだと、別シートに同じ図形名（既定の角丸1 等）があると
'   誤ってそちらを動かし、意図しないシートが前面に出ることがあります。
'   全シートから名前を解決し、ActiveSheet 上のものを優先します。
' =========================================================
Public Sub AnimateButtonPush()
=======
Private Sub AnimateButtonPush()
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' =========================================================
' マクロ実行中スプラッシュ（擬似モーダル）
' ・シート「設定」D3: true/TRUE でログ枠へ書き込み＋Exec 待機中のファイルポーリング。false/FALSE で無し・同期 Run・通常 cmd 表示（log\execution_log.txt への Python 出力は変わらず）
' ・シート「設定」D4: マクロ成功時の完了チャイム用 MP3 トラック番号 1?4（空・不正は 1）。ファイル名は標準モジュール MACRO_COMPLETE_MP3_1?4。sounds フォルダに配置。MP3 が無い／再生失敗時は macro_complete_chime.wav
' ・段階1／段階2のスプラッシュのみ: BGM（sounds 配下の Glass_Architecture1.mp3 等）を MCI ループ再生。終了時はフェードアウト後に close（完了チャイムより先）。他マクロのスプラッシュでは BGM・チャイムは再生しない
' ・UserForm「frmMacroSplash」をプロジェクトに追加（未追加時は表示せず続行）
' ・lockExcelUI=True のとき Application.Interactive=False でブック操作をブロック（対話マクロは False）
' ・ただし Interactive=False のままだと UserForm の再描画が滞り execution_log ポーリングが見えにくい。段階1/2 の Exec 待機中は一時的に True に戻す（RunCmdFileStageExecAndPoll）。
' ・終了・エラー時は必ず MacroSplash_Hide で Interactive を戻す
' ・作成手順とフォームコードは frmMacroSplash_VBA.txt
' ・完了の vbInformation MsgBox は原則やめ、段階1／段階2成功時はスプラッシュ最終文＋完了チャイム（MacroCompleteChime・設定 D4・sounds\*.mp3／WAV・失敗時 SystemAsterisk）
' =========================================================
Public Function SettingsSheet_IsSplashExecutionLogWriteEnabled() As Boolean
=======
Private Function SettingsSheet_IsSplashExecutionLogWriteEnabled() As Boolean
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
Public Function SettingsSheet_GetCompleteChimeTrack1to4() As Long
=======
Private Function SettingsSheet_GetCompleteChimeTrack1to4() As Long
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
Public Sub アニメ付き_スプラッシュ付きで実行(ByVal splashMessage As String, ByVal procName As String, Optional ByVal arg1 As Variant, Optional ByVal arg2 As Variant, Optional ByVal lockExcelUI As Boolean = True, Optional ByVal allowMacroSound As Boolean = False)
=======
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
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' =========================================================
' かっこいいボタンを自動生成するマクロ
' =========================================================
' グラデーション配色プリセット（CreateCoolButtonWithPreset の presetId）
' 1=ロイヤルブルー 2=ティール 3=オレンジ 4=フォレストグリーン 5=パープル
' 6=インディゴ 7=スレート 8=コーラル 9=アンバー 10=マゼンタ
Public Function CoolButtonGradientTop(ByVal presetId As Long) As Long
=======
Private Function CoolButtonGradientTop(ByVal presetId As Long) As Long
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
Public Function CoolButtonGradientBottom(ByVal presetId As Long) As Long
=======
Private Function CoolButtonGradientBottom(ByVal presetId As Long) As Long
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
Public Sub CreateCoolButtonWithPreset(btnText As String, macroName As String, posX As Single, posY As Single, ByVal presetId As Long, Optional stableShapeName As String)
    CreateCoolButton btnText, macroName, posX, posY, CoolButtonGradientTop(presetId), CoolButtonGradientBottom(presetId), stableShapeName
=======
Private Sub CreateCoolButtonWithPreset(btnText As String, macroName As String, posX As Single, posY As Single, ByVal presetId As Long)
    CreateCoolButton btnText, macroName, posX, posY, CoolButtonGradientTop(presetId), CoolButtonGradientBottom(presetId)
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' 配色プリセット P1～P10 の見本を配置（マクロは割り当てず、見た目確認・色選び用）
=======
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' アクティブシート上に、グラデーション＋押下アニメ用のクールボタンを1つ配置（InputBox で文言・マクロ名・座標・配色を指定）
' 割り当て先は「アニメ付き_*」など、先頭で AnimateButtonPush を呼ぶマクロを推奨（図形に本体を直割り当てするとアニメは動きません）
Public Sub アニメ付きマクロ用_クールボタンを対話配置()
    Dim cap As String
    Dim mac As String
    Dim ps As String
    Dim pr As Long
    Dim x As Single
    Dim y As Single
    Dim stable As String
    
    cap = InputBox( _
        "ボタンに表示する文字列を入力してください。", _
        "アニメ付きクールボタン (1/4)", _
        "実行")
    If Len(Trim$(cap)) = 0 Then Exit Sub
    
    mac = InputBox( _
        "割り当てるマクロ名を入力してください。" & vbCrLf & _
        "例: アニメ付き_計画生成を実行（このブック内の Public Sub 名）", _
        "アニメ付きクールボタン (2/4)", _
        "アニメ付き_計画生成を実行")
    If Len(Trim$(mac)) = 0 Then Exit Sub
    
    ps = InputBox( _
        "左位置と上位置をカンマ区切りで入力（ポイント）。例: 50, 120" & vbCrLf & _
        "空欄なら 50, 50 を使います。", _
        "アニメ付きクールボタン (3/4)", _
        "50, 50")
    If Len(Trim$(ps)) = 0 Then ps = "50, 50"
    If Not ParseTwoSingleCsv(ps, x, y) Then
        MsgBox "位置の形式が不正です。例: 50, 120", vbExclamation
        Exit Sub
    End If
    
    ps = InputBox( _
        "配色プリセット番号（1～10）を入力してください。" & vbCrLf & _
        "1=ロイヤルブルー … 10=マゼンタ（CreateCoolButtonWithPreset と同じ）", _
        "アニメ付きクールボタン (4/4)", _
        "1")
    pr = 1
    If Len(Trim$(ps)) > 0 And IsNumeric(ps) Then pr = CLng(CDbl(ps))
    If pr < 1 Or pr > 10 Then pr = 1
    
    Randomize
    stable = "AnimCool_" & Format(Now, "yyyymmddhhnnss") & "_" & Format(Int(1000000 * Rnd), "000000")
    
    CreateCoolButtonWithPreset Trim$(cap), Trim$(mac), x, y, pr, stable
    MsgBox "クールボタンを配置しました。" & vbCrLf & _
           "図形名: " & stable & vbCrLf & _
           "OnAction: " & Trim$(mac), vbInformation
End Sub

' カンマ区切りで2つの Single を読む（空白許容）
Private Function ParseTwoSingleCsv(ByVal s As String, ByRef outX As Single, ByRef outY As Single) As Boolean
    Dim p As Long
    Dim a As String
    Dim b As String
    p = InStr(1, s, ",")
    If p <= 0 Then Exit Function
    a = Trim$(Left$(s, p - 1))
    b = Trim$(Mid$(s, p + 1))
    If Len(a) = 0 Or Len(b) = 0 Then Exit Function
    If Not IsNumeric(a) Or Not IsNumeric(b) Then Exit Function
    outX = CSng(CDbl(a))
    outY = CSng(CDbl(b))
    ParseTwoSingleCsv = True
End Function

' ボタン生成の共通ロジック（stableShapeName を渡すと図形名を固定。AnimateButtonPush は Application.Caller=図形名のためアニメ付きマクロ用ボタンでは推奨）
Public Sub CreateCoolButton(btnText As String, macroName As String, posX As Single, posY As Single, colorTop As Long, colorBottom As Long, Optional stableShapeName As String)
=======
Private Sub CreateCoolButton(btnText As String, macroName As String, posX As Single, posY As Single, colorTop As Long, colorBottom As Long)
>>>>>>> hosokawa/main2
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
        If Len(Trim$(stableShapeName)) > 0 Then
            .Name = stableShapeName
        Else
            Randomize
            .Name = "CoolBtn_" & Format(Now, "yyyymmddhhnnss") & "_" & Format(Int(1000000 * Rnd), "000000")
        End If
        On Error GoTo 0
    End With
End Sub

<<<<<<< HEAD
' 「配台計画_タスク入力」1行目の右側付近に、試行順再計算用のクールボタンを1つ配置する（同一 OnAction の既存図形は削除してから作成）。
Public Sub 配台計画_タスク入力_配台試行順番再計算ボタンを配置()
    Const MACRO_ANIM As String = "アニメ付き_配台計画_タスク入力_配台試行順番を再計算"
    Const HDR_TRIAL As String = "配台試行順番"
    Dim ws As Worksheet
    Dim shp As Shape
    Dim oa As String
    Dim lastCol As Long
    Dim anchorCol As Long
    Dim leftPos As Single
    Dim topPos As Single
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_PLAN_INPUT_TASK)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "シート「" & SHEET_PLAN_INPUT_TASK & "」がありません。", vbExclamation, "ボタン配置"
        Exit Sub
    End If
    
    ws.Activate
    
    For Each shp In ws.Shapes
        On Error Resume Next
        oa = shp.OnAction
        On Error GoTo 0
        If InStr(1, oa, MACRO_ANIM, vbBinaryCompare) > 0 Then
            On Error Resume Next
            shp.Delete
            On Error GoTo 0
        End If
    Next shp
    
    anchorCol = FindColHeader(ws, HDR_TRIAL)
    If anchorCol <= 0 Then anchorCol = 1
    leftPos = ws.Cells(1, anchorCol).Left + ws.Cells(1, anchorCol).Width + 8
    topPos = ws.Cells(1, 1).Top + 4
    
    CreateCoolButtonWithPreset "試行順を更新", MACRO_ANIM, leftPos, topPos, 2
    MsgBox "「" & SHEET_PLAN_INPUT_TASK & "」にボタンを配置しました。" & vbCrLf & _
           "「配台不要」の手動クリア後などに押すと、Python で試行順を再計算して行を並べ替えます。", vbInformation, "ボタン配置"
End Sub

' 「配台試行順番」を小数キーで並べ替え 1..n 用（かっこいいボタン版）。試行順更新ボタンの下あたりに配置（同一マクロ割当の既存図形は削除）。
' グラデーション版は フォント管理 の「配台計画_タスク入力_試行順小数キー並べ替えボタンを配置」。
Public Sub 配台計画_タスク入力_試行順小数キー並べ替え_クールボタンを配置()
    Const MACRO_ANIM As String = "アニメ付き_配台計画_タスク入力_試行順を小数キーで並べ替え"
    Const HDR_TRIAL As String = "配台試行順番"
    Dim ws As Worksheet
    Dim shp As Shape
    Dim oa As String
    Dim anchorCol As Long
    Dim leftPos As Single
    Dim topPos As Single
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_PLAN_INPUT_TASK)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "シート「" & SHEET_PLAN_INPUT_TASK & "」がありません。", vbExclamation, "ボタン配置"
        Exit Sub
    End If
    
    ws.Activate
    
    For Each shp In ws.Shapes
        On Error Resume Next
        oa = shp.OnAction
        On Error GoTo 0
        If InStr(1, oa, MACRO_ANIM, vbBinaryCompare) > 0 Then
            On Error Resume Next
            shp.Delete
            On Error GoTo 0
        End If
    Next shp
    
    anchorCol = FindColHeader(ws, HDR_TRIAL)
    If anchorCol <= 0 Then anchorCol = 1
    leftPos = ws.Cells(1, anchorCol).Left + ws.Cells(1, anchorCol).Width + 8
    topPos = ws.Cells(1, 1).Top + 4 + 58
    
    CreateCoolButtonWithPreset "キー順に並べ替え", MACRO_ANIM, leftPos, topPos, 3
    MsgBox "「" & SHEET_PLAN_INPUT_TASK & "」にボタンを配置しました。" & vbCrLf & _
           "配台試行順番に 1, 2, 1.5 などを入れたあと押すと、キー昇順に行を並べ 1 から振り直します。", vbInformation, "ボタン配置"
End Sub

' =========================================================
' ① Python本体と必要なコンポーネントをインストールするマクロ（修正版）
' ・Python 3 の検出は py -3（ランチャーで 3 系を明示）
' ・未導入時: winget（Python.Python.3.12）→ 失敗時は公式 amd64 インストーラ
' ・pip は PowerShell 内で Machine/User の PATH を再合成（Excel 起動後でも py を拾いやすく）
' ・pip 依存は setup_environment.py が requirements.txt を読み込んで一括（cryptography 含む）。スクリプトは python\setup_environment.py を優先（旧: ブック直下）
' ・xlwings: 本ブックの段階1/2 は WScript.Shell で py を起動するため planning_core 側では未使用。
'           将来 Excel から xlwings（RunPython / UDF 等）で Python を呼ぶ拡張に備え、pip とアドインを導入する。
' ・setup_environment.py の最後で「xlwings addin install」（Excel アドインを XLSTART へ）
' ※公式 URL はモジュール先頭の PY_OFFICIAL_INSTALLER_URL で変更可能
' =========================================================
Public Sub DisableBackgroundDataRefreshAll()
=======
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
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' PQ 更新前: 接続先 IP へ ping 1 回（-w でタイムアウト）。成功時のみ True。
' 失敗時はデータ更新をスキップし、呼び出し元は従来どおり True で継続する。
Public Function PingHostOnceBeforeQueryRefresh(ByVal ipAddress As String, ByVal timeoutMs As Long) As Boolean
=======
Private Function PingHostOnceBeforeQueryRefresh(ByVal ipAddress As String, ByVal timeoutMs As Long) As Boolean
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' =========================================================
' Power Query / データ接続の更新（マクロ処理の先頭で呼ぶ）
' ※ 先に DisableBackgroundDataRefreshAll で同期更新に寄せ、RefreshAll 後に
'    CalculateUntilAsyncQueriesDone で取りこぼし待ち（背景オフ後はほぼ即時）。
'    これにより「未実行のデータ更新が取り消されます」系ダイアログを抑止しやすくする。
' ※ DisplayAlerts=False で接続／PQ 失敗時の Excel 標準ダイアログを抑止。VBA 側も MsgBox は出さず
'    m_lastRefreshQueriesErrMsg に詳細を残す（段階1・2のエラーメッセージに連結）。
' ※ PQ_REFRESH_PING_HOST へ ping（PQ_REFRESH_PING_TIMEOUT_MS）で応答がなければ RefreshAll は行わず、
'    成功として返す（既存データのまま段階1・2を継続）。
' =========================================================
Public Function TryRefreshWorkbookQueries() As Boolean
=======
Private Function TryRefreshWorkbookQueries() As Boolean
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' Python の execution_log は UTF-8(BOM 付き)。cmd の 2>&1 リダイレクトは環境で Shift_JIS になりがちなので BOM で切り替える。
Public Function ValidateMasterSkillsOpAsPriorityUnique(ByVal targetDir As String, ByRef errOut As String) As Boolean
    Dim wbPath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim openedHere As Boolean
    Dim lastCol As Long
    Dim lastRow As Long
    Dim c As Long
    Dim r As Long
    Dim pHdr As String
    Dim mHdr As String
    Dim combo As String
    Dim mem As String
    Dim cellV As String
    Dim roleCh As String
    Dim prVal As Long
    Dim okCell As Boolean
    Dim pmCount As Long
    Dim dict As Object
    Dim headerRow As Long
    Dim memCol As Long
    
    errOut = ""
    ValidateMasterSkillsOpAsPriorityUnique = False
    wbPath = targetDir & "\master.xlsm"
    If Len(Dir(wbPath)) = 0 Then
        errOut = "master.xlsm が見つかりません: " & wbPath
        Exit Function
    End If
    
    openedHere = False
    Set wb = Nothing
    On Error Resume Next
    Set wb = Workbooks("master.xlsm")
    On Error GoTo 0
    If wb Is Nothing Then
        On Error GoTo OpenFailSkills
        Set wb = Workbooks.Open(Filename:=wbPath, ReadOnly:=True, UpdateLinks:=0)
        openedHere = True
        On Error GoTo 0
    End If
    
    On Error Resume Next
    Set ws = wb.Worksheets("skills")
    On Error GoTo 0
    If ws Is Nothing Then
        errOut = "master.xlsm に「skills」シートがありません。"
        If openedHere Then
            On Error Resume Next
            wb.Close SaveChanges:=False
            On Error GoTo 0
        End If
        Exit Function
    End If
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastCol < 2 Or lastRow < 2 Then GoTo CloseOkSkills
    
    pmCount = 0
    For c = 2 To lastCol
        pHdr = Trim$(CStr(ws.Cells(1, c).Value))
        mHdr = Trim$(CStr(ws.Cells(2, c).Value))
        If Len(pHdr) > 0 And Len(mHdr) > 0 Then
            If StrComp(LCase$(pHdr), "nan", vbTextCompare) <> 0 And StrComp(LCase$(mHdr), "nan", vbTextCompare) <> 0 Then
                pmCount = pmCount + 1
            End If
        End If
    Next c
    
    If pmCount > 0 Then
        For c = 2 To lastCol
            pHdr = Trim$(CStr(ws.Cells(1, c).Value))
            mHdr = Trim$(CStr(ws.Cells(2, c).Value))
            If Len(pHdr) = 0 Or Len(mHdr) = 0 Then GoTo NextColTwo
            If StrComp(LCase$(pHdr), "nan", vbTextCompare) = 0 Or StrComp(LCase$(mHdr), "nan", vbTextCompare) = 0 Then GoTo NextColTwo
            combo = pHdr & "+" & mHdr
            Set dict = CreateObject("Scripting.Dictionary")
            For r = 3 To lastRow
                mem = Trim$(CStr(ws.Cells(r, 1).Value))
                If Len(mem) = 0 Or StrComp(LCase$(mem), "nan", vbTextCompare) = 0 Then GoTo NextRowTwo
                cellV = Trim$(CStr(ws.Cells(r, c).Value))
                okCell = ParseOpAsSkillCellForValidate(cellV, roleCh, prVal)
                If Not okCell Then GoTo NextRowTwo
                If Not dict.Exists(CStr(prVal)) Then
                    dict.Add CStr(prVal), mem & "(" & roleCh & ")"
                Else
                    errOut = "マスタ skills の優先度の数値が重複しています。" & vbCrLf & _
                        "列「" & combo & "」: 優先度 " & CStr(prVal) & " が重複（" & dict(CStr(prVal)) & " と " & mem & "(" & roleCh & ")）" & vbCrLf & _
                        "master.xlsm を修正してから再実行してください。"
                    Set dict = Nothing
                    GoTo CloseFailSkills
                End If
NextRowTwo:
            Next r
            Set dict = Nothing
NextColTwo:
        Next c
    Else
        headerRow = 1
        memCol = 1
        For c = 2 To lastCol
            combo = Trim$(CStr(ws.Cells(headerRow, c).Value))
            If Len(combo) = 0 Or StrComp(LCase$(combo), "nan", vbTextCompare) = 0 Then GoTo NextColOne
            Set dict = CreateObject("Scripting.Dictionary")
            For r = headerRow + 1 To lastRow
                mem = Trim$(CStr(ws.Cells(r, memCol).Value))
                If Len(mem) = 0 Or StrComp(LCase$(mem), "nan", vbTextCompare) = 0 Then GoTo NextRowOne
                cellV = Trim$(CStr(ws.Cells(r, c).Value))
                okCell = ParseOpAsSkillCellForValidate(cellV, roleCh, prVal)
                If Not okCell Then GoTo NextRowOne
                If Not dict.Exists(CStr(prVal)) Then
                    dict.Add CStr(prVal), mem & "(" & roleCh & ")"
                Else
                    errOut = "マスタ skills の優先度の数値が重複しています。" & vbCrLf & _
                        "列「" & combo & "」: 優先度 " & CStr(prVal) & " が重複（" & dict(CStr(prVal)) & " と " & mem & "(" & roleCh & ")）" & vbCrLf & _
                        "master.xlsm を修正してから再実行してください。"
                    Set dict = Nothing
                    GoTo CloseFailSkills
                End If
NextRowOne:
            Next r
            Set dict = Nothing
NextColOne:
        Next c
    End If

CloseOkSkills:
    If openedHere Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    ValidateMasterSkillsOpAsPriorityUnique = True
    Exit Function

CloseFailSkills:
    If openedHere Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    Exit Function

OpenFailSkills:
    errOut = "master.xlsm を開けませんでした: " & wbPath
End Function

' =========================================================
' 段階1: master.xlsm から「機械カレンダー」とメンバー別勤怠シートを本ブックへ置換コピーする（保護は段階1/2 終了時にまとめて適用）。
'         ※編集は master 側で実施し、段階1でスナップショットを同期。
'         ※配台用 Python は master.xlsm を直接参照するため本コピーは任意。既定は STAGE1_SYNC_MASTER_SHEETS_TO_MACRO_BOOK=0 でスキップ。
' =========================================================
Public Function 段階1_マスタシートを本ブックへ置換コピー( _
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

' master を開き（未オープンなら ReadOnly）、skills の A 列メンバー名に対応するシートを同期。戻り値=LOG に載せる短文。
Public Function 段階1_マスタ勤怠と機械カレンダーを同期し保護(ByVal targetDir As String) As String
    Dim wbPath As String
    Dim wb As Workbook
    Dim openedHere As Boolean
    Dim wsSkill As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim pmCount As Long
    Dim c As Long
    Dim pHdr As String
    Dim mHdr As String
    Dim r As Long
    Dim mem As String
    Dim startRow As Long
    Dim seen As Object
    Dim nOk As Long
    Dim nSkip As Long
    Dim wsMc As Worksheet
    Dim wm As Worksheet
    Dim parts As String
    
    段階1_マスタ勤怠と機械カレンダーを同期し保護 = ""
    wbPath = targetDir & "\" & MASTER_WORKBOOK_FILE
    If Len(Dir(wbPath)) = 0 Then Exit Function
    
    openedHere = False
    Set wb = Nothing
    On Error Resume Next
    Set wb = Workbooks(MASTER_WORKBOOK_FILE)
    On Error GoTo 0
    If wb Is Nothing Then
        On Error Resume Next
        Set wb = Workbooks.Open(Filename:=wbPath, ReadOnly:=True, UpdateLinks:=0)
        openedHere = Not (wb Is Nothing)
        On Error GoTo 0
    End If
    If wb Is Nothing Then
        段階1_マスタ勤怠と機械カレンダーを同期し保護 = "マスタ同期: ブックを開けませんでした"
        Exit Function
    End If
    
    On Error Resume Next
    Set wsMc = wb.Worksheets(SHEET_MACHINE_CALENDAR)
    On Error GoTo 0
    If Not wsMc Is Nothing Then
        If 段階1_マスタシートを本ブックへ置換コピー(wb, SHEET_MACHINE_CALENDAR) Then
            nOk = nOk + 1
        Else
            nSkip = nSkip + 1
        End If
    End If
    
    On Error Resume Next
    Set wsSkill = wb.Worksheets("skills")
    On Error GoTo 0
    If wsSkill Is Nothing Then GoTo CloseMasterWbSt1
    
    lastCol = wsSkill.Cells(1, wsSkill.Columns.Count).End(xlToLeft).Column
    lastRow = wsSkill.Cells(wsSkill.Rows.Count, 1).End(xlUp).Row
    If lastCol < 2 Or lastRow < 2 Then GoTo CloseMasterWbSt1
    
    pmCount = 0
    For c = 2 To lastCol
        pHdr = Trim$(CStr(wsSkill.Cells(1, c).Value))
        mHdr = Trim$(CStr(wsSkill.Cells(2, c).Value))
        If Len(pHdr) > 0 And Len(mHdr) > 0 Then
            If StrComp(LCase$(pHdr), "nan", vbTextCompare) <> 0 And StrComp(LCase$(mHdr), "nan", vbTextCompare) <> 0 Then
                pmCount = pmCount + 1
            End If
        End If
    Next c
    
    If pmCount > 0 Then
        startRow = 3
    Else
        startRow = 2
    End If
    
    Set seen = CreateObject("Scripting.Dictionary")
    For r = startRow To lastRow
        mem = Trim$(CStr(wsSkill.Cells(r, 1).Value))
        If Len(mem) = 0 Or StrComp(LCase$(mem), "nan", vbTextCompare) = 0 Then GoTo NextMemberSt1
        If seen.Exists(mem) Then GoTo NextMemberSt1
        seen.Add mem, True
        Set wm = Nothing
        On Error Resume Next
        Set wm = wb.Worksheets(mem)
        On Error GoTo 0
        If Not wm Is Nothing Then
            If 段階1_マスタシートを本ブックへ置換コピー(wb, mem) Then
                nOk = nOk + 1
            Else
                nSkip = nSkip + 1
            End If
        End If
NextMemberSt1:
    Next r
    
CloseMasterWbSt1:
    If openedHere Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    parts = "マスタ同期: 「" & SHEET_MACHINE_CALENDAR & "」+ メンバー勤怠をコピーし保護（成功シート数 " & CStr(nOk)
    If nSkip > 0 Then
        parts = parts & "・失敗 " & CStr(nSkip)
    End If
    parts = parts & "）※シート保護はマクロ終了時にまとめて適用"
    段階1_マスタ勤怠と機械カレンダーを同期し保護 = parts
End Function

' =========================================================
' 段階1/2・全シートフォント: 開始時に全シートを試行解除し、終了直前に「対象」だけ既定条件で再保護する
' （結果_* … 段階2 取込と同じオプション／機械カレンダー・マスタ skills のメンバー勤怠シート … 無パス UI のみブロック）
' =========================================================
Public Sub 配台マクロ_全シート保護を試行解除()
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

' master.xlsm の skills から、勤怠シート名が実在するメンバーのみ dict に追加（キーのみ使用）
Public Sub 配台_マスタSkillsから勤怠シート名を辞書に追加(ByVal targetDir As String, ByVal dict As Object)
    Dim wbPath As String
    Dim wb As Workbook
    Dim openedHere As Boolean
    Dim wsSkill As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim pmCount As Long
    Dim c As Long
    Dim pHdr As String
    Dim mHdr As String
    Dim r As Long
    Dim mem As String
    Dim startRow As Long
    Dim seen As Object
    Dim wm As Worksheet
    
    wbPath = targetDir & "\" & MASTER_WORKBOOK_FILE
    If Len(Dir(wbPath)) = 0 Then Exit Sub
    
    openedHere = False
    Set wb = Nothing
    On Error Resume Next
    Set wb = Workbooks(MASTER_WORKBOOK_FILE)
    On Error GoTo 0
    If wb Is Nothing Then
        On Error Resume Next
        Set wb = Workbooks.Open(Filename:=wbPath, ReadOnly:=True, UpdateLinks:=0)
        openedHere = Not (wb Is Nothing)
        On Error GoTo 0
    End If
    If wb Is Nothing Then Exit Sub
    
    On Error Resume Next
    Set wsSkill = wb.Worksheets("skills")
    On Error GoTo 0
    If wsSkill Is Nothing Then GoTo CloseMasterWbProt
    
    lastCol = wsSkill.Cells(1, wsSkill.Columns.Count).End(xlToLeft).Column
    lastRow = wsSkill.Cells(wsSkill.Rows.Count, 1).End(xlUp).Row
    If lastCol < 2 Or lastRow < 2 Then GoTo CloseMasterWbProt
    
    pmCount = 0
    For c = 2 To lastCol
        pHdr = Trim$(CStr(wsSkill.Cells(1, c).Value))
        mHdr = Trim$(CStr(wsSkill.Cells(2, c).Value))
        If Len(pHdr) > 0 And Len(mHdr) > 0 Then
            If StrComp(LCase$(pHdr), "nan", vbTextCompare) <> 0 And StrComp(LCase$(mHdr), "nan", vbTextCompare) <> 0 Then
                pmCount = pmCount + 1
            End If
        End If
    Next c
    
    If pmCount > 0 Then
        startRow = 3
    Else
        startRow = 2
    End If
    
    Set seen = CreateObject("Scripting.Dictionary")
    For r = startRow To lastRow
        mem = Trim$(CStr(wsSkill.Cells(r, 1).Value))
        If Len(mem) = 0 Or StrComp(LCase$(mem), "nan", vbTextCompare) = 0 Then GoTo NextMemberProt
        If seen.Exists(mem) Then GoTo NextMemberProt
        seen.Add mem, True
        Set wm = Nothing
        On Error Resume Next
        Set wm = wb.Worksheets(mem)
        On Error GoTo 0
        If Not wm Is Nothing Then
            If Not dict.Exists(mem) Then dict.Add mem, True
        End If
NextMemberProt:
    Next r

CloseMasterWbProt:
    If openedHere Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
End Sub

' targetDir 空なら ThisWorkbook.Path を使用（全シートフォント単体実行向け）
Public Sub 配台マクロ_対象シートを条件どおりに保護(Optional ByVal targetDir As String = "")
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

' =========================================================
' 段階1コア: task_extract_stage1.py → plan_input_tasks.xlsx →「配台計画_タスク入力」取込
' MsgBox は出さない。m_lastStage1ExitCode / m_lastStage1ErrMsg を参照（クエリ更新失敗時は m_lastRefreshQueriesErrMsg が連結される）
' =========================================================
Public Sub 段階1_コア実行()
    Dim wsh As Object
    Dim runBat As String
    Dim targetDir As String
    Dim wsLog As Worksheet
    Dim logFilePath As String
    Dim exitCode As Long
    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim adoStream As Object
    Dim outputText As String
    Dim logLines() As String
    Dim i As Long
    Dim warnRow As Long
    Dim st1XwErr As Long
    Dim st1XwDesc As String
    Dim missSt1 As String
    Dim st1DidUnlock As Boolean

    On Error GoTo ErrStage1
    m_lastStage1ExitCode = -1
    m_lastStage1ErrMsg = ""

    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    targetDir = ThisWorkbook.path
    If targetDir = "" Then
        m_lastStage1ErrMsg = "先にこのExcelファイルを保存してください。"
        Exit Sub
    End If

    Application.ScreenUpdating = False

    ' 接続更新より先に設定シートを確保（Refresh で止まる・失敗して Exit したとき無言でシート未作成になるのを防ぐ）
    MacroSplash_SetStep "段階1: 「設定_配台不要工程」シートを確認・作成・見出しを整えています…"
    設定_配台不要工程_シートを確保
    MacroSplash_SetStep "段階1: 「設定_環境変数」シートを確認・作成し不足キーのみ追記しています…"
    設定_環境変数_シートを確保
    MacroSplash_SetStep "段階1: 「設定_シート表示」シートを確認・作成しています…"
    設定_シート表示_シートを確保
    MacroSplash_SetStep "段階1: データ接続（Power Query 等）を更新しています…"

    If Not TryRefreshWorkbookQueries() Then
        m_lastStage1ErrMsg = "データ接続の更新に失敗したため段階1を中断しました。（「設定_配台不要工程」シートは作成済みの可能性があります）"
        If Len(m_lastRefreshQueriesErrMsg) > 0 Then
            m_lastStage1ErrMsg = m_lastStage1ErrMsg & vbCrLf & m_lastRefreshQueriesErrMsg
        End If
        Application.ScreenUpdating = prevScreenUpdating
        Application.DisplayAlerts = prevDisplayAlerts
        Exit Sub
    End If

    MacroSplash_SetStep "段階1: マスタ skills の運用優先度を検証しています…"
    Dim skErrSt1 As String
    If Not ValidateMasterSkillsOpAsPriorityUnique(targetDir, skErrSt1) Then
        m_lastStage1ErrMsg = skErrSt1
        m_lastStage1ExitCode = -1
        Application.ScreenUpdating = prevScreenUpdating
        Application.DisplayAlerts = prevDisplayAlerts
        Exit Sub
    End If

    MacroSplash_SetStep "段階1: ブックを保存し LOG シートを初期化します…"
    Application.StatusBar = "ブックを保存しています..."
    DoEvents
    On Error Resume Next
    ThisWorkbook.Save
    Application.StatusBar = False
    On Error GoTo ErrStage1

    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("LOG")
    On Error GoTo ErrStage1
    If wsLog Is Nothing Then
        m_lastStage1ErrMsg = "「LOG」シートが見つかりません。"
        Application.ScreenUpdating = prevScreenUpdating
        Application.DisplayAlerts = prevDisplayAlerts
        Exit Sub
    End If

    st1DidUnlock = False
    配台マクロ_全シート保護を試行解除
    st1DidUnlock = True

    MacroSplash_SetStep "段階1: LOG シートをクリアしヘッダを書き込んでいます…"
    wsLog.Cells.Clear
    wsLog.Cells(1, 1).Value = "実行ブック: " & ThisWorkbook.FullName
    If Not GeminiCredentialsJsonPathIsConfigured() Then
        wsLog.Cells(1, 2).Value = "【要設定】シート「設定」B1 に Gemini 認証 JSON のフルパス（例: Z:\社内\gemini_credentials.json）。gemini_credentials.example.json 参照。"
    End If
    Dim st1MasterSync As String
    If Stage1SyncMasterSheetsToMacroBookEffective() Then
        MacroSplash_SetStep "段階1: master.xlsm から勤怠・機械カレンダーを同期しています…"
        st1MasterSync = 段階1_マスタ勤怠と機械カレンダーを同期し保護(targetDir)
    Else
        MacroSplash_SetStep "段階1: master 勤怠のマクロブックへコピーをスキップ（配台は master.xlsm 直読み）…"
        st1MasterSync = "マスタ同期: スキップ（STAGE1_SYNC_MASTER_SHEETS_TO_MACRO_BOOK=0。配台は master.xlsm を直接参照）"
    End If
    If Len(st1MasterSync) > 0 Then
        wsLog.Cells(1, 3).Value = st1MasterSync
    End If
    Set wsh = CreateObject("WScript.Shell")
    wsh.Environment("Process")("TASK_INPUT_WORKBOOK") = ThisWorkbook.FullName

    On Error Resume Next
    Kill targetDir & "\log\execution_log.txt"
    Kill targetDir & "\log\stage_vba_exitcode.txt"
    On Error GoTo ErrStage1

    MacroSplash_SetStep "段階1: Python（task_extract）でタスク抽出を実行しています。完了までお待ちください…（詳細は LOG シート・log\execution_log.txt）"
    If STAGE12_USE_XLWINGS_RUNPYTHON And Not STAGE12_USE_XLWINGS_SPLASH_LOG Then
        wsh.Environment("Process")("PM_AI_SPLASH_XLWINGS") = ""
        m_splashExecutionLogPath = targetDir & "\log\execution_log.txt"
        m_stageVbaExitCodeLogDir = ""
        MacroSplash_ClearExecutionLogPane
        On Error Resume Next
        Err.Clear
        XwRunConsoleRunner "run_stage1_for_xlwings"
        If Err.Number <> 0 Then
            st1XwErr = Err.Number
            st1XwDesc = Err.Description
            Err.Clear
            On Error GoTo ErrStage1
            m_splashExecutionLogPath = ""
            m_stageVbaExitCodeLogDir = ""
            m_lastStage1ExitCode = -1
            m_lastStage1ErrMsg = "段階1: xlwings RunPython が失敗しました (" & CStr(st1XwErr) & "): " & st1XwDesc
            Application.StatusBar = False
            Application.ScreenUpdating = prevScreenUpdating
            Application.DisplayAlerts = prevDisplayAlerts
            If st1DidUnlock Then 配台マクロ_対象シートを条件どおりに保護 targetDir
            Exit Sub
        End If
        On Error GoTo ErrStage1
        exitCode = ReadStageVbaExitCodeFromFile(targetDir & "\log\stage_vba_exitcode.txt")
        If exitCode = &H7FFFFFFF Then exitCode = 1
        m_splashExecutionLogPath = ""
        m_stageVbaExitCodeLogDir = ""
        m_lastStage1ExitCode = exitCode
        MacroSplash_LoadExecutionLogFromPath targetDir & "\log\execution_log.txt"
    Else
        Dim hideStage12CmdSt1 As Boolean
        hideStage12CmdSt1 = Stage12CmdHideWindowEffective()
        wsh.CurrentDirectory = Environ("TEMP")
        ' 遅延環境変数で py 終了コードを exit /b し VBA に返す（一時 .cmd を cmd.exe /c で実行）
        ' 進捗表示は execution_log.txt のポーリングのみ（PM_AI_SPLASH_XLWINGS は使わない＝二重表示防止）
        runBat = "@echo off" & vbCrLf & "setlocal EnableDelayedExpansion" & vbCrLf & "pushd """ & targetDir & """" & vbCrLf & _
                 "if not exist log mkdir log" & vbCrLf & _
                 "chcp 65001>nul" & vbCrLf & _
                 "echo [stage1] Running Python... Progress below. See also LOG sheet and log\execution_log.txt" & vbCrLf & _
                 "py -3 -u python\task_extract_stage1.py" & vbCrLf & _
                 "set STAGE1_PY_EXIT=!ERRORLEVEL!" & vbCrLf & _
                 "echo." & vbCrLf & _
                 "echo [stage1] Finished. ERRORLEVEL=!STAGE1_PY_EXIT!" & vbCrLf & _
                 "(echo !STAGE1_PY_EXIT!)>log\stage_vba_exitcode.txt" & vbCrLf
        ' コンソール表示時のみ: Python 失敗後にウィンドウがすぐ閉じないよう pause（非表示・headless では付けない）
        If Not hideStage12CmdSt1 Then
            runBat = runBat & "if not !STAGE1_PY_EXIT! equ 0 (" & vbCrLf & _
                     "echo." & vbCrLf & _
                     "echo [stage1] Python error. Press any key to close this window..." & vbCrLf & _
                     "pause" & vbCrLf & _
                     ")" & vbCrLf
        End If
        runBat = runBat & "exit /b !STAGE1_PY_EXIT!"
        m_splashExecutionLogPath = targetDir & "\log\execution_log.txt"
        m_stageVbaExitCodeLogDir = ""
        MacroSplash_ClearExecutionLogPane
        exitCode = RunTempCmdWithConsoleLayout(wsh, runBat, Not hideStage12CmdSt1, hideStage12CmdSt1)
        m_splashExecutionLogPath = ""
        m_stageVbaExitCodeLogDir = ""
        m_lastStage1ExitCode = exitCode
    End If

    MacroSplash_SetStep "段階1: 配台不要工程シートへ TSV（A?D）を反映しています…"
    ' Gemini サマリ・設定シートは planning_core が openpyxl で保存（ブックが閉じているとき）。
    ' Excel で開いたままのときは log の TSV/テキストをマクロで反映する。
    On Error Resume Next
    Call 設定_配台不要工程_AからE_TSVから反映
    MacroSplash_SetStep "段階1: 配台不要工程シートの E 列（ロジック式）を TSV から反映しています…"
    Call 設定_配台不要工程_E列_TSVから反映
    MacroSplash_SetStep "段階1: メインシートの Gemini 利用サマリ（P 列）を反映しています…"
    Call メインシート_Gemini利用サマリをP列に反映(targetDir)
    On Error GoTo ErrStage1

    logFilePath = targetDir & "\log\execution_log.txt"
    If Len(Dir(logFilePath)) = 0 Then
        wsLog.Range("A2").Value = "execution_log.txt が見つかりませんでした。exitCode=" & CStr(exitCode)
        wsLog.Range("A3").Value = "xlwings 経路（STAGE12_USE_XLWINGS_RUNPYTHON=True）では Show Console の Python 出力も参照してください。runner は planning_core 読込前に log を作成するよう修正済みです。"
        missSt1 = Trim$(GeminiReadUtf8File(targetDir & "\log\stage2_blocking_message.txt"))
        If Len(missSt1) > 0 Then
            wsLog.Range("A4").Value = "log\stage2_blocking_message.txt: " & missSt1
        End If
    Else
        Set adoStream = CreateObject("ADODB.Stream")
        adoStream.charset = "UTF-8"
        adoStream.Open
        adoStream.LoadFromFile logFilePath
        outputText = adoStream.ReadText
        adoStream.Close
        Set adoStream = Nothing
        outputText = Replace(outputText, vbCrLf, vbLf)
        logLines = Split(outputText, vbLf)
        MacroSplash_SetStep "段階1: execution_log.txt の全文を LOG シートへ書き込んでいます…（行数 " & CStr(UBound(logLines) - LBound(logLines) + 1) & "）"
        Application.ScreenUpdating = False
        For i = LBound(logLines) To UBound(logLines)
            wsLog.Cells(i + 2, 1).Value = logLines(i)
        Next i
        If exitCode <> 0 Then
            warnRow = UBound(logLines) - LBound(logLines) + 3
            If warnRow < 1 Then warnRow = 2
            wsLog.Cells(warnRow, 1).Value = "■ Pythonの終了コード: " & CStr(exitCode) & " （詳細は上記ログ参照）"
        End If
    End If

    ' Python 失敗時はこの先（取り込み・フォント）をスキップ。フォント手前まで進んでから MsgBox すると原因が誤解されやすい。
    If exitCode <> 0 Then
        Application.ScreenUpdating = prevScreenUpdating
        Application.DisplayAlerts = prevDisplayAlerts
        If st1DidUnlock Then 配台マクロ_対象シートを条件どおりに保護 targetDir
        Exit Sub
    End If

    MacroSplash_SetStep "段階1: output\plan_input_tasks.xlsx を開き「配台計画_タスク入力」へ取り込んでいます…"
    If Not ImportPlanInputTasksFromOutput(targetDir) Then
        m_lastStage1ExitCode = -1
        Application.ScreenUpdating = prevScreenUpdating
        Application.DisplayAlerts = prevDisplayAlerts
        If st1DidUnlock Then 配台マクロ_対象シートを条件どおりに保護 targetDir
        Exit Sub
    End If
    On Error Resume Next
    配台計画_タスク入力を前へ並べ替え
    On Error GoTo 0

    MacroSplash_SetStep "段階1: フォント統一と表示調整を行っています…"
    Application.ScreenUpdating = True
    DoEvents
    On Error Resume Next
    配台_全シートフォントBIZ_UDP_自動適用
    On Error GoTo 0

    On Error Resume Next
    配台計画_タスク入力_A1を選択
    On Error GoTo 0

    MacroSplash_SetStep "段階1: 「設定_シート表示」を一覧更新しブックへ適用しています…"
    On Error Resume Next
    設定_シート表示_一覧をブックから再取得
    Err.Clear
    設定_シート表示_ブックへ適用
    Err.Clear
    On Error GoTo 0

    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts
    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0
    If st1DidUnlock Then 配台マクロ_対象シートを条件どおりに保護 targetDir
    Exit Sub

ErrStage1:
    m_lastStage1ExitCode = -1
    m_lastStage1ErrMsg = "段階1: " & Err.Description
    Application.StatusBar = False
    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts
    If st1DidUnlock Then 配台マクロ_対象シートを条件どおりに保護 targetDir
End Sub

' 互換・他モジュール用: 段階1のみ（エラー時 MsgBox。成功時はスプラッシュ文＋チャイム）
Public Function ImportPlanInputTasksFromOutput(ByVal targetDir As String) As Boolean
    Const PLAN_SHEET As String = "配台計画_タスク入力"
    Dim path As String
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim ws As Worksheet
    Dim da As Boolean
    Dim prevSUImp As Boolean
    Dim preserveFontName As String
    Dim preserveFontSize As Double
    Dim havePreserveFont As Boolean

    path = targetDir & "\output\plan_input_tasks.xlsx"
    If Len(Dir(path)) = 0 Then
        m_lastStage1ErrMsg = "plan_input_tasks.xlsx が見つかりません: " & path
        ImportPlanInputTasksFromOutput = False
        Exit Function
    End If

    da = Application.DisplayAlerts
    prevSUImp = Application.ScreenUpdating
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    MacroSplash_SetStep "段階1: plan_input_tasks.xlsx を開いています…"
    Set srcWb = Workbooks.Open(path)
    Set srcWs = srcWb.Sheets(1)

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(PLAN_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        MacroSplash_SetStep "段階1: 「配台計画_タスク入力」シートが無いため、出力ブックから新規シートとしてコピーしています…"
        srcWs.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        Set ws = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        On Error Resume Next
        ws.Name = PLAN_SHEET
        On Error GoTo 0
    Else
        MacroSplash_SetStep "段階1: 既存の「配台計画_タスク入力」をクリアし、出力データを貼り付けています…"
        preserveFontName = "": preserveFontSize = 0: havePreserveFont = False
        配台計画_タスク入力_既存シートの基準フォントを取得 ws, preserveFontName, preserveFontSize, havePreserveFont
        ws.Cells.Clear
        srcWs.UsedRange.Copy Destination:=ws.Range("A1")
    End If

    MacroSplash_SetStep "段階1: 取り込み元ブックを閉じ、列幅・罫線・配台試行順番ソートを適用しています…"
    srcWb.Close SaveChanges:=False
    Set srcWb = Nothing

    On Error Resume Next
    ws.UsedRange.Columns.AutoFit
    If havePreserveFont Then
        配台計画_タスク入力_UsedRangeにフォント名とサイズを適用 ws, preserveFontName, preserveFontSize
    End If
    With ws.UsedRange.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(226, 239, 218)
    End With
    ws.UsedRange.Borders.LineStyle = 1
    ws.UsedRange.Borders.Weight = 2

    ' 配台試行順番（昇順）でソートし、オートフィルタを有効化。列が無いときのみ従来どおり「指定納期」。
    ' ※ Python は to_excel 直前に試行順で並べ替え済みだが、貼り付け後にここで一度ソートすることで
    '   UsedRange の列検出・表示を安定させ、かつ試行順を正とする（指定納期だけだと順序が崩れる）。
    Dim colTrialOrder As Long
    Dim colSpecifiedDue As Long
    Dim sortCol As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim keyRange As Range
    Dim sortDataOpt As XlSortDataOption
    colTrialOrder = FindColHeader(ws, "配台試行順番")
    colSpecifiedDue = FindColHeader(ws, "指定納期")
    sortCol = 0
    sortDataOpt = xlSortNormal
    If colTrialOrder > 0 Then
        sortCol = colTrialOrder
        sortDataOpt = xlSortTextAsNumbers
    ElseIf colSpecifiedDue > 0 Then
        sortCol = colSpecifiedDue
    End If
    If sortCol > 0 Then
        lastRow = ws.Cells(ws.Rows.Count, sortCol).End(xlUp).Row
        lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        If lastRow >= 2 And lastCol >= 1 Then
            On Error Resume Next
            If ws.AutoFilterMode Then ws.AutoFilterMode = False
            On Error GoTo 0

            Set keyRange = ws.Range(ws.Cells(2, sortCol), ws.Cells(lastRow, sortCol))
            ws.Sort.SortFields.Clear
            ws.Sort.SortFields.Add Key:=keyRange, SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=sortDataOpt
            With ws.Sort
                .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .Apply
            End With

            ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter
        End If
    End If
    On Error GoTo 0

    ' 上書き入力列に薄い黄色（Python planning_core と同系色）? 取り込み後も確実に付与
    配台計画_タスク入力_上書き列に入力色を付与 ws

    Application.DisplayAlerts = da
    Application.ScreenUpdating = prevSUImp
    ImportPlanInputTasksFromOutput = True
End Function

' 配台計画_タスク入力シートを表示し A1 をアクティブにする（段階1終了時・取り込み後など）
Public Sub 配台計画_タスク入力_A1を選択()
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

' 配台計画_タスク入力: 上書き列（AI解析列を除く）のデータ行に入力用の薄黄色を付ける
Public Sub 配台計画_タスク入力_上書き列に入力色を付与(ByVal ws As Worksheet)
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

' 段階2: Python がブック保存で矛盾ハイライトを書けなかったとき、output の TSV を読み開いているシートへ適用
Public Sub ApplyPlanningConflictHighlightSidecar()
    Const SIDECAR As String = "planning_conflict_highlight.tsv"
    Dim p As String
    Dim adoStream As Object
    Dim txt As String
    Dim lines() As String
    Dim i As Long
    Dim ws As Worksheet
    Dim sheetName As String
    Dim numData As Long
    Dim clearCols() As String
    Dim headerMap As Object
    Dim r As Long
    Dim c As Long
    Dim colName As Variant
    Dim ci As Long
    Dim oneLine As String
    Dim parts() As String
    Dim cn As String
    Dim hv As Variant
    Dim lastRow As Long
    Dim lastCol As Long
    Dim prevSU As Boolean

    p = ThisWorkbook.path & "\log\" & SIDECAR
    If Len(Dir(p)) = 0 Then Exit Sub

    On Error GoTo CleanFail

    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.charset = "UTF-8"
    adoStream.Open
    adoStream.LoadFromFile p
    txt = adoStream.ReadText
    adoStream.Close
    Set adoStream = Nothing

    txt = Replace(txt, vbCrLf, vbLf)
    txt = Replace(txt, vbCr, vbLf)
    lines = Split(txt, vbLf)

    If UBound(lines) < 3 Then GoTo CleanDelete

    If Trim$(lines(0)) <> "V1" Then GoTo CleanDelete

    sheetName = Trim$(lines(1))
    numData = CLng(Val(Trim$(lines(2))))
    clearCols = Split(Trim$(lines(3)), vbTab)

    Set ws = ThisWorkbook.Sheets(sheetName)

    Set headerMap = CreateObject("Scripting.Dictionary")
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        hv = ws.Cells(1, c).Value
        If Not IsError(hv) Then
            If Not IsEmpty(hv) Then
                headerMap(Trim$(CStr(hv))) = c
            End If
        End If
    Next c

    lastRow = 1 + numData
    If lastRow < 2 Then lastRow = 2

    prevSU = Application.ScreenUpdating
    Application.ScreenUpdating = False

    ' 矛盾のないセルは段階1と同じ薄黄色へ。フォントは触らない（体裁維持）。
    ' AI解析列は黄色対象外（段階1の仕様に合わせる）
    For r = 2 To lastRow
        For Each colName In clearCols
            cn = Trim$(CStr(colName))
            If Len(cn) > 0 Then
                If headerMap.Exists(cn) Then
                    ci = headerMap(cn)
                    With ws.Cells(r, ci)
                        If StrComp(cn, "AI特別指定_解析", vbBinaryCompare) = 0 Then
                            .Interior.Pattern = xlNone
                        Else
                            .Interior.Color = RGB(255, 242, 204)
                        End If
                    End With
                End If
            End If
        Next colName
    Next r

    For i = 4 To UBound(lines)
        oneLine = Trim$(lines(i))
        If Len(oneLine) > 0 Then
            parts = Split(oneLine, vbTab, 2)
            If UBound(parts) >= 1 Then
                r = CLng(Val(parts(0)))
                cn = Trim$(parts(1))
                If r >= 2 Then
                    If headerMap.Exists(cn) Then
                        ci = headerMap(cn)
                        With ws.Cells(r, ci)
                            .Interior.Color = RGB(255, 0, 0)
                            .Font.Color = RGB(255, 255, 255)
                            .Font.Bold = True
                        End With
                    End If
                End If
            End If
        End If
    Next i

    Application.ScreenUpdating = prevSU
    Kill p
    Exit Sub

CleanDelete:
    On Error Resume Next
    Kill p
    Exit Sub

CleanFail:
    On Error Resume Next
    If Not adoStream Is Nothing Then
        adoStream.Close
        Set adoStream = Nothing
    End If
    Application.ScreenUpdating = True
End Sub

' production_plan 取り込み直後: Sheets(Count) だけだと末尾が _FontPick 等の固定シートになり、
' 実際にコピーされた「結果_*」ではなく _FontPick を参照してしまうことがある。名前優先で解決する。
Public Function 取込ブック内のコピー先シートを取得(ByVal wb As Workbook, ByVal expectedSheetName As String) As Worksheet
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
=======
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
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' production_plan 取り込み用: ベース名と同一、または Excel の番号付き複製
' （例: 「名前 (2)」「名前(2)」「名前 （2）」「名前（2）」等。先頭一致+括弧+数字のみを許可し「名前_別用途」と誤削除しない）
Public Function シート名は計画取込の同源名またはExcel番号付き複製か(ByVal nm As String, ByVal baseName As String) As Boolean
    Dim i As Long
    Dim j As Long
    Dim ch As String
    
    If StrComp(nm, baseName, vbBinaryCompare) = 0 Then
        シート名は計画取込の同源名またはExcel番号付き複製か = True
        Exit Function
    End If
=======
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
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' production_plan 取り込み: マクロブック側の「同名」および上記の番号付き複製をすべて削除してから Copy する。
Public Sub マクロブックから計画取込シート同源名シートを削除(ByVal wb As Workbook, ByVal sheetName As String)
=======
Private Sub マクロブックから計画取込シート同源名シートを削除(ByVal wb As Workbook, ByVal sheetName As String)
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' 段階2の取り込み結果をスプラッシュへ反映（成功・一部取込は MsgBox なし。未取得のみ警告ダイアログ）
Public Sub 段階2_取り込み結果を報告()
=======
Private Sub 段階2_取り込み結果を報告()
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' =========================================================
' 段階2コア: plan_simulation_stage2.py → 結果取り込み・メイン更新
' MsgBox は出さない。m_lastStage2ErrMsg / m_lastStage2ExitCode / m_stage2* を参照
' preserveStage1LogOnLogSheet=True … 段階1+2 連続時。LOG に段階1ログを残してから段階2を追記
' =========================================================
Public Sub 段階2_コア実行(Optional ByVal preserveStage1LogOnLogSheet As Boolean = False)
    Dim wsh As Object
    Dim runBat As String
    Dim outputText As String
    Dim targetDir As String
    Dim wsLog As Worksheet
    Dim logLines() As String
    Dim i As Long
    Dim adoStream As Object
    Dim logFilePath As String
    Dim cmdLogPath As String
    Dim outputFilePath As String
    Dim exitCode As Long
    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim sourceWb As Workbook
    Dim targetWb As Workbook
    Dim sourceWs As Worksheet
    Dim ws As Worksheet
    Dim sheetName As String
    Dim memberWb As Workbook
    Dim memberPath As String
    Dim newSheetName As String
    Dim planImported As Boolean
    Dim memberImported As Boolean
    Dim warnRow2 As Long
    Dim preserved As Collection
    Dim logStartRow As Long
    Dim logWriteRow As Long
    Dim lastLogR As Long
    Dim r As Long
    Dim pr As Long
    Dim st2XwErr As Long
    Dim st2XwDesc As String
    Dim missSt2 As String
    Dim st2DidUnlock As Boolean
    
    On Error GoTo ErrHandler
    
    m_lastStage2ErrMsg = ""
    m_lastStage2ExitCode = -1
    m_stage2PlanImported = False
    m_stage2MemberImported = False
    
    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    
    ' 1. 現在のExcelファイルの場所を取得 (UNCパス対応のため)
    targetDir = ThisWorkbook.path
    If targetDir = "" Then
        m_lastStage2ErrMsg = "先にこのExcelファイルを保存してください。"
        Exit Sub
    End If
    
    ' 2. 「LOG」シートが存在するか確認
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("LOG")
    On Error GoTo 0
    If wsLog Is Nothing Then
        m_lastStage2ErrMsg = "「LOG」シートが見つかりません。"
        Exit Sub
    End If

    Dim skErrSt2 As String
    If Not ValidateMasterSkillsOpAsPriorityUnique(targetDir, skErrSt2) Then
        m_lastStage2ErrMsg = skErrSt2
        Exit Sub
    End If

    設定_配台不要工程_シートを確保
    設定_環境変数_シートを確保
    設定_シート表示_シートを確保
    MacroSplash_SetStep "段階2: データ接続（Power Query 等）を更新しています…"

    If Not TryRefreshWorkbookQueries() Then
        m_lastStage2ErrMsg = "データ接続の更新に失敗したため段階2を中断しました。"
        If Len(m_lastRefreshQueriesErrMsg) > 0 Then
            m_lastStage2ErrMsg = m_lastStage2ErrMsg & vbCrLf & m_lastRefreshQueriesErrMsg
        End If
        Exit Sub
    End If
    
    MacroSplash_SetStep "段階2: LOG シートを準備しています（段階1ログの連結含む）…"
    Set preserved = New Collection
    If preserveStage1LogOnLogSheet Then
        lastLogR = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
        For r = 2 To lastLogR
            preserved.Add wsLog.Cells(r, 1).Value
        Next r
    End If
    
    ' ここでLOGシートは一旦クリア（連続実行時は直前に段階1行を退避済み）
    wsLog.Cells.Clear
    wsLog.Cells(1, 1).Value = "実行ブック: " & ThisWorkbook.FullName
    
    logStartRow = 2
    If preserveStage1LogOnLogSheet And preserved.Count > 0 Then
        wsLog.Cells(2, 1).Value = "---- 段階1（task_extract_stage1）----"
        logStartRow = 3
        For pr = 1 To preserved.Count
            wsLog.Cells(logStartRow, 1).Value = preserved(pr)
            logStartRow = logStartRow + 1
        Next pr
        wsLog.Cells(logStartRow, 1).Value = "---- 段階2（plan_simulation_stage2）----"
        logStartRow = logStartRow + 1
    End If
    
    Set wsh = CreateObject("WScript.Shell")
    
    If Not GeminiCredentialsJsonPathIsConfigured() Then
        wsLog.Cells(1, 2).Value = "【要設定】シート「設定」B1 に Gemini 認証 JSON のフルパス（例: Z:\社内\gemini_credentials.json）。gemini_credentials.example.json 参照。"
    End If

    ' タスク入力：TASK_INPUT_WORKBOOK でブックパスを渡す
    ' 段階2(plan_simulation_stage2.py) は「配台計画_タスク入力」シートを読みます
    ' 任意: シート「列設定_結果_タスク一覧」（列名・表示）で結果_タスク一覧の列順と表示/非表示を変更可。
    '       表示=FALSE の列は結果シートで列非表示。マクロ「列設定_結果_タスク一覧_チェックボックスを配置」でチェックボックスを表示列(B)に連動可能。
    wsh.Environment("Process")("TASK_INPUT_WORKBOOK") = ThisWorkbook.FullName
    
    ' ★修正：削除対象のログファイルのパスを output フォルダ配下に変更
    On Error Resume Next
    Kill targetDir & "\log\execution_log.txt"
    Kill targetDir & "\log\stage_vba_exitcode.txt"
    Kill targetDir & "\log\cmd_stage2.log"
    On Error GoTo 0
    
    ' ---------------------------------------------------------
    ' 【重要】UNCパス警告を回避する確実な方法
    ' ---------------------------------------------------------
    ' ① コマンドプロンプトが起動する瞬間の警告を防ぐため、裏で一時フォルダを指定
    wsh.CurrentDirectory = Environ("TEMP")
    
    ' 段階2: Python は TASK_INPUT_WORKBOOK のディスク上のファイルを読むため、
    ' 未保存の「配台計画_タスク入力」等が反映されないのを防ぐ
    MacroSplash_SetStep "段階2: ブックを保存しています…"
    Application.StatusBar = "ブックを保存しています..."
    DoEvents
    ThisWorkbook.Save
    Application.StatusBar = False
    
    st2DidUnlock = False
    配台マクロ_全シート保護を試行解除
    st2DidUnlock = True
    
    ' ② 「pushd」コマンドを使用し、UNCパスに一時的なドライブ文字を割り当てて確実に移動してからPythonを実行する
    ' リダイレクトは付けない（付けるとコンソールが真っ黒になる）。ログは Python が execution_log.txt にも出力する。
    cmdLogPath = targetDir & "\log\cmd_stage2.log"
    ' /v:on … py の終了コードを exit /b で返す（末尾の echo だけだと ERRORLEVEL=0 になり VBA が正常終了と誤認しがち）
    MacroSplash_SetStep "段階2: Python（plan_simulation）で計画シミュレーションを実行しています。完了までお待ちください…（詳細は LOG シート・log\execution_log.txt）"
    If STAGE12_USE_XLWINGS_RUNPYTHON And Not STAGE12_USE_XLWINGS_SPLASH_LOG Then
        wsh.Environment("Process")("PM_AI_SPLASH_XLWINGS") = ""
        m_splashExecutionLogPath = targetDir & "\log\execution_log.txt"
        m_stageVbaExitCodeLogDir = ""
        MacroSplash_ClearExecutionLogPane
        On Error Resume Next
        Err.Clear
        XwRunConsoleRunner "run_stage2_for_xlwings"
        If Err.Number <> 0 Then
            st2XwErr = Err.Number
            st2XwDesc = Err.Description
            Err.Clear
            On Error GoTo ErrHandler
            m_splashExecutionLogPath = ""
            m_stageVbaExitCodeLogDir = ""
            m_lastStage2ExitCode = -1
            m_lastStage2ErrMsg = "段階2: xlwings RunPython が失敗しました (" & CStr(st2XwErr) & "): " & st2XwDesc
            GoTo Finish
        End If
        On Error GoTo ErrHandler
        exitCode = ReadStageVbaExitCodeFromFile(targetDir & "\log\stage_vba_exitcode.txt")
        If exitCode = &H7FFFFFFF Then exitCode = 1
        m_splashExecutionLogPath = ""
        m_stageVbaExitCodeLogDir = ""
        m_lastStage2ExitCode = exitCode
        MacroSplash_LoadExecutionLogFromPath targetDir & "\log\execution_log.txt"
    Else
        Dim hideStage12CmdSt2 As Boolean
        hideStage12CmdSt2 = Stage12CmdHideWindowEffective()
        runBat = "@echo off" & vbCrLf & "setlocal EnableDelayedExpansion" & vbCrLf & "pushd """ & targetDir & """" & vbCrLf & _
                 "if not exist log mkdir log" & vbCrLf & _
                 "chcp 65001>nul" & vbCrLf & _
                 "echo [stage2] Running plan simulation... Progress below. Log file: log\execution_log.txt" & vbCrLf & _
                 "py -3 -u python\plan_simulation_stage2.py" & vbCrLf & _
                 "set STAGE2_PY_EXIT=!ERRORLEVEL!" & vbCrLf & _
                 "echo." & vbCrLf & _
                 "echo [stage2] Finished. ERRORLEVEL=!STAGE2_PY_EXIT!" & vbCrLf & _
                 "(echo !STAGE2_PY_EXIT!)>log\stage_vba_exitcode.txt" & vbCrLf
        ' コンソール表示時のみ: Python 失敗後にウィンドウがすぐ閉じないよう pause（非表示・headless では付けない）
        If Not hideStage12CmdSt2 Then
            runBat = runBat & "if not !STAGE2_PY_EXIT! equ 0 (" & vbCrLf & _
                     "echo." & vbCrLf & _
                     "echo [stage2] Python error. Press any key to close this window..." & vbCrLf & _
                     "pause" & vbCrLf & _
                     ")" & vbCrLf
        End If
        runBat = runBat & "exit /b !STAGE2_PY_EXIT!"
        ' 4. cmd 完了まで待機（execution_log を txtExecutionLog へポーリング）
        m_splashExecutionLogPath = targetDir & "\log\execution_log.txt"
        m_stageVbaExitCodeLogDir = ""
        MacroSplash_ClearExecutionLogPane
        exitCode = RunTempCmdWithConsoleLayout(wsh, runBat, Not hideStage12CmdSt2, hideStage12CmdSt2)
        m_splashExecutionLogPath = ""
        m_stageVbaExitCodeLogDir = ""
        m_lastStage2ExitCode = exitCode
    End If
    ' Python が検証エラー（例: exit 3）のとき log\stage2_blocking_message.txt に1行メッセージを残す。計画生成の MsgBox 用。
    If exitCode <> 0 Then
        Dim stage2Block As String
        stage2Block = Trim$(GeminiReadUtf8File(targetDir & "\log\stage2_blocking_message.txt"))
        If Len(stage2Block) > 0 Then
            m_lastStage2ErrMsg = stage2Block
        Else
            m_lastStage2ErrMsg = "Python の終了コードが " & CStr(exitCode) & " です。LOG シートおよび log\execution_log.txt を確認してください。（優先度重複などの検証中止時は log\stage2_blocking_message.txt も参照）"
        End If
    End If

    MacroSplash_SetStep "段階2: ログ・設定（配台不要工程・Gemini）をブックへ反映しています…"
    ' Gemini サマリ・設定シートは Python が openpyxl で保存を試みる。開きっぱなし時は log をマクロで反映。
    On Error Resume Next
    Call 設定_配台不要工程_AからE_TSVから反映
    Call 設定_配台不要工程_E列_TSVから反映
    Call メインシート_Gemini利用サマリをP列に反映(targetDir)
    On Error GoTo ErrHandler
    
    LOG_AIシートへ特別指定Geminiファイルを反映 targetDir
    
    ' 5. Python側で生成したログファイル(UTF-8)を読み込む
    logFilePath = targetDir & "\log\execution_log.txt"
    
    If Len(Dir(logFilePath)) = 0 Then
        wsLog.Cells(logStartRow, 1).Value = "execution_log.txt が見つかりませんでした。exitCode=" & CStr(exitCode)
        wsLog.Cells(logStartRow + 1, 1).Value = "xlwings 経路（STAGE12_USE_XLWINGS_RUNPYTHON=True）では Show Console の Python 出力も参照してください。runner は planning_core 読込前に log を作成するよう修正済みです。"
        missSt2 = Trim$(GeminiReadUtf8File(targetDir & "\log\stage2_blocking_message.txt"))
        If Len(missSt2) > 0 Then
            wsLog.Cells(logStartRow + 2, 1).Value = "log\stage2_blocking_message.txt: " & missSt2
        End If
        GoTo Finish
    End If
    
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.charset = "UTF-8"
    adoStream.Open
    
    adoStream.LoadFromFile logFilePath
    
    outputText = adoStream.ReadText
    adoStream.Close
    Set adoStream = Nothing
    
    ' 5. 改行コードを統一して配列に分割
    outputText = Replace(outputText, vbCrLf, vbLf)
    logLines = Split(outputText, vbLf)
    
    ' 6. LOGシートに一行ずつ書き出す（段階1退避があるときは logStartRow から）
    Application.ScreenUpdating = False
    logWriteRow = logStartRow
    For i = LBound(logLines) To UBound(logLines)
        wsLog.Cells(logWriteRow, 1).Value = logLines(i)
        logWriteRow = logWriteRow + 1
    Next i
    Application.ScreenUpdating = prevScreenUpdating
    
    If exitCode <> 0 Then
        warnRow2 = logWriteRow
        If warnRow2 < 1 Then warnRow2 = 2
        wsLog.Cells(warnRow2, 1).Value = "■ Pythonの終了コード: " & CStr(exitCode) & " （詳細は上記・実行時のコンソール・log\execution_log.txt を参照）"
    End If

    ' cmd.exe の標準出力/標準エラーも LOG シート末尾に追記（リダイレクトは環境により UTF-8 または Shift_JIS）
    If Len(Dir(cmdLogPath)) > 0 Then
        Dim cmdText As String
        Dim cmdLines() As String
        Dim baseRow As Long
        cmdText = ReadCmdCaptureLogText(cmdLogPath)
        cmdText = Replace(cmdText, vbCrLf, vbLf)
        cmdLines = Split(cmdText, vbLf)
        baseRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 2
        wsLog.Cells(baseRow, 1).Value = "---- cmd.exe stdout/stderr ----"
        For i = LBound(cmdLines) To UBound(cmdLines)
            wsLog.Cells(baseRow + 1 + i, 1).Value = cmdLines(i)
        Next i
    End If

    ' 計画検証エラー等で m_lastStage2ErrMsg が設定されたときは結果ブック取り込みをスキップ（誤って前回出力を取り込まない）
    If exitCode <> 0 And Len(m_lastStage2ErrMsg) > 0 Then
        GoTo Finish
    End If

    ' ブックが開いたままだと Python 側の openpyxl 保存が失敗することがある → TSV 経由でハイライトを反映
    ApplyPlanningConflictHighlightSidecar
    
    ' ---------------------------------------------------------
    ' 7. 生成されたExcelファイルのシートをこのブックに取り込む
    ' ---------------------------------------------------------
    MacroSplash_SetStep "段階2: 出力 xlsx から結果シート・個人シートを取り込みます…"
    planImported = False
    memberImported = False
    Set targetWb = ThisWorkbook
    
    ' 7a. production_plan_multi_day_*.xlsx（結果_* シート）
    outputFilePath = GetLatestOutputFile(targetDir & "\output", "production_plan_multi_day_*.xlsx")
    
    If outputFilePath <> "" Then
        ' 画面描画と警告を一時停止（削除確認ダイアログ等を非表示）
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        ' 列設定シートは削除せずに取り込むと同名で (2) が増えるため、事前に掃除
        列設定結果タスク一覧_番号付き重複シートを削除 targetWb
        
        ' 出力されたブックを開く
        Set sourceWb = Workbooks.Open(outputFilePath)
        
        For Each sourceWs In sourceWb.Sheets
            sheetName = Trim$(sourceWs.Name)
            
            ' Python 出力と同名のシートがマクロブックに残っていると、Copy 時に Excel が (2) を付けて複製する。
            ' 従来は「結果_*」と列設定のみ事前削除していたため、TEMP_設備毎の時間割・ブロックテーブル等が重複した。
            ' 既に残っている「名前 (2)」だけの場合もあるため、同源名（正確一致 + 「名前 (」で始まる複製）をまとめて削除する。
            マクロブックから計画取込シート同源名シートを削除 targetWb, sheetName
            Set ws = Nothing
            
            ' シートをコピー（ターゲットブックの末尾に）
            sourceWs.Copy After:=targetWb.Sheets(targetWb.Sheets.Count)
            
            ' コピーしたシートの書式設定（列幅、罫線、見出し）
            ' ※ Sheets(Count) だけだと末尾が _FontPick のとき誤参照するため、取り込み元と同名で引き直す
            Set ws = 取込ブック内のコピー先シートを取得(targetWb, sheetName)
            
            ' (1) セルフォントは上書きしない（Python 出力・ユーザーが「全シートフォント」で変更した体裁を段階2で維持する）
            
            ' (1b) 列幅: Python 出力では列幅を書かない。設備ガントは専用、それ以外は AutoFit。
            '     結果_タスク一覧 は非表示列があるため、全列 Select+AutoFit すると非表示が解除される。
            If StrComp(sheetName, "結果_設備ガント", vbBinaryCompare) = 0 Then
                結果_設備ガント_列幅を設定 ws
            ElseIf StrComp(sheetName, "結果_タスク一覧", vbBinaryCompare) = 0 Then
                結果シート_列幅_AutoFit非表示を維持 ws
                結果_タスク一覧_配完回答指定16時_いいえを強調 ws
            Else
                結果シート_列幅_AutoFit安定 ws
            End If
            
            ' (2) 使用している範囲全体に罫線(実線・細線)を引く
            ws.UsedRange.Borders.LineStyle = 1 ' xlContinuous
            ws.UsedRange.Borders.Weight = 2    ' xlThin
            ' 罫線付与で列幅が変わる環境があるため、設備ガントは専用幅を再適用
            If StrComp(sheetName, "結果_設備ガント", vbBinaryCompare) = 0 Then
                結果_設備ガント_列幅を設定 ws
            End If
            
            ' (3) 見出し（1行目）：太字・薄い黄緑（表形式シート向け）
            '     結果_設備ガント は 1 行目がレポートタイトル（Python でサイズ・背景を指定）のため上書きしない
            If StrComp(sheetName, "結果_設備ガント", vbBinaryCompare) <> 0 Then
                With ws.UsedRange.Rows(1)
                    .Font.Bold = True
                    .Interior.Color = RGB(226, 239, 218) ' 薄い黄緑色
                End With
            End If
            
            ' (3b) 結果_設備ガントのみ：タイトル A1（結合先頭）を強制的に左寄せ
            If StrComp(sheetName, "結果_設備ガント", vbBinaryCompare) = 0 Then
                結果_設備ガント_タイトルA1を左寄せに固定 ws
            End If
            
            ' (4) 結果_* のみ：メインシートへ戻るリンク（1行目・見出し行の右余白）
            If Left$(sheetName, 3) = "結果_" Then
                On Error Resume Next
                結果シート_メインへ戻るリンクを付与 ws
                Err.Clear
                On Error GoTo ErrHandler
            End If
            
            ' (5) 結果_* の保護は段階2 終了時（Finish）の 配台マクロ_対象シートを条件どおりに保護 でまとめて適用（処理中は全シート解除済み）
            
        Next sourceWs
        
        ' (6) master.xlsm メインの工場稼働(A12/B12)・定常(A15/B15)を結果_設備毎の時間割・結果_設備毎の時間割_機械名毎・結果_設備ガントに反映（UserInterfaceOnly 保護後もマクロから可。依頼NO薄緑は機械名毎のみ追加）
        On Error Resume Next
        取込後_結果シートへマスタ時刻を反映 targetWb
        Err.Clear
        ' マスタ反映・保護後も、設備ガントの列幅は専用設定に戻す（AutoFit 混入防止）
        Set ws = Nothing
        Set ws = targetWb.Worksheets("結果_設備ガント")
        If Err.Number = 0 Then
            結果_設備ガント_列幅を設定 ws
            結果_設備ガント_タイトルA1を左寄せに固定 ws
        End If
        Err.Clear
        On Error GoTo ErrHandler
        
        ' ソースブックを閉じる（保存しない）
        sourceWb.Close SaveChanges:=False
        Set sourceWb = Nothing
        
        planImported = True
        
        ' 画面描画と警告を元に戻す
        Application.DisplayAlerts = prevDisplayAlerts
        Application.ScreenUpdating = prevScreenUpdating
        
        ' 最初（一番左）のシートを選択状態にする（お好みで）
        targetWb.Sheets(1).Activate
    End If
    
    MacroSplash_SetStep "段階2: 個人別スケジュール（member_schedule）を取り込んでいます…"
    ' 7b. member_schedule_*.xlsx（メンバー名シート → 個人_プレフィックスで取り込み）
    memberPath = GetLatestOutputFile(targetDir & "\output", "member_schedule_*.xlsx")
    If Len(memberPath) > 0 Then
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        Set memberWb = Workbooks.Open(memberPath)
        
        For Each sourceWs In memberWb.Sheets
            sheetName = sourceWs.Name
            newSheetName = SafePersonalSheetName(sheetName)
            
            ' 既に「個人_*」シートがある場合は削除せず、内容をクリアしてから上書き
            On Error Resume Next
            Set ws = targetWb.Sheets(newSheetName)
            If Err.Number <> 0 Then
                Err.Clear
                On Error GoTo ErrHandler
                sourceWs.Copy After:=targetWb.Sheets(targetWb.Sheets.Count)
                Set ws = 取込ブック内のコピー先シートを取得(targetWb, sheetName)
                On Error Resume Next
                ws.Name = newSheetName
                On Error GoTo ErrHandler
            Else
                Err.Clear
                On Error GoTo ErrHandler
                ws.Cells.Clear
                sourceWs.UsedRange.Copy Destination:=ws.Range("A1")
            End If
            
            ' 個人_* もセルフォントは上書きしない（同上）
            ' 個人シートの列幅も Python 側では設定しない（同上 AutoFit）
            結果シート_列幅_AutoFit安定 ws
            ws.UsedRange.Borders.LineStyle = 1
            ws.UsedRange.Borders.Weight = 2
            With ws.UsedRange.Rows(1)
                .Font.Bold = True
                .Interior.Color = RGB(226, 239, 218)
            End With
        Next sourceWs
        
        memberWb.Close SaveChanges:=False
        Set memberWb = Nothing
        memberImported = True
        
        Application.DisplayAlerts = prevDisplayAlerts
        Application.ScreenUpdating = prevScreenUpdating
    End If
    
    MacroSplash_SetStep "段階2: メインシート・シート順・フォント後処理を実行しています…"
    ' メインシート：メンバーへのリンク ＋ 前日から12日間の出退勤（失敗しても本処理は継続）
    On Error Resume Next
    メインシート_メンバー一覧と出勤表示 True
    ' 個人_* シートをブック末尾へ（失敗しても継続）
    個人シートを末尾へ並べ替え
    ' 「設定」の一つ前に列設定シートを置く（取り込みでは末尾に付くため）
    On Error Resume Next
    列設定_結果_タスク一覧を設定の直前へ移動 ThisWorkbook
    On Error GoTo ErrHandler

    MacroSplash_SetStep "段階2: 「設定_シート表示」を一覧更新しブックへ適用しています…"
    On Error Resume Next
    設定_シート表示_一覧をブックから再取得
    Err.Clear
    設定_シート表示_ブックへ適用
    Err.Clear
    On Error GoTo ErrHandler
    
    ' 完了ダイアログ直前はメインシートを表示（A1）
    On Error Resume Next
    Application.ScreenUpdating = True
    メインシートA1を選択
    DoEvents
    On Error GoTo ErrHandler
    
    m_stage2PlanImported = planImported
    m_stage2MemberImported = memberImported

    On Error Resume Next
    配台_全シートフォントBIZ_UDP_自動適用
    On Error GoTo 0

Finish:
    On Error Resume Next
    If Not adoStream Is Nothing Then
        adoStream.Close
        Set adoStream = Nothing
    End If
    If Not sourceWb Is Nothing Then
        sourceWb.Close SaveChanges:=False
        Set sourceWb = Nothing
    End If
    If Not memberWb Is Nothing Then
        memberWb.Close SaveChanges:=False
        Set memberWb = Nothing
    End If
    On Error GoTo 0
    
    Application.DisplayAlerts = prevDisplayAlerts
    Application.ScreenUpdating = prevScreenUpdating
    
    If st2DidUnlock Then
        On Error Resume Next
        配台マクロ_対象シートを条件どおりに保護 targetDir
        On Error GoTo 0
    End If
    
    On Error Resume Next
    If planImported Then
        結果プレフィックスシートの表示倍率を設定 ThisWorkbook, 100
        結果_設備ガント_表示倍率を設定 ThisWorkbook, 85
        結果_設備毎の時間割_B2選択して窓枠固定
        結果_タスク一覧_F2選択して窓枠固定
        結果_カレンダー出勤簿_A2選択して窓枠固定
    End If
    メインシートA1を選択
    On Error GoTo 0
    
    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0
    
    Exit Sub
    
ErrHandler:
    m_lastStage2ErrMsg = "VBAエラー: " & Err.Number & " / " & Err.Description
    If Not wsLog Is Nothing Then
        wsLog.Cells(1, 1).Value = m_lastStage2ErrMsg
    End If
    Resume Finish
End Sub

' 互換・他モジュール用: 段階2のみ（エラー時 MsgBox。成功時はスプラッシュ＋チャイム）
Public Function SafePersonalSheetName(ByVal baseName As String) As String
=======
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
>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' =========================================================
' 【補助関数】最新の出力ファイルを取得する
' =========================================================
=======
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

>>>>>>> hosokawa/main2
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

<<<<<<< HEAD
' ブックを開いたときに Ctrl+Shift+テンキー - を登録（ThisWorkbook の BeforeClose で解除する例は 生産管理_AI配台テスト_ThisWorkbook_VBA.txt）
=======
Sub Auto_Open()
    ShortcutMainSheet_OnKeyRegister
End Sub

>>>>>>> hosokawa/main2
