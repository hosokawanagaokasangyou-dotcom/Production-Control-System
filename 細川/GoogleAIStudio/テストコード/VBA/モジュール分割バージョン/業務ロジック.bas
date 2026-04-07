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
 =========================================================
 ★ 図形に登録するためのアニメーション付き起動マクロ ★
 処理本体は 段階1_コア実行 / 段階2_コア実行、ダイアログ付きの公開入口は RunPythonStage1 / RunPython / RunPythonStage1ThenStage2
 段階1・段階2のコアが成功で終わった直後、配台_全シートフォントBIZ_UDP_自動適用 で全シートを BIZ UDPゴシックに統一し、結果_主要4結果シート_列オートフィット で主要4結果シートの列幅を調整（完了の vbInformation MsgBox は使わずスプラッシュ＋システム音）
 段階2 Finish: 取り込み成功時は「結果_」で始まる全シートの表示倍率を 100% にし、その後 結果_設備ガント のみ 85% に戻す。結果_設備毎の時間割(B2)・結果_タスク一覧(F2)・結果_カレンダー(出勤簿)(A2) で窓枠固定を付与したうえで、最後にメインシート A1 をアクティブにして終了する
 =========================================================
Sub アニメ付き_計画生成を実行()
    Call AnimateButtonPush
    アニメ付き_スプラッシュ付きで実行 "シミュレーション（計画生成）を実行しています…", "RunPython", False, , True, True
End Sub
 段階1: 加工計画DATA からタスク抽出 → output に xlsx 出力し「配台計画_タスク入力」へ取り込み
Sub アニメ付き_タスク抽出を実行()
    Call AnimateButtonPush
    アニメ付き_スプラッシュ付きで実行 "タスク抽出（段階1）を実行しています…", "RunPythonStage1", , , True, True
End Sub
 段階1→保存反映→段階2を続けて実行（配台計画シートの手編集を挟まない一括実行）
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
 図形ボタン用：Caller が取れるのは「この Sub が OnAction のとき」だけ。本体を直接割り当てるとアニメは動かない。
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
 =========================================================
 Gemini API キーを暗号化 JSON にし「設定」B1 にパスを書く（押下アニメ付きはアニメ付き_* を図形に割当）
 暗号化パスフレーズは InputBox で入力し --passphrase-file 経由で Python に渡す。B2 にはパスフレーズを書かない。
 Python: python\encrypt_gemini_credentials.py（要 cryptography）。起動は py -3 を推奨。
 =========================================================
Sub アニメ付き_Gemini認証を暗号化してB1に保存()
    Call AnimateButtonPush
    ' InputBox 等があるためグリッド操作ブロックは使わない（スプラッシュのみ）
    アニメ付き_スプラッシュ付きで実行 "Gemini 認証を暗号化して保存しています…", "設定_Gemini認証を暗号化してB1に保存", , , False
End Sub
 列設定シートの内容を「結果_タスク一覧」へ反映（Python）。図形の OnAction には本マクロを指定（本体を直指定するとアニメは動かない）。
Sub アニメ付き_列設定_結果_タスク一覧_列順表示をPython適用()
    Call AnimateButtonPush
    アニメ付き_スプラッシュ付きで実行 "列設定を結果タスク一覧に反映しています…", "列設定_結果_タスク一覧_列順表示をPython適用"
End Sub
 列設定シート A:B のみ重複列名を削除（結果シートは触らない）。図形には「アニメ付き_列設定_結果_タスク一覧_重複列名を整理」。
Sub アニメ付き_列設定_結果_タスク一覧_重複列名を整理()
    Call AnimateButtonPush
    アニメ付き_スプラッシュ付きで実行 "列設定シートの重複列名を整理しています…", "列設定_結果_タスク一覧_重複列名を整理"
End Sub
Private Function GeminiJsonStringEscape(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCr, "\r")
    t = Replace(t, vbLf, "\n")
    t = Replace(t, vbTab, "\t")
    GeminiJsonStringEscape = t
End Function
Private Sub GeminiWriteUtf8File(ByVal filePath As String, ByVal textContent As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.charset = "UTF-8"
    stm.Open
    stm.WriteText textContent
    stm.SaveToFile filePath, 2
    stm.Close
    Set stm = Nothing
End Sub
 ログ表示用（暗号化失敗時の stderr など）
Private Function GeminiReadUtf8File(ByVal filePath As String) As String
    Dim stm As Object
    GeminiReadUtf8File = ""
    If Len(Dir(filePath)) = 0 Then Exit Function
    On Error GoTo CleanFail
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.charset = "UTF-8"
    stm.Open
    stm.LoadFromFile filePath
    GeminiReadUtf8File = stm.ReadText
    stm.Close
    Set stm = Nothing
    Exit Function
CleanFail:
    On Error Resume Next
    If Not stm Is Nothing Then stm.Close
    Set stm = Nothing
End Function
 Python が execution_log を開きっぱなしのとき LoadFromFile が共有違反で失敗することがある。一時コピーから読む。
Private Function GeminiReadUtf8FileViaTempCopy(ByVal filePath As String) As String
    Dim tmp As String
    GeminiReadUtf8FileViaTempCopy = ""
    If Len(Dir(filePath)) = 0 Then Exit Function
    Randomize
    tmp = Environ("TEMP") & "\pm_ai_sp_" & Replace(Replace(Replace(CStr(Now), "/", ""), ":", ""), " ", "_") & "_" & CStr(Int(100000 * Rnd)) & ".txt"
    On Error Resume Next
    FileCopy filePath, tmp
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    GeminiReadUtf8FileViaTempCopy = GeminiReadUtf8File(tmp)
    On Error Resume Next
    Kill tmp
End Function
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
 メインシートのリンク・出退勤のみ再反映したいとき（手動実行可）
Sub メインシート_メンバー一覧と出勤表示_手動()
    メインシート_メンバー一覧と出勤表示 False
End Sub
 同じフォルダの master.xlsm を開く（既に開いていればアクティブ化）
Public Sub メインシート_masterブックを開く()
    Dim path As String
    Dim folder As String
    Dim wb As Workbook
    Dim wbMaster As Workbook
    
    folder = ThisWorkbook.path
    If Len(folder) = 0 Then
        MsgBox "ブックを一度保存してから実行してください。", vbExclamation
        Exit Sub
    End If
    path = folder & "\" & MASTER_WORKBOOK_FILE
    If Len(Dir(path)) = 0 Then
        MsgBox "次のファイルが見つかりません。" & vbCrLf & path, vbExclamation
        Exit Sub
    End If
    
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, path, vbTextCompare) = 0 Then
            wb.Activate
            MacroSplash_SetStep "master.xlsm は既に開いています（アクティブにしました）。"
            m_animMacroSucceeded = True
            Exit Sub
        End If
    Next wb
    
    On Error GoTo OpenFail
    MacroSplash_SetStep "master.xlsm を開いています…"
    Set wbMaster = Application.Workbooks.Open(Filename:=path)
    wbMaster.Activate
    MacroSplash_SetStep "master.xlsm を開きました。"
    m_animMacroSucceeded = True
    Exit Sub
OpenFail:
    MsgBox "master.xlsm を開けませんでした: " & Err.Description, vbCritical
End Sub
Sub アニメ付き_メインシート_masterブックを開く()
    Call AnimateButtonPush
    メインシート_masterブックを開く
End Sub
 初回のみ推奨: メインシート上に「master.xlsm を開く」図形ボタンを1つ追加（重複したら不要分を削除）
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
 =========================================================
 メインシート A列上段：結果_* シートへのリンク
 B7～：個人シートへのリンク ＋ 前日から12日間の出退勤
 （結果_カレンダー(出勤簿) から取得。シート名は「メイン」「Main」または名前に「メイン」を含むもの）
 ★段階2(planning_core): 任意で見出しセルに「グローバルコメント」と書き、その直下のセルに「再優先特別記載」を入力可能。
   同文言は Gemini で解釈され、指示に応じてスキル無視・必要人数1名化などが通常ルールより最優先で適用される。
 ・勤怠セル: master.xlsm メイン A15/B15 の定常開始/終了と同じ「HH:MM / HH:MM」なら通常（背景なし）。読めないときは 08:45 / 17:00 基準。
 =========================================================
 master メイン A12/B12 のセル値を時刻として解釈（時分）。解釈不能は False。
 ※時刻のみのセルは Double になり IsDate が False になり得るため、数値型を明示処理する。
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
 planning_core.RESULT_OUTSIDE_REGULAR_TIME_FILL（FCE4D6）相当＝定常外
 工場稼働枠外（メイン A12/B12 の半開区間と重ならない帯）は薄い青で区別
Private Function 時刻を分に(ByVal t As Date) As Long
    時刻を分に = CLng(Hour(t)) * 60& + CLng(Minute(t))
End Function
 半開区間 [a0,a1) と [b0,b1) が重なるか（分単位・同一日内想定）
Private Function 半開区間が重なる分(ByVal a0 As Long, ByVal a1 As Long, ByVal b0 As Long, ByVal b1 As Long) As Boolean
    半開区間が重なる分 = (a0 < b1) And (a1 > b0)
End Function
 結果_設備毎の時間割「日時帯」セル（HH:MM-HH:MM 等）を解釈。■ を含む行は False
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
 master.xlsm 内のメイン設定シート（テストコード master_xlsm_VBA の MasterGetMainWorksheet と同趣旨）
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
 master メイン A12/B15 等: 結合セルでも左上の実値を取得
Private Function マスタメイン_結合左上の値(ByVal ws As Worksheet, ByVal cellAddr As String) As Variant
    Dim rng As Range
    On Error GoTo FailMMTL
    Set rng = ws.Range(cellAddr)
    マスタメイン_結合左上の値 = rng.MergeArea.Cells(1, 1).Value
    Exit Function
FailMMTL:
    マスタメイン_結合左上の値 = Empty
End Function
 master.xlsm メイン A12/B12（工場稼働）・A15/B15（定常）を読む。欠損・不正・開始>=終了は *Ok=False
Private Sub マスタメイン_工場稼働と定常を取得( _
    ByRef facOk As Boolean, ByRef facS As Date, ByRef facE As Date, _
    ByRef regOk As Boolean, ByRef regS As Date, ByRef regE As Date)
    
    Dim folder As String
    Dim p As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim openedHere As Boolean
    Dim v As Variant
    Dim tS As Date
    Dim tE As Date
    
    facOk = False
    regOk = False
    folder = ThisWorkbook.path
    If Len(folder) = 0 Then Exit Sub
    p = folder & "\" & MASTER_WORKBOOK_FILE
    If Len(Dir(p)) = 0 Then Exit Sub
    
    openedHere = False
    Set wb = Nothing
    On Error Resume Next
    Set wb = Workbooks(MASTER_WORKBOOK_FILE)
    On Error GoTo 0
    If wb Is Nothing Then
        On Error Resume Next
        Set wb = Workbooks.Open(Filename:=p, ReadOnly:=True, UpdateLinks:=0)
        openedHere = Not (wb Is Nothing)
        On Error GoTo 0
    End If
    If wb Is Nothing Then Exit Sub
    
    Set ws = マスタブック_メイン設定シートを取得(wb)
    If ws Is Nothing Then GoTo CloseMasterWb
    
    v = マスタメイン_結合左上の値(ws, "A12")
    If マスタメイン_セルを時刻Dateへ(v, tS) Then
        v = マスタメイン_結合左上の値(ws, "B12")
        If マスタメイン_セルを時刻Dateへ(v, tE) Then
            If TimeValue(tS) < TimeValue(tE) Then
                facOk = True
                facS = tS
                facE = tE
            End If
        End If
    End If
    
    v = マスタメイン_結合左上の値(ws, "A15")
    If マスタメイン_セルを時刻Dateへ(v, tS) Then
        v = マスタメイン_結合左上の値(ws, "B15")
        If マスタメイン_セルを時刻Dateへ(v, tE) Then
            If TimeValue(tS) < TimeValue(tE) Then
                regOk = True
                regS = tS
                regE = tE
            End If
        End If
    End If
CloseMasterWb:
    If openedHere Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
End Sub
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
 planning_core.RESULT_DISPATCHED_REQUEST_FILL（C6EFCE）相当＝機械名列の依頼NO
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
 結果_設備毎の時間割（および TEMP）: 設備セルに「(日次始業準備)」「(加工前準備)」「(依頼切替後始末)」が含まれるとき薄緑（進度列は除外）
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
 段階2: production_plan 取り込み直後に呼ぶ（当該マクロ内は終了時まで保護しない）
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
 手動: master を変更したあと結果シートだけ着色を合わせ直す（再取り込み不要）
Public Sub 結果シート_マスタ工場稼働と定常を再適用()
    取込後_結果シートへマスタ時刻を反映 ThisWorkbook
End Sub
 master.xlsm メイン A15/B15（定常）を「hh:nn / hh:nn」で返す（読めなければ 08:45 / 17:00）
Private Function マスタメイン_工場標準勤怠表示文字列() As String
    Const FB As String = "08:45 / 17:00"
    Dim folder As String
    Dim p As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim openedHere As Boolean
    Dim vS As Variant, vE As Variant
    Dim tS As Date, tE As Date
    
    マスタメイン_工場標準勤怠表示文字列 = FB
    On Error GoTo CleanExit
    
    folder = ThisWorkbook.path
    If Len(folder) = 0 Then Exit Function
    p = folder & "\" & MASTER_WORKBOOK_FILE
    If Len(Dir(p)) = 0 Then Exit Function
    
    openedHere = False
    Set wb = Nothing
    On Error Resume Next
    Set wb = Workbooks(MASTER_WORKBOOK_FILE)
    On Error GoTo CleanExit
    If wb Is Nothing Then
        On Error Resume Next
        Set wb = Workbooks.Open(Filename:=p, ReadOnly:=True, UpdateLinks:=0)
        On Error GoTo CleanExit
        openedHere = Not (wb Is Nothing)
    End If
    If wb Is Nothing Then Exit Function
    
    Set ws = マスタブック_メイン設定シートを取得(wb)
    If ws Is Nothing Then GoTo CloseWb
    
    vS = マスタメイン_結合左上の値(ws, "A15")
    vE = マスタメイン_結合左上の値(ws, "B15")
    If Not マスタメイン_セルを時刻Dateへ(vS, tS) Then GoTo CloseWb
    If Not マスタメイン_セルを時刻Dateへ(vE, tE) Then GoTo CloseWb
    If TimeValue(tS) >= TimeValue(tE) Then GoTo CloseWb
    
    マスタメイン_工場標準勤怠表示文字列 = Format$(tS, "hh:nn") & " / " & Format$(tE, "hh:nn")
CloseWb:
    If openedHere Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    Exit Function
CleanExit:
    On Error Resume Next
    If openedHere And Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Function
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
 結果_カレンダー(出勤簿) の「出勤 / 退勤」が master メイン A15/B15 の定常枠と一致するか
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
 メイン B7～（メンバー列＋日付12列 C～N）に表全体の細枠罫線を付与。ClearContents 後も B 列だけ線が無い状態を防ぐ。
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
 メインシートの A～N 列オートフィット（メインの勤怠12日分＋A列リンク。フォント変更後・段階2後のレイアウト用）
 ※ScreenUpdating=False 中は効かないことがあるため、必要なら True にしてから実行
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
 planning_core が log\gemini_usage_summary_for_main.txt（UTF-8）に出力した Gemini 利用サマリを、
 メインシート P16 以降に反映する。openpyxl が保存できない（ブック開きっぱなし）とき用。段階1/2 コアから呼ぶ。
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
 メインシートを表示し A1 をアクティブにする（シート名「メイン_」＝GetMainWorksheet と同じ）
Private Sub メインシートA1を選択()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = GetMainWorksheet()
    If ws Is Nothing Then Exit Sub
    ws.Activate
    ws.Range("A1").Select
    On Error GoTo 0
End Sub
' ShortcutMainSheet_CtrlShift0 / OnKeyRegister / OnKeyUnregister は 起動ショートカット.bas に集約
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
 ハイパーリンク再付与後に既定の青リンク体へ戻らないよう、クリア前のフォントを記憶・復元する
Private Sub メインシート_フォント属性を取得( _
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
Private Sub メインシート_フォント属性を適用( _
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
 範囲内の各セルに付いたハイパーリンクだけ削除（書式は維持）
Private Sub メインシート_指定範囲のハイパーリンクを削除(ByVal wsMain As Worksheet, ByVal Target As Range)
    Dim c As Range
    If wsMain Is Nothing Or Target Is Nothing Then Exit Sub
    On Error Resume Next
    For Each c In Target.Cells
        If c.Hyperlinks.Count > 0 Then c.Hyperlinks.Delete
    Next c
    On Error GoTo 0
End Sub
 メインシート A1～：ブック内で名前が「結果_」で始まるシートへのジャンプリンクを並べる
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
 結果_*（設備ガント以外）・個人_*: 実験コードと同じ手順で列オートフィット
 ・呼び出し元が ScreenUpdating=False のとき、Select 前に True に戻さないと AutoFit が効かないことがある
 ・元の ScreenUpdating は必ず復帰
 ・引数名は targetWs（RunPython 等の呼び出し側にも「ws」があり、ウォッチで親フレームの ws と混同しやすいため）
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
 結果_タスク一覧 専用: 非表示列に EntireColumn.AutoFit をかけると列が再表示されるため、表示列のみ AutoFit する。
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
 結果_タスク一覧: 列「配完_回答指定16時まで」（旧名「配完_基準16時まで」）が「いいえ」のセルを赤背景・白文字・太字にする。
 段階2の xlsx 取り込み直後に呼ぶ（openpyxl 側の書式に加え、列幅調整後の見た目を確実にする）。
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
 過去の取り込み不具合で残った「列設定_結果_タスク一覧 (2)」等を削除（本体シートは残す）
 ※呼び出し元が DisplayAlerts=False のとき、終了時に True に戻さない（シート削除確認が出るのを防ぐ）
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
 段階2 取り込み後: 「設定」タブの直前に列設定シートを置く（個人_*・LOG・設定の並べ替えの後に呼ぶ）
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
 結果_* シートの1行目右側に、メインシートへ戻る内部リンクを1つ置く（取り込み直後に呼ぶ）
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
 段階2 完了時: 結果_設備毎の時間割 を表示し B2 を選択、1行目と A 列を窓枠固定。シートが無ければ False。
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
 段階2 完了時: 結果_タスク一覧 を表示し F2 を選択、1 行目と A～E 列を窓枠固定。シートが無ければ False。
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
 段階2 完了時: 結果_カレンダー(出勤簿) を表示し A2 を選択、1 行目を窓枠固定。シートが無ければ False。
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
 段階2 完了間際: 名前が「結果_」で始まるシートの表示倍率を指定％にする（各シートを一度アクティブにして ActiveWindow.Zoom を設定）
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
 結果_設備ガントのみ表示倍率を設定（シートをアクティブにして ActiveWindow.Zoom）
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
 結果_設備ガント：取り込み直後に列幅を設定（Python 本体では列幅を書かない）
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
 結果_設備ガント：タイトルは結合セル先頭 A1。取り込み後の 1 行目一括書式のあと左寄せが崩れることがあるため固定する。
Private Sub 結果_設備ガント_タイトルA1を左寄せに固定(ByVal ws As Worksheet)
    On Error Resume Next
    With ws.Range("A1")
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
End Sub
 planning_core のガント罫線 thin color 666666 に合わせる（ハイライト解除時）
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
 Python が挿入する日付ブロック間の黒帯（行高さ約 5pt）
Private Function 結果_設備ガント_行は区切り行か(ByVal ws As Worksheet, ByVal r As Long) As Boolean
    Dim rh As Double
    On Error Resume Next
    rh = ws.Rows(r).RowHeight
    On Error GoTo 0
    If rh > 0# And rh <= 5.6 Then 結果_設備ガント_行は区切り行か = True
End Function
 タイトル・メタ（1?2 行）・表頭・区切り行以外＝計画行・実績行（日付は A 列縦結合のため列 A にも ? が現れる）
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
 ThisWorkbook.SheetSelectionChange から呼ぶ（標準モジュールは Public 必須）
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
 既存ブックで設備ガントが「セルの書式設定」無効の保護のとき、ハイライト罫線が効かない。パスワードは SHEET_FONT_UNPROTECT_PASSWORD のみ対応（手動パスワードはユーザーが一度解除してから実行）
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
 段階1/2 終盤・全シートフォント適用後: 結果の主要4シートの列オートフィット
 ・結果_タスク一覧 は非表示列を開かない（結果シート_列幅_AutoFit非表示を維持）
 ・結果_設備ガント は専用列幅（時刻列を潰さない）＋タイトル A1 左寄せ再固定
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
 =========================================================
 シート並び：個人_*（名前昇順）→ その後ろに LOG → 最後に「設定」
 （シート名は正確に LOG / 設定。無い場合はスキップ）
 =========================================================
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
 =========================================================
 シート並び：配台計画_タスク入力を前へ
 （step1完了時点で「個人_*」「LOG」「設定」の前の方へ配置）
 =========================================================
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
 =========================================================
 共通：ボタンを押し込むアニメーション処理
 ※ActiveSheet.Shapes(名前) だけだと、別シートに同じ図形名（既定の角丸1 等）があると
   誤ってそちらを動かし、意図しないシートが前面に出ることがあります。
   全シートから名前を解決し、ActiveSheet 上のものを優先します。
 =========================================================
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
 =========================================================
 マクロ実行中スプラッシュ（擬似モーダル）
 ・シート「設定」D3: true/TRUE でログ枠へ書き込み＋Exec 待機中のファイルポーリング。false/FALSE で無し・同期 Run・通常 cmd 表示（log\execution_log.txt への Python 出力は変わらず）
 ・シート「設定」D4: マクロ成功時の完了チャイム用 MP3 トラック番号 1?4（空・不正は 1）。ファイル名は標準モジュール MACRO_COMPLETE_MP3_1?4。sounds フォルダに配置。MP3 が無い／再生失敗時は macro_complete_chime.wav
 ・段階1／段階2のスプラッシュのみ: BGM（sounds 配下の Glass_Architecture1.mp3 等）を MCI ループ再生。終了時はフェードアウト後に close（完了チャイムより先）。他マクロのスプラッシュでは BGM・チャイムは再生しない
 ・UserForm「frmMacroSplash」をプロジェクトに追加（未追加時は表示せず続行）
 ・lockExcelUI=True のとき Application.Interactive=False でブック操作をブロック（対話マクロは False）
 ・ただし Interactive=False のままだと UserForm の再描画が滞り execution_log ポーリングが見えにくい。段階1/2 の Exec 待機中は一時的に True に戻す（RunCmdFileStageExecAndPoll）。
 ・終了・エラー時は必ず MacroSplash_Hide で Interactive を戻す
 ・作成手順とフォームコードは frmMacroSplash_VBA.txt
 ・完了の vbInformation MsgBox は原則やめ、段階1／段階2成功時はスプラッシュ最終文＋完了チャイム（MacroCompleteChime・設定 D4・sounds\*.mp3／WAV・失敗時 SystemAsterisk）
 =========================================================
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
Private Sub MacroSplash_SetStep(ByVal stepMessage As String)
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
 ログは末尾が最新。キャレットを最後に置き txtExecutionLog にフォーカス（UserForm には SetFocus がない）
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
 m_splashExecutionLogPath の UTF-8 ログを txtExecutionLog へ（長いときは末尾のみ）
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
    frmMacroSplash.lblMessage.Caption = "…（実行ログの表示に失敗 ? 下記の【ログ表示エラー】を参照）"
    frmMacroSplash.Repaint
    DoEvents
    If Not prevSU Then Application.ScreenUpdating = False
    m_splashReadErrShown = True
End Sub
 RunPython 終了直後など、同期 Python でポーリングできなかったあとに execution_log を一括表示（Interactive 一時 True で描画）
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
            errBanner = "【ログ表示エラー】execution_log.txt を読み込めませんでした。パス: " & fullPath & vbCrLf & vbCrLf
            prevInt = Application.Interactive
            Application.Interactive = True
            tb.text = errBanner & tb.text
            m_splashLastLogSnapshot = tb.text
            frmMacroSplash.lblMessage.Caption = "…（実行ログの一括表示に失敗 ? 下記を参照）"
            frmMacroSplash.Repaint
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
Private Function MacroCompleteChime_EnsureWavPath() As String
    Dim p As String
    Dim dirSounds As String
    p = MacroCompleteChime_LocalWavPath()
    If Len(p) = 0 Then Exit Function
    If Len(Dir(p)) > 0 Then
        MacroCompleteChime_EnsureWavPath = p
        Exit Function
    End If
    dirSounds = ThisWorkbook.path & "\" & MACRO_COMPLETE_CHIME_REL_DIR
    On Error Resume Next
    MkDir dirSounds
    On Error GoTo 0
    If MacroCompleteChime_HttpDownloadBinary(MACRO_COMPLETE_CHIME_DOWNLOAD_URL, p) Then
        If Len(Dir(p)) > 0 Then MacroCompleteChime_EnsureWavPath = p
    End If
End Function
Private Sub MacroCompleteChime()
    On Error Resume Next
    If Not m_splashAllowMacroSound Then Exit Sub
    Dim track As Long
    Dim mp3 As String
    Dim wav As String
    track = SettingsSheet_GetCompleteChimeTrack1to4()
    mp3 = MacroCompleteChime_LocalMp3Path(track)
    If Len(mp3) > 0 And Len(Dir(mp3)) > 0 Then
        If MacroCompleteChime_MciPlayMp3(mp3) Then Exit Sub
    End If
    wav = MacroCompleteChime_EnsureWavPath()
    If Len(wav) > 0 Then
        PlaySoundW StrPtr(wav), 0&, SND_FILENAME Or SND_ASYNC
    Else
        PlaySound "SystemAsterisk", 0&, SND_ALIAS Or SND_ASYNC
    End If
End Sub
' PlayFinishSound は サウンド制御.bas に集約（内部で MacroCompleteChime 相当を実行）
Private Function MacroStartBgm_FullPath() As String
    Dim folder As String
    folder = ThisWorkbook.path
    If Len(folder) = 0 Then Exit Function
    MacroStartBgm_FullPath = folder & "\" & MACRO_COMPLETE_CHIME_REL_DIR & "\" & MACRO_START_BGM_FILENAME
End Function
Private Sub MacroStartBgm_CloseHard()
    On Error Resume Next
    If m_macroStartBgmOpen Then
        mciSendStringW StrPtr("close " & MACRO_START_BGM_ALIAS), 0&, 0, 0&
    End If
    m_macroStartBgmOpen = False
End Sub
Private Sub MacroStartBgm_FadeOutAndClose()
    Dim i As Long
    Dim vol As Long
    On Error Resume Next
    If Not m_macroStartBgmOpen Then Exit Sub
    For i = 10 To 0 Step -1
        vol = 100& * i
        mciSendStringW StrPtr("setaudio " & MACRO_START_BGM_ALIAS & " volume to " & CStr(vol)), 0&, 0, 0&
        Sleep 45
        DoEvents
    Next i
    mciSendStringW StrPtr("close " & MACRO_START_BGM_ALIAS), 0&, 0, 0&
    m_macroStartBgmOpen = False
End Sub
Private Sub MacroStartBgm_StartIfAvailable()
    Dim p As String
    Dim r As Long
    Dim cmdOpen As String
    On Error Resume Next
    If Not m_splashAllowMacroSound Then Exit Sub
    p = MacroStartBgm_FullPath()
    If Len(p) = 0 Or Len(Dir(p)) = 0 Then Exit Sub
    MacroStartBgm_CloseHard
    cmdOpen = "open " & Chr$(34) & p & Chr$(34) & " type mpegvideo alias " & MACRO_START_BGM_ALIAS
    r = mciSendStringW(StrPtr(cmdOpen), 0&, 0, 0&)
    If r <> 0 Then Exit Sub
    mciSendStringW StrPtr("setaudio " & MACRO_START_BGM_ALIAS & " volume to 1000"), 0&, 0, 0&
    r = mciSendStringW(StrPtr("play " & MACRO_START_BGM_ALIAS & " repeat"), 0&, 0, 0&)
    If r <> 0 Then r = mciSendStringW(StrPtr("play " & MACRO_START_BGM_ALIAS), 0&, 0, 0&)
    If r = 0 Then m_macroStartBgmOpen = True
End Sub
 Excel メインウィンドウ（Application.hwnd）の下端・水平中央へ UserForm を SetWindowPos（モードレスのため API で座標指定）
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
 モードレス UserForm が背後に残ると再描画・ログ更新が止まったように見えることがある。Show 直後に前面へ（ユーザーがクリックしたときと同趣旨）。
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
        message = "処理中です。しばらくお待ちください。"
    End If
    frmMacroSplash.Caption = SPLASH_FORM_WINDOW_TITLE
    frmMacroSplash.lblMessage.Caption = message
    frmMacroSplash.StartUpPosition = 2  ' 初期のみ。直後に MacroSplash_PositionDockExcelBottomCenter で Excel 下端中央へ
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
' SplashLog_AppendChunk は スプラッシュ表示.bas に集約（xlwings からの呼び出しは PM_AI_XLWINGS_SPLASH_MACRO でモジュール名を指定可）
' アニメ付き_* から呼び出し：スプラッシュ表示 → マクロ実行（引数は最大2つまで Application.Run に委譲）
' lockExcelUI：False = InputBox／フォントダイアログなど Excel 対話が必要なマクロ向け
' allowMacroSound：True = 段階1／段階2と同様に BGM・成功時チャイムを許可（既定 False）
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
 =========================================================
 かっこいいボタンを自動生成するマクロ
 =========================================================
 グラデーション配色プリセット（CreateCoolButtonWithPreset の presetId）
 1=ロイヤルブルー 2=ティール 3=オレンジ 4=フォレストグリーン 5=パープル
 6=インディゴ 7=スレート 8=コーラル 9=アンバー 10=マゼンタ
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
 配色プリセット P1～P10 の見本を配置（マクロは割り当てず、見た目確認・色選び用）
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
 ボタン生成の共通ロジック
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
' InstallComponents および Python 環境セットアップ用の Private 補助は 環境セットアップ.bas に集約
' =========================================================
' Gemini 認証: Python は「設定」B1 の JSON ファイルパスからキーを読む（平文または暗号化）。
' 暗号化 JSON の復号は planning_core のソース内定数のみ。パスフレーズはシートに保存しない（B2 は未使用またはクリア）。
' =========================================================
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
 =========================================================
 外部データ／PQ の接続を「バックグラウンド更新しない」にそろえる。
 背景更新のまま RefreshAll が先に返り、その直後の Save 等で Excel が
 「この操作を実行すると、まだ実行されていないデータの更新が取り消されます」と出すのを防ぐ。
 =========================================================
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
 PQ 更新前: 接続先 IP へ ping 1 回（-w でタイムアウト）。成功時のみ True。
 失敗時はデータ更新をスキップし、呼び出し元は従来どおり True で継続する。
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
 =========================================================
 Power Query / データ接続の更新（マクロ処理の先頭で呼ぶ）
 ※ 先に DisableBackgroundDataRefreshAll で同期更新に寄せ、RefreshAll 後に
    CalculateUntilAsyncQueriesDone で取りこぼし待ち（背景オフ後はほぼ即時）。
    これにより「未実行のデータ更新が取り消されます」系ダイアログを抑止しやすくする。
 ※ DisplayAlerts=False で接続／PQ 失敗時の Excel 標準ダイアログを抑止。VBA 側も MsgBox は出さず
    m_lastRefreshQueriesErrMsg に詳細を残す（段階1・2のエラーメッセージに連結）。
 ※ PQ_REFRESH_PING_HOST へ ping（PQ_REFRESH_PING_TIMEOUT_MS）で応答がなければ RefreshAll は行わず、
    成功として返す（既存データのまま段階1・2を継続）。
 =========================================================
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
 Python の execution_log は UTF-8(BOM 付き)。cmd の 2>&1 リダイレクトは環境で Shift_JIS になりがちなので BOM で切り替える。
Private Function FileHasUtf8Bom(ByVal filePath As String) As Boolean
    Dim ff As Integer
    Dim b1 As Byte, b2 As Byte, b3 As Byte
    On Error GoTo CleanFail
    If Len(Dir(filePath)) = 0 Then FileHasUtf8Bom = False: Exit Function
    ff = FreeFile
    Open filePath For Binary Access Read As #ff
    Get #ff, 1, b1
    Get #ff, 2, b2
    Get #ff, 3, b3
    Close #ff
    FileHasUtf8Bom = (b1 = &HEF And b2 = &HBB And b3 = &HBF)
    Exit Function
CleanFail:
    On Error Resume Next
    If ff <> 0 Then Close #ff
    FileHasUtf8Bom = False
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
 cmd.exe が生成した capture ログ用（UTF-8 BOM が無ければ日本語環境では Shift_JIS として読む）
Private Function ReadCmdCaptureLogText(ByVal filePath As String) As String
    On Error GoTo EH
    If Len(Dir(filePath)) = 0 Then Exit Function
    If FileHasUtf8Bom(filePath) Then
        ReadCmdCaptureLogText = ReadTextFileWithCharset(filePath, "utf-8")
    Else
        ReadCmdCaptureLogText = ReadTextFileWithCharset(filePath, "Windows-932")
    End If
    Exit Function
EH:
    ReadCmdCaptureLogText = ""
End Function
 Excel で式として解釈される先頭 "=" を文字列として保持する
Private Function EscapeExcelFormulaText(ByVal s As String) As String
    If Len(s) > 0 Then
        If Left$(s, 1) = "=" Then
            EscapeExcelFormulaText = "'" & s
            Exit Function
        End If
    End If
    EscapeExcelFormulaText = s
End Function
 段階2 完了後: 特別指定_備考用 Gemini のプロンプト・応答ログを LOG_AI シートに転記（pause の代わりにブック内で確認）
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
 =========================================================
 設定_配台不要工程: シートの新規作成と見出し行のみ VBA（Python は工程+機械行の同期・AI・保存）
 手動で空シートだけ用意したい場合も本マクロを実行可。
 =========================================================
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
 =========================================================
 設定_環境変数: シートの新規作成・見出し行・テンプレにあってシートに無い変数名行のみ追記
 （python/workbook_env_bootstrap.py・設定_環境変数_雛形.tsv と整合）
 =========================================================
Private Function 設定_環境変数_1行目は見出し(ByVal ws As Worksheet) As Boolean
    Dim t As String
    t = LCase$(Trim$(CStr(ws.Cells(1, 1).Value)))
    If Len(t) = 0 Then
        設定_環境変数_1行目は見出し = False
        Exit Function
