Option Explicit

Public Function GetOrCreateFontScratchSheet() As Worksheet
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

Public Sub RestoreCellFontProps(ByVal r As Range, ByVal oldName As String, _
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

Public Function FontNameExistsInExcel(ByVal fontName As String) As Boolean
    Dim i As Long
    For i = 1 To Application.FontNames.Count
        If StrComp(Application.FontNames(i), fontName, vbTextCompare) = 0 Then
            FontNameExistsInExcel = True
            Exit Function
        End If
    Next i
End Function

Public Sub ApplyFontToAllSheetCells(ByVal fontName As String, ByRef skippedOut As String)
    Dim ws As Worksheet
    Dim ur As Range
    Dim rangeErr As Boolean
    
    skippedOut = ""
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        rangeErr = False
        Set ur = Nothing
        Set ur = ws.UsedRange
        If Err.Number <> 0 Then
            skippedOut = skippedOut & "・" & ws.Name & "（UsedRange: " & Err.Description & "）" & vbCrLf
            Err.Clear
            rangeErr = True
        End If
        If Not rangeErr Then
            ur.Font.Name = fontName
            If Err.Number <> 0 Then
                skippedOut = skippedOut & "・" & ws.Name & "（Font.Name: " & Err.Description & "）" & vbCrLf
                Err.Clear
            End If
        End If
        On Error GoTo 0
    Next ws
End Sub

' 段階1・段階2のコア処理が成功した直後に呼ぶ。MsgBox なしで BIZ UDPゴシックを全シート UsedRange に付与しメインを AutoFit。
Public Sub 配台_全シートフォントBIZ_UDP_自動適用()
    Dim skipped As String
    On Error Resume Next
    ApplyFontToAllSheetCells BIZ_UDP_GOTHIC_FONT_NAME, skipped
    メインシート_AからK列_AutoFit
    結果_主要4結果シート_列オートフィット
    On Error GoTo 0
End Sub

' Excel 標準の［フォント］ダイアログで選んで全シートに適用
Public Sub 全シートフォントをリストから選択して統一()
    Dim wsScratch As Worksheet
    Dim r As Range
    Dim prevWs As Worksheet
    Dim prevVis As XlSheetVisibility
    Dim oldName As String
    Dim oldSize As Variant
    Dim oldBold As Variant
    Dim oldItalic As Variant
    Dim oldUnderline As Variant
    Dim oldColor As Variant
    Dim oldStrike As Variant
    Dim picked As String
    Dim skipped As String
    
    Set prevWs = ActiveSheet
    Set wsScratch = GetOrCreateFontScratchSheet()
    prevVis = wsScratch.Visible
    
    Set r = wsScratch.Range("A1")
    With r.Font
        oldName = .Name
        oldSize = .Size
        oldBold = .Bold
        oldItalic = .Italic
        oldUnderline = .Underline
        oldColor = .Color
        oldStrike = .Strikethrough
    End With
    
    wsScratch.Visible = xlSheetVisible
    wsScratch.Activate
    r.Select
    
    If Not Application.Dialogs(xlDialogFormatFont).Show Then
        RestoreCellFontProps r, oldName, oldSize, oldBold, oldItalic, oldUnderline, oldColor, oldStrike
        wsScratch.Visible = prevVis
        On Error Resume Next
        prevWs.Activate
        On Error GoTo 0
        Exit Sub
    End If
    
    picked = r.Font.Name
    RestoreCellFontProps r, oldName, oldSize, oldBold, oldItalic, oldUnderline, oldColor, oldStrike
    wsScratch.Visible = prevVis
    On Error Resume Next
    prevWs.Activate
    On Error GoTo 0
    
    配台マクロ_全シート保護を試行解除
    MacroSplash_SetStep "フォント「" & picked & "」を全シートへ適用しています…"
    On Error GoTo Fail
    ApplyFontToAllSheetCells picked, skipped
    On Error GoTo 0
    
    If Len(skipped) = 0 Then
        MacroSplash_SetStep "全シートのフォントを「" & picked & "」に設定しました。"
        m_animMacroSucceeded = True
    Else
        MacroSplash_SetStep "フォントは適用しましたが、一部シートをスキップしました（ダイアログで詳細を確認してください）。"
        MsgBox "フォント「" & picked & "」を設定しました。スキップしたシート:" & vbCrLf & vbCrLf & skipped, vbExclamation
        m_animMacroSucceeded = True
    End If
    On Error Resume Next
    メインシート_AからK列_AutoFit
    結果_主要4結果シート_列オートフィット
    配台マクロ_対象シートを条件どおりに保護
    On Error GoTo 0
    Exit Sub
    
Fail:
    On Error Resume Next
    配台マクロ_対象シートを条件どおりに保護
    On Error GoTo 0
    MsgBox "フォント設定でエラー: " & Err.Description, vbCritical
End Sub

Public Sub 全シートフォントを選択して統一()
    Call アニメ付き_全シートフォントをリストから選択して統一
End Sub

Public Sub 全シートフォントを手入力で統一()
    Dim v As Variant
    Dim fontName As String
    Dim skipped As String
    
    v = Application.InputBox( _
        "適用するフォント名を入力してください。" & vbCrLf & _
        "（ホームのフォントボックスと同じ表記）", _
        "全シートのフォント統一（手入力）", _
        BIZ_UDP_GOTHIC_FONT_NAME, _
        Type:=2)
    If VarType(v) = vbBoolean Then Exit Sub
    
    fontName = Trim$(CStr(v))
    If Len(fontName) = 0 Then
        MsgBox "フォント名が空のため中止しました。", vbExclamation
        Exit Sub
    End If
    
    If Not FontNameExistsInExcel(fontName) Then
        If MsgBox( _
            "フォント「" & fontName & "」が一覧に見つかりませんでした。" & vbCrLf & _
            "このまま適用を試みますか？", _
            vbQuestion Or vbYesNo, "確認") = vbNo Then
            Exit Sub
        End If
    End If
    
    配台マクロ_全シート保護を試行解除
    MacroSplash_SetStep "フォント「" & fontName & "」を全シートへ適用しています…"
    On Error GoTo FailHand
    ApplyFontToAllSheetCells fontName, skipped
    On Error GoTo 0
    
    If Len(skipped) = 0 Then
        MacroSplash_SetStep "全シートのフォントを「" & fontName & "」に設定しました。"
        m_animMacroSucceeded = True
    Else
        MacroSplash_SetStep "フォントは適用しましたが、一部シートをスキップしました（ダイアログで詳細を確認してください）。"
        MsgBox "フォント「" & fontName & "」を設定しました。スキップ:" & vbCrLf & vbCrLf & skipped, vbExclamation
        m_animMacroSucceeded = True
    End If
    On Error Resume Next
    メインシート_AからK列_AutoFit
    結果_主要4結果シート_列オートフィット
    配台マクロ_対象シートを条件どおりに保護
    On Error GoTo 0
    Exit Sub
    
FailHand:
    On Error Resume Next
    配台マクロ_対象シートを条件どおりに保護
    On Error GoTo 0
    MsgBox "フォント設定でエラー: " & Err.Description, vbCritical
End Sub

Public Sub 全シートフォント_BIZ_UDPゴシックに統一()
    Dim skipped As String
    MacroSplash_SetStep "全シートのフォントを「" & BIZ_UDP_GOTHIC_FONT_NAME & "」へ適用しています…"
    配台マクロ_全シート保護を試行解除
    On Error GoTo FailB
    ApplyFontToAllSheetCells BIZ_UDP_GOTHIC_FONT_NAME, skipped
    On Error GoTo 0
    
    If Len(skipped) = 0 Then
        MacroSplash_SetStep "全シートのフォントを「" & BIZ_UDP_GOTHIC_FONT_NAME & "」に設定しました。"
        m_animMacroSucceeded = True
    Else
        MacroSplash_SetStep "フォントは適用しましたが、一部シートをスキップしました（ダイアログで詳細を確認してください）。"
        MsgBox "フォントを設定しました。スキップ:" & vbCrLf & vbCrLf & skipped, vbExclamation
        m_animMacroSucceeded = True
    End If
    On Error Resume Next
    メインシート_AからK列_AutoFit
    結果_主要4結果シート_列オートフィット
    配台マクロ_対象シートを条件どおりに保護
    On Error GoTo 0
    Exit Sub
    
FailB:
    On Error Resume Next
    配台マクロ_対象シートを条件どおりに保護
    On Error GoTo 0
    MsgBox "フォント設定でエラー: " & Err.Description, vbCritical
End Sub

'==============================================================================
' 列設定_結果_タスク一覧 ? 表示列(B)と連動するフォームのチェックボックスを配置
' 開発タブ → 挿入 → フォーム コントロールの「チェックボックス」と同等。
' 再実行すると既存のチェックボックスを削除して付け直します（セルの TRUE/FALSE は保持）。
'==============================================================================
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

'==============================================================================
' 列設定_結果_タスク一覧: 「結果_タスク一覧」の1行目（列見出し）から A:B を再構成する（VBA のみ）。
' 見出し行は「列名」「表示」（planning_core と同じ）。各列の表示は True で初期化。
' 同一見出し（英大文字小文字無視）は1回だけ。空の見出しセルはスキップ。
' 既存のデータ行が新しい行数より下に残る場合は A:B の余白を ClearContents。
' 図形のマクロ: 「アニメ付き_列設定_結果_タスク一覧_結果シート見出しから再構成」。
' チェックボックスを B 列に付けている場合は、実行後に「列設定_結果_タスク一覧_チェックボックスを配置」を推奨。
'==============================================================================
Public Sub 列設定_結果_タスク一覧_結果シート見出しから再構成()
    Dim wsRes As Worksheet
    Dim wsCfg As Worksheet
    Dim lastCol As Long
    Dim c As Long
    Dim r As Long
    Dim hdr As String
    Dim prevScr As Boolean
    Dim oldLR As Long
    Dim seen As Object

    prevScr = Application.ScreenUpdating
    On Error GoTo FailRebuild

    On Error Resume Next
    Set wsRes = ThisWorkbook.Worksheets(SHEET_RESULT_TASK_LIST)
    Set wsCfg = ThisWorkbook.Worksheets(SHEET_COL_CONFIG_RESULT_TASK)
    On Error GoTo FailRebuild
    If wsRes Is Nothing Then
        MsgBox "シート「" & SHEET_RESULT_TASK_LIST & "」がありません。", vbExclamation, "列設定の再構成"
        Exit Sub
    End If
    If wsCfg Is Nothing Then
        MsgBox "シート「" & SHEET_COL_CONFIG_RESULT_TASK & "」がありません。", vbExclamation, "列設定の再構成"
        Exit Sub
    End If

    lastCol = wsRes.Cells(1, wsRes.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then
        MsgBox "「" & SHEET_RESULT_TASK_LIST & "」の1行目に列見出しがありません。", vbExclamation, "列設定の再構成"
        Exit Sub
    End If

    oldLR = wsCfg.Cells(wsCfg.Rows.Count, 1).End(xlUp).Row
    If oldLR < 1 Then oldLR = 1

    Set seen = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    seen.CompareMode = vbTextCompare
    On Error GoTo FailRebuild

    Application.ScreenUpdating = False

    wsCfg.Cells(1, 1).Value = "列名"
    wsCfg.Cells(1, 2).Value = "表示"

    r = 1
    For c = 1 To lastCol
        hdr = Trim$(CStr(wsRes.Cells(1, c).Value))
        If Len(hdr) = 0 Then GoTo NextHdrCol
        If seen.Exists(hdr) Then GoTo NextHdrCol
        seen.Add hdr, True
        r = r + 1
        wsCfg.Cells(r, 1).Value = hdr
        wsCfg.Cells(r, 2).Value = True
NextHdrCol:
    Next c

    If r < 2 Then
        Application.ScreenUpdating = prevScr
        MsgBox "有効な列見出しが見つかりませんでした。", vbExclamation, "列設定の再構成"
        Exit Sub
    End If

    If oldLR > r Then
        wsCfg.Range(wsCfg.Cells(r + 1, 1), wsCfg.Cells(oldLR, 2)).ClearContents
    End If

    Application.ScreenUpdating = prevScr
    m_animMacroSucceeded = True
    MsgBox "「" & SHEET_COL_CONFIG_RESULT_TASK & "」を " & CStr(r - 1) & " 列で更新しました。" & vbCrLf & _
           "チェックボックスを使っている場合は「列設定_結果_タスク一覧_チェックボックスを配置」を再実行してください。", vbInformation
    Exit Sub

FailRebuild:
    On Error Resume Next
    Application.ScreenUpdating = prevScr
    On Error GoTo 0
    MsgBox "列設定の再構成でエラー: " & Err.Description, vbCritical
End Sub

'==============================================================================
' 列設定_結果_タスク一覧 → 結果_タスク一覧 へ列順・列非表示を適用（Python / xlwings）
' 図形のマクロ: 「アニメ付き_列設定_結果_タスク一覧_列順表示をPython適用」（押下アニメ付き）。
'   重複列名の整理のみ: 「アニメ付き_列設定_結果_タスク一覧_重複列名を整理」。
'   本体を直指定すると AnimateButtonPush が動かない。
' ・事前に「列設定」シートで列名・表示を編集してから実行。
' ・Excel で本ブックを開いたまま（xlwings が接続）。保存してから実行推奨。
'==============================================================================
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
             "py -" & PM_AI_SETUP_PY_MINOR & " -u python\apply_result_task_column_layout.py" & vbCrLf & _
             "echo." & vbCrLf & _
             "echo [column-layout] ERRORLEVEL=%ERRORLEVEL%" & vbCrLf & _
             "exit /b %ERRORLEVEL%"
    exitCode = RunTempCmdWithConsoleLayout(wsh, runBat, False, Stage12CmdHideWindowEffective())
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

'==============================================================================
' 列設定_結果_タスク一覧: 列名の重複行を除き A:B を書き直す（Python）。結果_タスク一覧は変更しない。
' チェックボックスを B 列に付けている場合は、実行後に「列設定_結果_タスク一覧_チェックボックスを配置」を推奨。
'==============================================================================
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
             "py -" & PM_AI_SETUP_PY_MINOR & " -u python\dedupe_result_task_column_config_sheet.py" & vbCrLf & _
             "echo." & vbCrLf & _
             "echo [dedupe-column-config] ERRORLEVEL=%ERRORLEVEL%" & vbCrLf & _
             "exit /b %ERRORLEVEL%"
    exitCode = RunTempCmdWithConsoleLayout(wsh, runBat, False, Stage12CmdHideWindowEffective())
    Application.ScreenUpdating = prevScreen

    If exitCode <> 0 Then
        MsgBox "Python の終了コードが " & CStr(exitCode) & " です。" & vbCrLf _
            & "log\execution_log.txt を確認してください。", vbExclamation, "列設定の整理"
    Else
        MacroSplash_SetStep "「" & SHEET_COL_CONFIG_RESULT_TASK & "」の重複列名を除き A:B を更新しました。（チェックボックス利用時は配置マクロの再実行を推奨）"
        m_animMacroSucceeded = True
    End If
End Sub

'==============================================================================
' 配台計画_タスク入力: 「配台試行順番」を Python（xlwings）で再計算・行並べ替え
' 図形のマクロ: 「アニメ付き_配台計画_タスク入力_配台試行順番を再計算」
' 図形の自動作成: 「アニメ付き_配台計画_タスク入力_配台試行順再計算ボタンを配置」
' ・_apply_planning_sheet_post_load_mutations（設定シート行同期・分割行の自動配台不要）は従来どおり。
'   ただし「設定_配台不要工程」の C/E による計画シートへの配台不要の強制上書きは行わない（手動クリアを維持）。
' ・Excel で本ブックを開いたまま。保存してからの実行を推奨。
'==============================================================================
Public Sub 配台計画_タスク入力_配台試行順番をPythonで再計算()
    Dim wsh As Object
    Dim runBat As String
    Dim targetDir As String
    Dim exitCode As Long
    Dim wsPlan As Worksheet
    Dim prevScreen As Boolean

    targetDir = ThisWorkbook.path
    If Len(targetDir) = 0 Then
        MsgBox "先にこの Excel ファイルを保存してください。", vbExclamation, "配台試行順番の再計算"
        Exit Sub
    End If

    On Error Resume Next
    Set wsPlan = ThisWorkbook.Worksheets(SHEET_PLAN_INPUT_TASK)
    On Error GoTo 0
    If wsPlan Is Nothing Then
        MsgBox "シート「" & SHEET_PLAN_INPUT_TASK & "」がありません。", vbExclamation, "配台試行順番の再計算"
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    Set wsh = CreateObject("WScript.Shell")
    wsh.Environment("Process")("TASK_INPUT_WORKBOOK") = ThisWorkbook.FullName

    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    MacroSplash_SetStep "配台計画: Python で配台試行順番を再計算しています…"
    runBat = "@echo off" & vbCrLf & "pushd """ & targetDir & """" & vbCrLf & "chcp 65001>nul" & vbCrLf & _
             "py -" & PM_AI_SETUP_PY_MINOR & " -u python\apply_plan_input_dispatch_trial_order.py" & vbCrLf & _
             "echo." & vbCrLf & _
             "echo [plan-dispatch-trial-order] ERRORLEVEL=%ERRORLEVEL%" & vbCrLf & _
             "exit /b %ERRORLEVEL%"
    exitCode = RunTempCmdWithConsoleLayout(wsh, runBat, False, Stage12CmdHideWindowEffective())
    Application.ScreenUpdating = prevScreen

    If exitCode <> 0 Then
        MsgBox "Python の終了コードが " & CStr(exitCode) & " です。" & vbCrLf _
            & "log\execution_log.txt を確認してください。", vbExclamation, "配台試行順番の再計算"
    Else
        MacroSplash_SetStep "「" & SHEET_PLAN_INPUT_TASK & "」の配台試行順番を更新し、行を並べ替えました。"
        m_animMacroSucceeded = True
    End If
End Sub

'==============================================================================
' 配台計画_タスク入力: シートの「配台試行順番」を小数キーとして昇順に並べ替え 1..n（マスタ・上書き連携なし）
' 図形のマクロ: 「アニメ付き_配台計画_タスク入力_試行順を小数キーで並べ替え」
' 図形の自動作成: 「配台計画_タスク入力_試行順小数キー並べ替えボタンを配置」
'==============================================================================
Public Sub 配台計画_タスク入力_試行順を小数キーでPython並べ替え()
    Dim wsh As Object
    Dim runBat As String
    Dim targetDir As String
    Dim exitCode As Long
    Dim wsPlan As Worksheet
    Dim prevScreen As Boolean

    targetDir = ThisWorkbook.path
    If Len(targetDir) = 0 Then
        MsgBox "先にこの Excel ファイルを保存してください。", vbExclamation, "試行順の並べ替え"
        Exit Sub
    End If

    On Error Resume Next
    Set wsPlan = ThisWorkbook.Worksheets(SHEET_PLAN_INPUT_TASK)
    On Error GoTo 0
    If wsPlan Is Nothing Then
        MsgBox "シート「" & SHEET_PLAN_INPUT_TASK & "」がありません。", vbExclamation, "試行順の並べ替え"
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    Set wsh = CreateObject("WScript.Shell")
    wsh.Environment("Process")("TASK_INPUT_WORKBOOK") = ThisWorkbook.FullName

    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    MacroSplash_SetStep "配台計画: 配台試行順番を小数キーで並べ替えています…"
    runBat = "@echo off" & vbCrLf & "pushd """ & targetDir & """" & vbCrLf & "chcp 65001>nul" & vbCrLf & _
             "py -" & PM_AI_SETUP_PY_MINOR & " -u python\apply_plan_input_dispatch_trial_order_sort_by_float_keys.py" & vbCrLf & _
             "echo." & vbCrLf & _
             "echo [plan-dispatch-trial-float-keys] ERRORLEVEL=%ERRORLEVEL%" & vbCrLf & _
             "exit /b %ERRORLEVEL%"
    exitCode = RunTempCmdWithConsoleLayout(wsh, runBat, False, Stage12CmdHideWindowEffective())
    Application.ScreenUpdating = prevScreen

    If exitCode <> 0 Then
        MsgBox "Python の終了コードが " & CStr(exitCode) & " です。" & vbCrLf _
            & "log\execution_log.txt を確認してください。", vbExclamation, "試行順の並べ替え"
    Else
        MacroSplash_SetStep "「" & SHEET_PLAN_INPUT_TASK & "」の配台試行順番をキー順に並べ、1 から振り直しました。"
        m_animMacroSucceeded = True
    End If
End Sub

'==============================================================================
' 配台試行順: 複数パターン（P1～P3）の一覧シートを Python（xlwings）で作成
' 図形のマクロ: 「アニメ付き_配台計画_タスク入力_試行順パターン一覧シートを作成」
' 図形の自動作成: 「配台計画_タスク入力_試行順パターン一覧ボタンを配置」
' ・出力シート名は planning_core の DISPATCH_TRIAL_PATTERN_LIST_SHEET_NAME（既定「配台試行順_パターン一覧」）
' ・決定論パターンは P1～P3 を常に出力（一覧シートの列挙件数は段階2の DISPATCH_PATTERN_STAGE2_MAX_PATTERNS とは独立）
'==============================================================================
Public Sub 配台計画_タスク入力_試行順パターン一覧シートをPythonで作成()
    Dim wsh As Object
    Dim runBat As String
    Dim targetDir As String
    Dim exitCode As Long
    Dim wsPlan As Worksheet
    Dim prevScreen As Boolean

    targetDir = ThisWorkbook.path
    If Len(targetDir) = 0 Then
        MsgBox "先にこの Excel ファイルを保存してください。", vbExclamation, "試行順パターン一覧"
        Exit Sub
    End If

    On Error Resume Next
    Set wsPlan = ThisWorkbook.Worksheets(SHEET_PLAN_INPUT_TASK)
    On Error GoTo 0
    If wsPlan Is Nothing Then
        MsgBox "シート「" & SHEET_PLAN_INPUT_TASK & "」がありません。", vbExclamation, "試行順パターン一覧"
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    Set wsh = CreateObject("WScript.Shell")
    wsh.Environment("Process")("TASK_INPUT_WORKBOOK") = ThisWorkbook.FullName

    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    MacroSplash_SetStep "配台試行順: パターン一覧シートを作成しています…"
    runBat = "@echo off" & vbCrLf & "pushd """ & targetDir & """" & vbCrLf & "chcp 65001>nul" & vbCrLf & _
             "py -" & PM_AI_SETUP_PY_MINOR & " -u python\apply_dispatch_trial_pattern_list_sheet.py" & vbCrLf & _
             "echo." & vbCrLf & _
             "echo [dispatch-trial-pattern-list] ERRORLEVEL=%ERRORLEVEL%" & vbCrLf & _
             "exit /b %ERRORLEVEL%"
    exitCode = RunTempCmdWithConsoleLayout(wsh, runBat, False, Stage12CmdHideWindowEffective())
    Application.ScreenUpdating = prevScreen

    If exitCode <> 0 Then
        MsgBox "Python の終了コードが " & CStr(exitCode) & " です。" & vbCrLf _
            & "log\execution_log.txt を確認してください。", vbExclamation, "試行順パターン一覧"
    Else
        MacroSplash_SetStep "シート「" & SHEET_DISPATCH_TRIAL_PATTERN_LIST & "」を更新しました。"
        m_animMacroSucceeded = True
    End If
End Sub

'==============================================================================
' 配台試行順: 各パターン（P1～P4/R*）で段階2を実行し output に別ブック保存、サマリシートにリンクとスコア
' 図形のマクロ: 「アニメ付き_配台計画_タスク入力_試行順パターン別段階2を実行」
' 図形の自動作成: 「配台計画_タスク入力_試行順パターン別段階2ボタンを配置」
' ・サマリシート名は planning_core の DISPATCH_PATTERN_STAGE2_SUMMARY_SHEET_NAME（既定「配台試行順_パターン別段階2」）
' ・python\apply_dispatch_trial_pattern_stage2_batch.py（所要時間大）
'==============================================================================
Public Sub 配台計画_タスク入力_試行順パターン別段階2をPythonで作成()
    Dim wsh As Object
    Dim runBat As String
    Dim targetDir As String
    Dim exitCode As Long
    Dim wsPlan As Worksheet
    Dim prevScreen As Boolean

    targetDir = ThisWorkbook.path
    If Len(targetDir) = 0 Then
        MsgBox "先にこの Excel ファイルを保存してください。", vbExclamation, "パターン別段階2"
        Exit Sub
    End If

    On Error Resume Next
    Set wsPlan = ThisWorkbook.Worksheets(SHEET_PLAN_INPUT_TASK)
    On Error GoTo 0
    If wsPlan Is Nothing Then
        MsgBox "シート「" & SHEET_PLAN_INPUT_TASK & "」がありません。", vbExclamation, "パターン別段階2"
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    Set wsh = CreateObject("WScript.Shell")
    wsh.Environment("Process")("TASK_INPUT_WORKBOOK") = ThisWorkbook.FullName

    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    MacroSplash_SetStep "配台試行順: 各パターンで段階2を実行しています（完了までお待ちください）…"
    runBat = "@echo off" & vbCrLf & "pushd """ & targetDir & """" & vbCrLf & "chcp 65001>nul" & vbCrLf & _
             "py -" & PM_AI_SETUP_PY_MINOR & " -u python\apply_dispatch_trial_pattern_stage2_batch.py" & vbCrLf & _
             "echo." & vbCrLf & _
             "echo [dispatch-trial-pattern-stage2-batch] ERRORLEVEL=%ERRORLEVEL%" & vbCrLf & _
             "exit /b %ERRORLEVEL%"
    exitCode = RunTempCmdWithConsoleLayout(wsh, runBat, False, Stage12CmdHideWindowEffective())
    Application.ScreenUpdating = prevScreen

    If exitCode <> 0 Then
        MsgBox "Python の終了コードが " & CStr(exitCode) & " です。" & vbCrLf _
            & "log\execution_log.txt を確認してください。", vbExclamation, "パターン別段階2"
    Else
        MacroSplash_SetStep "シート「" & SHEET_DISPATCH_PATTERN_STAGE2_SUMMARY & "」を更新しました（output\dispatch_pattern_stage2 に結果ブック）。"
        m_animMacroSucceeded = True
    End If
End Sub

'==============================================================================
' 配台試行順: サマリで選んだパターンの試行順を「配台計画_タスク入力」へ反映（Python）
' 図形のマクロ: 「アニメ付き_配台計画_タスク入力_試行順パターン採用を実行」
' ・python\apply_dispatch_pattern_stage2_selection.py
'==============================================================================
Public Sub 配台計画_タスク入力_試行順パターン採用をPythonで実行()
    Dim wsh As Object
    Dim runBat As String
    Dim targetDir As String
    Dim exitCode As Long
    Dim wsPlan As Worksheet
    Dim prevScreen As Boolean

    targetDir = ThisWorkbook.path
    If Len(targetDir) = 0 Then
        MsgBox "先にこの Excel ファイルを保存してください。", vbExclamation, "パターン採用反映"
        Exit Sub
    End If

    On Error Resume Next
    Set wsPlan = ThisWorkbook.Worksheets(SHEET_PLAN_INPUT_TASK)
    On Error GoTo 0
    If wsPlan Is Nothing Then
        MsgBox "シート「" & SHEET_PLAN_INPUT_TASK & "」がありません。", vbExclamation, "パターン採用反映"
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    Set wsh = CreateObject("WScript.Shell")
    wsh.Environment("Process")("TASK_INPUT_WORKBOOK") = ThisWorkbook.FullName

    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    MacroSplash_SetStep "配台試行順: サマリで選んだパターンを計画シートへ反映しています…"
    runBat = "@echo off" & vbCrLf & "pushd """ & targetDir & """" & vbCrLf & "chcp 65001>nul" & vbCrLf & _
             "py -" & PM_AI_SETUP_PY_MINOR & " -u python\apply_dispatch_pattern_stage2_selection.py" & vbCrLf & _
             "echo." & vbCrLf & _
             "echo [dispatch-pattern-stage2-selection] ERRORLEVEL=%ERRORLEVEL%" & vbCrLf & _
             "exit /b %ERRORLEVEL%"
    exitCode = RunTempCmdWithConsoleLayout(wsh, runBat, False, Stage12CmdHideWindowEffective())
    Application.ScreenUpdating = prevScreen

    If exitCode <> 0 Then
        MsgBox "Python の終了コードが " & CStr(exitCode) & " です。" & vbCrLf _
            & "log\execution_log.txt を確認してください。", vbExclamation, "パターン採用反映"
    Else
        MacroSplash_SetStep "配台試行順: 採用パターンを「" & SHEET_PLAN_INPUT_TASK & "」に反映しました。"
        m_animMacroSucceeded = True
    End If
End Sub

' グラデーション＋影付き図形（メインの「かっこいいボタン」と同趣旨）。shapeName で図形名を区別する。
Private Sub PlanInputSheet_AddGradientActionButton( _
    ByVal ws As Worksheet, _
    ByVal btnText As String, _
    ByVal onActionFull As String, _
    ByVal leftPt As Single, _
    ByVal topPt As Single, _
    ByVal colorTop As Long, _
    ByVal colorBottom As Long, _
    ByVal shapeName As String)
    Dim shp As Shape
    Const BTN_W As Single = 268
    Const BTN_H As Single = 48
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPt, topPt, BTN_W, BTN_H)
    shp.Name = shapeName
    With shp
        With .TextFrame2.TextRange
            .text = btnText
            .Font.Name = "メイリオ"
            .Font.Size = 12
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
        .OnAction = onActionFull
    End With
End Sub

'==============================================================================
' 配台計画_タスク入力: 上記「配台試行順を再計算」用のグラデーション図形を 1 行目付近に配置
' 開発タブ → マクロ → 「配台計画_タスク入力_配台試行順再計算ボタンを配置」または
' 「アニメ付き_配台計画_タスク入力_配台試行順再計算ボタンを配置」
'==============================================================================
Public Sub 配台計画_タスク入力_配台試行順再計算ボタンを配置()
    Dim ws As Worksheet
    Dim ur As Range
    Dim anchorCol As Long
    Dim leftPt As Single
    Dim topPt As Single
    Dim sh As Shape
    Dim wbQuoted As String
    Dim macroAnim As String
    Dim i As Long
    On Error GoTo FailBtn
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_PLAN_INPUT_TASK)
    On Error GoTo FailBtn
    If ws Is Nothing Then
        MsgBox "シート「" & SHEET_PLAN_INPUT_TASK & "」がありません。", vbExclamation, "配台試行順ボタン"
        Exit Sub
    End If
    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        Set sh = ws.Shapes(i)
        If StrComp(sh.Name, SHAPE_PLAN_INPUT_DISPATCH_TRIAL_ORDER, vbTextCompare) = 0 Then
            sh.Delete
        End If
    Next i
    On Error GoTo FailBtn
    Set ur = Nothing
    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo FailBtn
    anchorCol = 4
    If Not ur Is Nothing Then
        anchorCol = ur.Column + ur.Columns.Count + 1
        If anchorCol < 4 Then anchorCol = 4
        If anchorCol > 80 Then anchorCol = 80
    End If
    leftPt = ws.Cells(1, anchorCol).Left
    topPt = ws.Rows(1).Top + 1.5
    wbQuoted = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'"
    macroAnim = wbQuoted & "!アニメ付き_配台計画_タスク入力_配台試行順番を再計算"
    PlanInputSheet_AddGradientActionButton ws, "配台試行順を更新", macroAnim, leftPt, topPt, RGB(100, 120, 220), RGB(40, 50, 120), SHAPE_PLAN_INPUT_DISPATCH_TRIAL_ORDER
    ws.Activate
    On Error Resume Next
    ws.Range("A1").Select
    On Error GoTo 0
    MsgBox "「" & SHEET_PLAN_INPUT_TASK & "」にボタンを配置しました。" & vbCrLf & _
           "（配台不要を手動でクリアしたあと、クリックで試行順を再計算します）", vbInformation, "配台試行順ボタン"
    Exit Sub
FailBtn:
    MsgBox "ボタン配置でエラー: " & Err.Description, vbCritical, "配台試行順ボタン"
End Sub

'==============================================================================
' 配台計画_タスク入力: 「小数キーで並べ替え→1..n」用グラデーション図形を 1 行目付近に配置（再計算ボタンの下）
' 開発タブ → マクロ → 「配台計画_タスク入力_試行順小数キー並べ替えボタンを配置」
'==============================================================================
Public Sub 配台計画_タスク入力_試行順小数キー並べ替えボタンを配置()
    Dim ws As Worksheet
    Dim ur As Range
    Dim anchorCol As Long
    Dim leftPt As Single
    Dim topPt As Single
    Dim sh As Shape
    Dim wbQuoted As String
    Dim macroAnim As String
    Dim i As Long
    Const BTN_H As Single = 48
    Const BTN_GAP As Single = 6
    On Error GoTo FailBtn2
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_PLAN_INPUT_TASK)
    On Error GoTo FailBtn2
    If ws Is Nothing Then
        MsgBox "シート「" & SHEET_PLAN_INPUT_TASK & "」がありません。", vbExclamation, "試行順キーボタン"
        Exit Sub
    End If
    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        Set sh = ws.Shapes(i)
        If StrComp(sh.Name, SHAPE_PLAN_INPUT_DISPATCH_TRIAL_ORDER_FLOAT_KEYS, vbTextCompare) = 0 Then
            sh.Delete
        End If
    Next i
    On Error GoTo FailBtn2
    Set ur = Nothing
    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo FailBtn2
    anchorCol = 4
    If Not ur Is Nothing Then
        anchorCol = ur.Column + ur.Columns.Count + 1
        If anchorCol < 4 Then anchorCol = 4
        If anchorCol > 80 Then anchorCol = 80
    End If
    leftPt = ws.Cells(1, anchorCol).Left
    topPt = ws.Rows(1).Top + 1.5 + BTN_H + BTN_GAP
    wbQuoted = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'"
    macroAnim = wbQuoted & "!アニメ付き_配台計画_タスク入力_試行順を小数キーで並べ替え"
    PlanInputSheet_AddGradientActionButton ws, "試行順をキーで並べ替え", macroAnim, leftPt, topPt, RGB(0, 150, 140), RGB(0, 75, 70), SHAPE_PLAN_INPUT_DISPATCH_TRIAL_ORDER_FLOAT_KEYS
    ws.Activate
    On Error Resume Next
    ws.Range("A1").Select
    On Error GoTo 0
    MsgBox "「" & SHEET_PLAN_INPUT_TASK & "」にボタンを配置しました。" & vbCrLf & _
           "（配台試行順番に 1, 2, 1.5 などキーを入れたあと、クリックで昇順に並べ 1 から振り直します）", vbInformation, "試行順キーボタン"
    Exit Sub
FailBtn2:
    MsgBox "ボタン配置でエラー: " & Err.Description, vbCritical, "試行順キーボタン"
End Sub

'==============================================================================
' 配台計画_タスク入力: 「試行順パターン一覧」用グラデーション図形（試行順キーボタンの下）
' 開発タブ → マクロ → 「配台計画_タスク入力_試行順パターン一覧ボタンを配置」
'==============================================================================
Public Sub 配台計画_タスク入力_試行順パターン一覧ボタンを配置()
    Dim ws As Worksheet
    Dim ur As Range
    Dim anchorCol As Long
    Dim leftPt As Single
    Dim topPt As Single
    Dim sh As Shape
    Dim wbQuoted As String
    Dim macroAnim As String
    Dim i As Long
    Const BTN_H As Single = 48
    Const BTN_GAP As Single = 6
    On Error GoTo FailBtn3
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_PLAN_INPUT_TASK)
    On Error GoTo FailBtn3
    If ws Is Nothing Then
        MsgBox "シート「" & SHEET_PLAN_INPUT_TASK & "」がありません。", vbExclamation, "パターン一覧ボタン"
        Exit Sub
    End If

    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        Set sh = ws.Shapes(i)
        If StrComp(sh.Name, SHAPE_PLAN_INPUT_DISPATCH_PATTERN_LIST, vbTextCompare) = 0 Then
            sh.Delete
        End If
    Next i
    On Error GoTo FailBtn3
    Set ur = Nothing
    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo FailBtn3
    anchorCol = 4
    If Not ur Is Nothing Then
        anchorCol = ur.Column + ur.Columns.Count + 1
        If anchorCol < 4 Then anchorCol = 4
        If anchorCol > 80 Then anchorCol = 80
    End If
    leftPt = ws.Cells(1, anchorCol).Left
    topPt = ws.Rows(1).Top + 1.5 + 2 * (BTN_H + BTN_GAP)
    wbQuoted = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'"
    macroAnim = wbQuoted & "!アニメ付き_配台計画_タスク入力_試行順パターン一覧シートを作成"
    PlanInputSheet_AddGradientActionButton ws, "試行順パターン一覧", macroAnim, leftPt, topPt, RGB(120, 90, 200), RGB(60, 40, 120), SHAPE_PLAN_INPUT_DISPATCH_PATTERN_LIST
    ws.Activate
    On Error Resume Next
    ws.Range("A1").Select
    On Error GoTo 0
    MsgBox "「" & SHEET_PLAN_INPUT_TASK & "」にボタンを配置しました。" & vbCrLf _
        & "クリックでシート「" & SHEET_DISPATCH_TRIAL_PATTERN_LIST & "」に各パターンの試行順一覧を書き込みます。", vbInformation, "パターン一覧ボタン"
    Exit Sub
FailBtn3:
    MsgBox "ボタン配置でエラー: " & Err.Description, vbCritical, "パターン一覧ボタン"
End Sub

'==============================================================================
' 配台計画_タスク入力: 「試行順パターン別段階2」用グラデーション図形（試行順パターン一覧の下）
' 開発タブ → マクロ → 「配台計画_タスク入力_試行順パターン別段階2ボタンを配置」
'==============================================================================
Public Sub 配台計画_タスク入力_試行順パターン別段階2ボタンを配置()
    Dim ws As Worksheet
    Dim ur As Range
    Dim anchorCol As Long
    Dim leftPt As Single
    Dim topPt As Single
    Dim sh As Shape
    Dim wbQuoted As String
    Dim macroAnim As String
    Dim i As Long
    Const BTN_H As Single = 48
    Const BTN_GAP As Single = 6
    On Error GoTo FailBtn4
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_PLAN_INPUT_TASK)
    On Error GoTo FailBtn4
    If ws Is Nothing Then
        MsgBox "シート「" & SHEET_PLAN_INPUT_TASK & "」がありません。", vbExclamation, "パターン別段階2ボタン"
        Exit Sub
    End If

    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        Set sh = ws.Shapes(i)
        If StrComp(sh.Name, SHAPE_PLAN_INPUT_DISPATCH_PATTERN_STAGE2, vbTextCompare) = 0 Then
            sh.Delete
        ElseIf StrComp(sh.Name, SHAPE_PLAN_INPUT_DISPATCH_PATTERN_STAGE2_SELECT, vbTextCompare) = 0 Then
            sh.Delete
        End If
    Next i
    On Error GoTo FailBtn4
    Set ur = Nothing
    On Error Resume Next
    Set ur = ws.UsedRange
    On Error GoTo FailBtn4
    anchorCol = 4
    If Not ur Is Nothing Then
        anchorCol = ur.Column + ur.Columns.Count + 1
        If anchorCol < 4 Then anchorCol = 4
        If anchorCol > 80 Then anchorCol = 80
    End If
    leftPt = ws.Cells(1, anchorCol).Left
    topPt = ws.Rows(1).Top + 1.5 + 3 * (BTN_H + BTN_GAP)
    wbQuoted = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'"
    macroAnim = wbQuoted & "!アニメ付き_配台計画_タスク入力_試行順パターン別段階2を実行"
    PlanInputSheet_AddGradientActionButton ws, "試行順パターン別段階2", macroAnim, leftPt, topPt, RGB(30, 110, 170), RGB(15, 55, 95), SHAPE_PLAN_INPUT_DISPATCH_PATTERN_STAGE2
    macroAnim = wbQuoted & "!アニメ付き_配台計画_タスク入力_試行順パターン採用を実行"
    PlanInputSheet_AddGradientActionButton ws, "試行順パターン採用を計画へ", macroAnim, leftPt, ws.Rows(1).Top + 1.5 + 4 * (BTN_H + BTN_GAP), RGB(20, 130, 90), RGB(10, 70, 50), SHAPE_PLAN_INPUT_DISPATCH_PATTERN_STAGE2_SELECT
    ws.Activate
    On Error Resume Next
    ws.Range("A1").Select
    On Error GoTo 0
    MsgBox "「" & SHEET_PLAN_INPUT_TASK & "」にボタンを配置しました。" & vbCrLf _
        & "上: 各パターンの段階2 → シート「" & SHEET_DISPATCH_PATTERN_STAGE2_SUMMARY & "」。" & vbCrLf _
        & "下: サマリ B3 で選んだパターンの試行順を計画シートへ反映。", vbInformation, "パターン別段階2ボタン"
    Exit Sub
FailBtn4:
    MsgBox "ボタン配置でエラー: " & Err.Description, vbCritical, "パターン別段階2ボタン"
End Sub

'==============================================================================
' シート「配台試行順_パターン一覧」に、一覧更新・段階2バッチ・採用反映のグラデ3ボタンを一括配置
' 開発タブ → マクロ → 「配台試行順_パターン一覧シートに試行順操作ボタン3つを配置」
' ・シートが無いときは先に「試行順パターン一覧」相当の処理でシートを作成するか、手動で同名シートを用意
'==============================================================================
Public Sub 配台試行順_パターン一覧シートに試行順操作ボタン3つを配置()
    Dim ws As Worksheet
    Dim wbQuoted As String
    Dim macroAnim As String
    Dim sh As Shape
    Dim i As Long
    Dim leftPt As Single
    Dim topPt As Single
    Dim cPlace As Long
    Const BTN_H As Single = 48
    Const BTN_GAP As Single = 6
    On Error GoTo FailBtn5
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_DISPATCH_TRIAL_PATTERN_LIST)
    On Error GoTo FailBtn5
    If ws Is Nothing Then
        MsgBox "シート「" & SHEET_DISPATCH_TRIAL_PATTERN_LIST & "」がありません。" & vbCrLf _
            & "先に「試行順パターン一覧」ボタン等で一覧を作成するか、同名シートを追加してください。", vbExclamation, "パターン一覧シート"
        Exit Sub
    End If

    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        Set sh = ws.Shapes(i)
        If StrComp(sh.Name, SHAPE_DISPATCH_TRIAL_PATTERN_LIST_SHEET_BTN_LIST, vbTextCompare) = 0 Then
            sh.Delete
        ElseIf StrComp(sh.Name, SHAPE_DISPATCH_TRIAL_PATTERN_LIST_SHEET_BTN_STAGE2, vbTextCompare) = 0 Then
            sh.Delete
        ElseIf StrComp(sh.Name, SHAPE_DISPATCH_TRIAL_PATTERN_LIST_SHEET_BTN_APPLY, vbTextCompare) = 0 Then
            sh.Delete
        End If
    Next i
    On Error GoTo FailBtn5

    cPlace = 2
    On Error Resume Next
    If Not ws.UsedRange Is Nothing Then
        cPlace = ws.UsedRange.Column + ws.UsedRange.Columns.Count + 1
    End If
    On Error GoTo FailBtn5
    If cPlace < 2 Then cPlace = 2
    If cPlace > 40 Then cPlace = 40
    leftPt = ws.Cells(1, cPlace).Left
    topPt = ws.Rows(1).Top + 6

    wbQuoted = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'"
    macroAnim = wbQuoted & "!アニメ付き_配台計画_タスク入力_試行順パターン一覧シートを作成"
    PlanInputSheet_AddGradientActionButton ws, "試行順パターン一覧", macroAnim, leftPt, topPt, RGB(120, 90, 200), RGB(60, 40, 120), SHAPE_DISPATCH_TRIAL_PATTERN_LIST_SHEET_BTN_LIST
    topPt = topPt + (BTN_H + BTN_GAP)
    macroAnim = wbQuoted & "!アニメ付き_配台計画_タスク入力_試行順パターン別段階2を実行"
    PlanInputSheet_AddGradientActionButton ws, "試行順パターン別段階2", macroAnim, leftPt, topPt, RGB(30, 110, 170), RGB(15, 55, 95), SHAPE_DISPATCH_TRIAL_PATTERN_LIST_SHEET_BTN_STAGE2
    topPt = topPt + (BTN_H + BTN_GAP)
    macroAnim = wbQuoted & "!アニメ付き_配台計画_タスク入力_試行順パターン採用を実行"
    PlanInputSheet_AddGradientActionButton ws, "試行順パターン採用を計画へ", macroAnim, leftPt, topPt, RGB(20, 130, 90), RGB(10, 70, 50), SHAPE_DISPATCH_TRIAL_PATTERN_LIST_SHEET_BTN_APPLY

    ws.Activate
    On Error Resume Next
    ws.Range("A1").Select
    On Error GoTo 0
    MsgBox "シート「" & SHEET_DISPATCH_TRIAL_PATTERN_LIST & "」にボタンを3つ配置しました。" & vbCrLf _
        & "上から: 一覧更新 → パターン別段階2 → 採用を計画へ反映。", vbInformation, "パターン一覧シート"
    Exit Sub
FailBtn5:
    MsgBox "ボタン配置でエラー: " & Err.Description, vbCritical, "パターン一覧シート"
End Sub

'==============================================================================
' デバッグ: ブックを開いたまま「どのシートが COM 的に触りにくいか」を一覧する
' 開発タブ → マクロ → 「COM操作テスト_全シートをログに出す」を実行。
' シート「COM操作テストログ」を末尾に作成し、シートごとの OK/NG を出します。
' （Excel 本体からの操作の目安。外部 Python/pywin32 の COM とはプロセスが異なる場合があります）
'==============================================================================


