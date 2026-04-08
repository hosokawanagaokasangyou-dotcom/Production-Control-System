Attribute VB_Name = "フォント管理"
Option Explicit

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

Private Sub 配台計画_タスク入力_既存シートの基準フォントを取得( _
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

Private Sub 配台計画_タスク入力_UsedRangeにフォント名とサイズを適用( _
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

Private Sub ApplyFontToAllSheetCells(ByVal fontName As String, ByRef skippedOut As String)
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

Private Sub 配台_全シートフォントBIZ_UDP_自動適用()
    Dim skipped As String
    On Error Resume Next
    ApplyFontToAllSheetCells BIZ_UDP_GOTHIC_FONT_NAME, skipped
    メインシート_AからK列_AutoFit
    結果_主要4結果シート_列オートフィット
    On Error GoTo 0
End Sub

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

