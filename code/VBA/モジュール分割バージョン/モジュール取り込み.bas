Option Explicit

' =========================================================
' リポジトリの code\VBA\モジュール分割バージョン 配下の
' .bas / .cls / .frm を、ブックが code\ にある前提で一括インポートする。
' 既定ではファイル名（拡張子除く）をモジュール名に合わせる。
' 事前に「ファイル」→「オプション」→「トラストセンター」→
' 「マクロの設定」→「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が必要。
' =========================================================

' 同名モジュールが既にあるときの扱い（ImportVBAFiles の dupPolicy）
Public Enum PmDupImportPolicy
    pmDupSkip = 0      ' 同名があればそのファイルはインポートしない（既存を維持）
    pmDupReplace = 1   ' 既存の同名コンポーネントを削除してから取り込む（中身を差し替え）
    pmDupSuffix = 2    ' 既存があれば「名前_2」「名前_3」のように連番で別名にする
End Enum

' Excel の「マクロ」(Alt+F8)・クイックアクセス・シート上のボタンから実行するときは
' 引数を取らない Sub だけが選べる。既定フォルダ取り込みは ImportVBAFiles_Default を使う。
' VBE 上で F5 する場合も、こちらにカーソルを置いて実行すると確実。
' 既定は pmDupSuffix（既存を消さない）。
Public Sub ImportVBAFiles_Default()
    ImportVBAFiles vbNullString, pmDupSuffix
End Sub

' リポジトリと完全に揃えたいとき：同名は既存モジュールを削除してから取り込む。
Public Sub ImportVBAFiles_Default_Replace()
    ImportVBAFiles vbNullString, pmDupReplace
End Sub

' folderPath を空にすると GetSplitModulesFolderPath の結果を使う（コードや「イミディエイト」から呼ぶ用）。
' dupPolicy を省略時は pmDupSuffix。
Public Sub ImportVBAFiles( _
    Optional ByVal folderPath As String = vbNullString, _
    Optional ByVal dupPolicy As PmDupImportPolicy = pmDupSuffix)

    Dim fso As Object
    Dim targetFolder As Object
    Dim fileItem As Object
    Dim extension As String
    Dim wb As Workbook
    Dim importCount As Long
    Dim resolvedPath As String

    Set wb = ThisWorkbook

    If Len(Trim$(folderPath)) = 0 Then
        resolvedPath = GetSplitModulesFolderPath()
    Else
        resolvedPath = Trim$(folderPath)
    End If

    If Len(resolvedPath) = 0 Then
        MsgBox "フォルダパスを解決できません。" & vbCrLf & _
               "ブックを一度保存してから実行するか、folderPath を明示してください。", vbExclamation, "エラー"
        Exit Sub
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(resolvedPath) Then
        MsgBox "指定されたフォルダが見つかりません。" & vbCrLf & resolvedPath, vbExclamation, "エラー"
        Exit Sub
    End If

    Set targetFolder = fso.GetFolder(resolvedPath)
    importCount = 0

    For Each fileItem In targetFolder.Files
        extension = LCase$(fso.GetExtensionName(fileItem.Name))

        If extension = "bas" Or extension = "cls" Or extension = "frm" Then
            importCount = importCount + ImportOneVBFile(wb, fileItem.Path, fileItem.Name, fso, dupPolicy)
        End If
    Next fileItem

    If importCount > 0 Then
        MsgBox importCount & " 個のファイルをインポートしました。", vbInformation, "完了"
    Else
        MsgBox "インポート対象のファイル(.bas, .cls, .frm)が見つからないか、" & vbCrLf & _
               "すべてスキップ／失敗しました。" & vbCrLf & resolvedPath, vbInformation, "完了"
    End If

    Set fileItem = Nothing
    Set targetFolder = Nothing
    Set fso = Nothing
End Sub

' ブックが code\ 直下にあるとき、ThisWorkbook.Path\VBA\モジュール分割バージョン を返す。
' 未保存ブックなどで Path が取れないときは "" 。
Public Function GetSplitModulesFolderPath() As String
    Dim basePath As String

    basePath = Trim$(ThisWorkbook.Path)
    If Len(basePath) = 0 Then
        GetSplitModulesFolderPath = vbNullString
        Exit Function
    End If

    If Right$(basePath, 1) <> Application.PathSeparator Then
        basePath = basePath & Application.PathSeparator
    End If

    GetSplitModulesFolderPath = basePath & "VBA" & Application.PathSeparator & "モジュール分割バージョン"
End Function

' -------------------------------------------------------------------------
' 以下 Private
' -------------------------------------------------------------------------

Private Function ImportOneVBFile( _
    ByVal wb As Workbook, _
    ByVal filePath As String, _
    ByVal fileNameOnly As String, _
    ByVal fso As Object, _
    ByVal dupPolicy As PmDupImportPolicy) As Long

    Dim vbp As Object
    Dim desired As String
    Dim comp As Object
    Dim finalName As String

    ImportOneVBFile = 0
    Set vbp = wb.VBProject

    desired = SanitizeVBModuleName(fso.GetBaseName(fileNameOnly))
    If Len(desired) = 0 Then desired = "ImportedModule"

    Select Case dupPolicy
        Case pmDupSkip
            If VBComponentExists(vbp, desired) Then
                Debug.Print fileNameOnly & " をスキップしました（既に '" & desired & "' があります）。"
                Exit Function
            End If
            If Not DoImport(vbp, filePath, comp) Then Exit Function
            If Not TryRenameComponent(comp, desired) Then
                Debug.Print fileNameOnly & " のリネームに失敗しました。 Err=" & Err.Number
            End If

        Case pmDupReplace
            If VBComponentExists(vbp, desired) Then
                If Not SafeRemoveComponent(vbp, desired) Then
                    Debug.Print fileNameOnly & " をスキップしました（'" & desired & "' を削除できません）。"
                    Exit Function
                End If
            End If
            If Not DoImport(vbp, filePath, comp) Then Exit Function
            If Not TryRenameComponent(comp, desired) Then
                Debug.Print fileNameOnly & " のリネームに失敗しました。 Err=" & Err.Number
            End If

        Case pmDupSuffix
            finalName = NextAvailableModuleName(vbp, desired)
            If Not DoImport(vbp, filePath, comp) Then Exit Function
            If StrComp(comp.Name, finalName, vbTextCompare) <> 0 Then
                If Not TryRenameComponent(comp, finalName) Then
                    Debug.Print fileNameOnly & " のリネームに失敗しました。 Err=" & Err.Number
                End If
            End If

        Case Else
            If Not DoImport(vbp, filePath, comp) Then Exit Function
    End Select

    Debug.Print fileNameOnly & " をインポートしました（モジュール名: " & comp.Name & "）。"
    ImportOneVBFile = 1
End Function

Private Function DoImport(ByVal vbp As Object, ByVal filePath As String, ByRef comp As Object) As Boolean
    On Error Resume Next
    Set comp = vbp.VBComponents.Import(filePath)
    If Err.Number <> 0 Then
        Debug.Print filePath & " の Import に失敗しました。 Err=" & Err.Number & " " & Err.Description
        Err.Clear
        Set comp = Nothing
        DoImport = False
    Else
        DoImport = Not comp Is Nothing
    End If
    On Error GoTo 0
End Function

Private Function TryRenameComponent(ByVal comp As Object, ByVal newName As String) As Boolean
    On Error Resume Next
    comp.Name = newName
    TryRenameComponent = (Err.Number = 0)
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
End Function

Private Function VBComponentExists(ByVal vbp As Object, ByVal compName As String) As Boolean
    Dim test As Object
    On Error Resume Next
    Set test = vbp.VBComponents(compName)
    VBComponentExists = Not test Is Nothing
    Set test = Nothing
    Err.Clear
    On Error GoTo 0
End Function

Private Function SafeRemoveComponent(ByVal vbp As Object, ByVal compName As String) As Boolean
    Dim c As Object
    On Error Resume Next
    Set c = vbp.VBComponents(compName)
    If c Is Nothing Then
        SafeRemoveComponent = True
        Exit Function
    End If
    vbp.VBComponents.Remove c
    SafeRemoveComponent = (Err.Number = 0)
    If Err.Number <> 0 Then Err.Clear
    Set c = Nothing
    On Error GoTo 0
End Function

' VBA のモジュール名は最大 31 文字。禁則文字は除去・置換する。
Private Function SanitizeVBModuleName(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    t = Replace(t, "\", "_")
    t = Replace(t, "/", "_")
    t = Replace(t, ":", "_")
    t = Replace(t, "*", "_")
    t = Replace(t, "?", "_")
    t = Replace(t, """", "_")
    t = Replace(t, "<", "_")
    t = Replace(t, ">", "_")
    t = Replace(t, "|", "_")
    t = Replace(t, "[", "_")
    t = Replace(t, "]", "_")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    t = Replace(t, " ", "_")
    If Len(t) > 31 Then t = Left$(t, 31)
    If Len(t) = 0 Then t = "ImportedModule"
    SanitizeVBModuleName = t
End Function

' baseName が空いていればそれを返す。取られていれば baseName_2, baseName_3 …（31文字以内に収める）。
Private Function NextAvailableModuleName(ByVal vbp As Object, ByVal baseName As String) As String
    Dim n As Long
    Dim suffix As String
    Dim candidate As String
    Dim maxBase As Long

    If Not VBComponentExists(vbp, baseName) Then
        NextAvailableModuleName = baseName
        Exit Function
    End If

    n = 2
    Do
        suffix = "_" & CStr(n)
        maxBase = 31 - Len(suffix)
        If maxBase < 1 Then maxBase = 1
        candidate = Left$(baseName, maxBase) & suffix
        If Not VBComponentExists(vbp, candidate) Then
            NextAvailableModuleName = candidate
            Exit Function
        End If
        n = n + 1
        If n > 9999 Then Exit Do
    Loop

    NextAvailableModuleName = Left$(baseName, 20) & "_" & CStr(Timer)
End Function
