Option Explicit

' =========================================================
' リポジトリの code\VBA\モジュール分割バージョン 配下の
' .bas / .cls / .frm を、ブックが code\ にある前提で一括インポートする。
' 事前に「ファイル」→「オプション」→「トラストセンター」→
' 「マクロの設定」→「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が必要。
' =========================================================

' Excel の「マクロ」(Alt+F8)・クイックアクセス・シート上のボタンから実行するときは
' 引数を取らない Sub だけが選べる。既定フォルダ取り込みは ImportVBAFiles_Default を使う。
' VBE 上で F5 する場合も、こちらにカーソルを置いて実行すると確実。
Public Sub ImportVBAFiles_Default()
    ImportVBAFiles
End Sub

' folderPath を空にすると GetSplitModulesFolderPath の結果を使う（コードや「イミディエイト」から呼ぶ用）。
Public Sub ImportVBAFiles(Optional ByVal folderPath As String = vbNullString)
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
            On Error Resume Next
            wb.VBProject.VBComponents.Import fileItem.Path

            If Err.Number = 0 Then
                importCount = importCount + 1
                Debug.Print fileItem.Name & " をインポートしました。"
            Else
                Debug.Print fileItem.Name & " のインポートに失敗しました（同名モジュールが存在する可能性があります）。 Err=" & Err.Number
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next fileItem

    If importCount > 0 Then
        MsgBox importCount & " 個のファイルをインポートしました。", vbInformation, "完了"
    Else
        MsgBox "インポート対象のファイル(.bas, .cls, .frm)が見つからないか、" & vbCrLf & _
               "すべて失敗しました。" & vbCrLf & resolvedPath, vbInformation, "完了"
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
