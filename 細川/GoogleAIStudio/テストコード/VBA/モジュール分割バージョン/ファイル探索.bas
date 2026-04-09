Option Explicit

Function GetLatestOutputFile(folderPath As String, filePattern As String) As String
    Dim latestPath As String
    Dim latestDate As Date

    If Len(Dir(folderPath, vbDirectory)) = 0 Then
        GetLatestOutputFile = ""
        Exit Function
    End If

    latestDate = 0
    latestPath = ""
    CollectLatestOutputFileRecursive folderPath, filePattern, latestPath, latestDate
    GetLatestOutputFile = latestPath
End Function

Public Sub CollectLatestOutputFileRecursive(ByVal folderPath As String, ByVal filePattern As String, ByRef latestPath As String, ByRef latestDate As Date)
    Dim fso As Object
    Dim fldr As Object
    Dim subFldr As Object
    Dim fil As Object
    Dim currDate As Date
    If Len(folderPath) = 0 Then Exit Sub

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Sub
    If Not fso.FolderExists(folderPath) Then Exit Sub
    Set fldr = fso.GetFolder(folderPath)
    If fldr Is Nothing Then Exit Sub
    On Error GoTo 0

    For Each fil In fldr.Files
        If LCase$(fil.Name) Like LCase$(filePattern) Then
            On Error Resume Next
            currDate = fil.DateLastModified
            If Err.Number = 0 Then
                If currDate > latestDate Then
                    latestDate = currDate
                    latestPath = CStr(fil.path)
                End If
            Else
                Err.Clear
            End If
            On Error GoTo 0
        End If
    Next fil

    For Each subFldr In fldr.SubFolders
        CollectLatestOutputFileRecursive CStr(subFldr.path), filePattern, latestPath, latestDate
    Next subFldr
End Sub

' =========================================================
' 全シートのセルフォントを統一（手動実行・ボタン割当可）
' ※セルグリッドのみ。図形・グラフ内テキストは対象外。
' ※保護の解除・再保護は呼び出し元（段階1/2 コアまたは全シートフォント各マクロ）が 配台マクロ_* で実施。本サブは UsedRange のフォント名のみ変更。
' ※解除できないシートはスキップし、ダイアログに列挙。
' ※フォント後の列幅調整はメイン A:K と 結果_主要4結果シート_列オートフィット のみ。結果_設備ガントは専用列幅（オートフィットしない）。
' ※「リスト選択」は Excel 標準の［セルの書式設定］→［フォント］ダイアログを使用。
' ※図形のマクロには「アニメ付き_全シートフォントをリストから選択して統一」を指定（押下アニメ用。本体を直指定すると AnimateButtonPush が動かない）。
' =========================================================
