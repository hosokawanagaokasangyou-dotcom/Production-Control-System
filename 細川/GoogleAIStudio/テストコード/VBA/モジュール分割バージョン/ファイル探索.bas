Option Explicit

Function GetLatestOutputFile(folderPath As String, filePattern As String) As String
    Dim latestPath As String
    Dim latestDate As Date

    If Len(Dir(folderPath, vbDirectory)) = 0 Then
        最新の出力ファイルパスを取得 = ""
        Exit Function
    End If

    latestDate = 0
    latestPath = ""
    最新出力ファイルを再帰検索 folderPath, filePattern, latestPath, latestDate
    最新の出力ファイルパスを取得 = latestPath
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
        最新出力ファイルを再帰検索 CStr(subFldr.path), filePattern, latestPath, latestDate
    Next subFldr
End Sub

' =========================================================
' ファイル探索（最新の出力ファイルパスを返す）
' GetLatestOutputFile:
'   folderPath 配下を再帰的に走査し、filePattern（Like パターン）に一致するファイルのうち
'   DateLastModified が最大のもののフルパスを返す。該当なし・フォルダ不正時は "" 。
' CollectLatestOutputFileRecursive:
'   上記の内部処理。Scripting.FileSystemObject で Files / SubFolders を辿る。
' =========================================================
