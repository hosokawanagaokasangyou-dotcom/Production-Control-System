Attribute VB_Name = "ï∂éöóÒì¸èoóÕã§í "
Option Explicit

Public Function GeminiJsonStringEscape(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, """", "\""")
    t = Replace(t, vbCr, "\r")
    t = Replace(t, vbLf, "\n")
    t = Replace(t, vbTab, "\t")
    GeminiJsonStringEscape = t
End Function

Public Sub GeminiWriteUtf8File(ByVal filePath As String, ByVal textContent As String)
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

Public Function GeminiReadUtf8File(ByVal filePath As String) As String
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

Public Function GeminiReadUtf8FileViaTempCopy(ByVal filePath As String) As String
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

