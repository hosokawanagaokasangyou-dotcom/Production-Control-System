Attribute VB_Name = "ÉTÉEÉìÉhêßå‰"
Option Explicit

Public Sub PlayFinishSound()
    MacroCompleteChime
End Sub

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

