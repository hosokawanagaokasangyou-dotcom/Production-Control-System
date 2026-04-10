Option Explicit

Public Function MacroCompleteChime_LocalWavPath() As String
    Dim folder As String
    folder = ThisWorkbook.path
    If Len(folder) = 0 Then Exit Function
    完了チャイム_ローカルWAVパス = folder & "\" & MACRO_COMPLETE_CHIME_REL_DIR & "\" & MACRO_COMPLETE_CHIME_FILE_NAME
End Function

Public Function MacroCompleteChime_LocalMp3Path(ByVal track1to4 As Long) As String
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
    完了チャイム_MP3パスを取得 = folder & "\" & MACRO_COMPLETE_CHIME_REL_DIR & "\" & fn
End Function

Public Function MacroCompleteChime_MciPlayMp3(ByVal fullPath As String) As Boolean
    Dim a As String
    Dim cmdOpen As String
    Dim r As Long
    完了チャイム_MP3をMCI再生 = False
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
    完了チャイム_MP3をMCI再生 = True
    Exit Function
Fail:
    On Error Resume Next
    If Len(a) > 0 Then r = mciSendStringW(StrPtr("close " & a), 0&, 0, 0&)
End Function

Public Function MacroCompleteChime_HttpDownloadBinary(ByVal url As String, ByVal destPath As String) As Boolean
    Dim xhr As Object
    Dim stm As Object
    On Error GoTo Fail
    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
    xhr.Open "GET", url, False
    xhr.setRequestHeader "User-Agent", "Excel-VBA-完了チャイムを再生処理/1"
    xhr.Send
    If xhr.Status < 200 Or xhr.Status >= 300 Then GoTo Fail
    If LenB(xhr.responseBody) = 0 Then GoTo Fail
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 1
    stm.Open
    stm.Write xhr.responseBody
    stm.SaveToFile destPath, 2
    stm.Close
    完了チャイム_HTTPでバイナリ取得 = True
    Exit Function
Fail:
    On Error Resume Next
    If Not stm Is Nothing Then stm.Close
    完了チャイム_HTTPでバイナリ取得 = False
End Function

Public Function MacroCompleteChime_EnsureWavPath() As String
    Dim p As String
    Dim dirSounds As String
    p = 完了チャイム_ローカルWAVパス()
    If Len(p) = 0 Then Exit Function
    If Len(Dir(p)) > 0 Then
        完了チャイム_WAVパスを確保 = p
        Exit Function
    End If
    dirSounds = ThisWorkbook.path & "\" & MACRO_COMPLETE_CHIME_REL_DIR
    On Error Resume Next
    MkDir dirSounds
    On Error GoTo 0
    If 完了チャイム_HTTPでバイナリ取得(MACRO_COMPLETE_CHIME_DOWNLOAD_URL, p) Then
        If Len(Dir(p)) > 0 Then 完了チャイム_WAVパスを確保 = p
    End If
End Function

Public Sub MacroCompleteChime()
    On Error Resume Next
    If Not m_splashAllowMacroSound Then Exit Sub
    Dim track As Long
    Dim mp3 As String
    Dim wav As String
    track = 設定シート_完了チャイムトラック番号()
    mp3 = 完了チャイム_MP3パスを取得(track)
    If Len(mp3) > 0 And Len(Dir(mp3)) > 0 Then
        If 完了チャイム_MP3をMCI再生(mp3) Then Exit Sub
    End If
    wav = 完了チャイム_WAVパスを確保()
    If Len(wav) > 0 Then
        PlaySoundW StrPtr(wav), 0&, SND_FILENAME Or SND_ASYNC
    Else
        PlaySound "SystemAsterisk", 0&, SND_ALIAS Or SND_ASYNC
    End If
End Sub

' MP3 は sndPlaySound（別名 PlaySoundA 系）ではなく MCI（mciSendStringW）で再生。WAV のみの場合は PlaySoundW でも可。
Public Sub 完了音を再生()
    完了チャイムを再生処理
End Sub

Public Function MacroStartBgm_FullPath() As String
    Dim folder As String
    folder = ThisWorkbook.path
    If Len(folder) = 0 Then Exit Function
    起動BGM_フルパス = folder & "\" & MACRO_COMPLETE_CHIME_REL_DIR & "\" & MACRO_START_BGM_FILENAME
End Function

Public Sub MacroStartBgm_CloseHard()
    On Error Resume Next
    If m_macroStartBgmOpen Then
        mciSendStringW StrPtr("close " & MACRO_START_BGM_ALIAS), 0&, 0, 0&
    End If
    m_macroStartBgmOpen = False
End Sub

Public Sub MacroStartBgm_FadeOutAndClose()
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

Public Sub MacroStartBgm_StartIfAvailable()
    Dim p As String
    Dim r As Long
    Dim cmdOpen As String
    On Error Resume Next
    If Not m_splashAllowMacroSound Then Exit Sub
    p = 起動BGM_フルパス()
    If Len(p) = 0 Or Len(Dir(p)) = 0 Then Exit Sub
    起動BGMを強制クローズ
    cmdOpen = "open " & Chr$(34) & p & Chr$(34) & " type mpegvideo alias " & MACRO_START_BGM_ALIAS
    r = mciSendStringW(StrPtr(cmdOpen), 0&, 0, 0&)
    If r <> 0 Then Exit Sub
    mciSendStringW StrPtr("setaudio " & MACRO_START_BGM_ALIAS & " volume to 1000"), 0&, 0, 0&
    r = mciSendStringW(StrPtr("play " & MACRO_START_BGM_ALIAS & " repeat"), 0&, 0, 0&)
    If r <> 0 Then r = mciSendStringW(StrPtr("play " & MACRO_START_BGM_ALIAS), 0&, 0, 0&)
    If r = 0 Then m_macroStartBgmOpen = True
End Sub

' Excel メインウィンドウ（Application.hwnd）の下端・水平中央へ UserForm を SetWindowPos（モードレスのため API で座標指定）
