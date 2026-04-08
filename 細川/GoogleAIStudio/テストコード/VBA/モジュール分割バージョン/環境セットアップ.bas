Attribute VB_Name = "環境セットアップ"
Option Explicit

Private Function SetupEnvironmentScriptRelativePath(ByVal workDir As String) As String
    If Len(Dir(workDir & "\python\setup_environment.py")) > 0 Then
        SetupEnvironmentScriptRelativePath = "python\setup_environment.py"
    ElseIf Len(Dir(workDir & "\setup_environment.py")) > 0 Then
        SetupEnvironmentScriptRelativePath = "setup_environment.py"
    Else
        SetupEnvironmentScriptRelativePath = ""
    End If
End Function

Sub InstallComponents()
    Dim wsh As Object
    Dim wingetExit As Long
    Dim pipExit As Long
    Dim msg As String
    Dim workDir As String
    Dim setupRel As String
    
    Set wsh = CreateObject("WScript.Shell")
    
    workDir = Trim$(ThisWorkbook.path)
    If Len(workDir) = 0 Then
        MsgBox "先にこのブックを保存してから環境構築を実行してください。" & vbCrLf & _
               "（setup_environment.py は python\ またはブック直下、requirements.txt は python\ に配置）", vbExclamation
        Exit Sub
    End If
    setupRel = SetupEnvironmentScriptRelativePath(workDir)
    If Len(setupRel) = 0 Then
        MsgBox "次のいずれのファイルも見つかりません:" & vbCrLf & _
               workDir & "\python\setup_environment.py" & vbCrLf & _
               "または " & workDir & "\setup_environment.py" & vbCrLf & vbCrLf & _
               "テストコード一式（python フォルダ含む）をブックと同じフォルダにコピーしてから再実行してください。", vbCritical
        Exit Sub
    End If
    
    If Not IsPython3Available(wsh) Then
        MsgBox "Python 3（py -3）が見つかりません。" & vbCrLf & _
               "自動インストールを開始します（winget → 失敗時は python.org のインストーラ）。" & vbCrLf & _
               "管理者権限や UAC の承認が求められる場合があります。数分かかることがあります。", vbInformation
        
        On Error Resume Next
        wingetExit = wsh.Run("cmd.exe /c winget --version", 0, True)
        On Error GoTo 0
        
        If wingetExit = 0 Then
            Call TryInstallPythonViaWinget(wsh)
        End If
        
        If Not IsPython3Available(wsh) Then
            If Not TryInstallPythonViaOfficialInstaller(wsh) Then
                MsgBox "公式インストーラによるセットアップに失敗しました。" & vbCrLf & _
                       "https://www.python.org/downloads/windows/ から Python 3.12 をインストール（Add python.exe to PATH）後、本マクロを再実行してください。", vbCritical
                Exit Sub
            End If
        End If
        
        If Not IsPython3Available(wsh) Then
            MsgBox "インストール後も py -3 を認識できませんでした。" & vbCrLf & _
                   "Excel をいったん終了し、再起動してから再度「環境構築」を実行してください。", vbExclamation
            Exit Sub
        End If
        
        msg = Trim$(Py3VersionOutput(wsh))
        MacroSplash_SetStep "Python を検出しました（" & Left$(msg, 80) & "…）。pip でライブラリをインストールします…"
    End If
    
    MacroSplash_SetStep "環境構築: setup_environment.py を実行しています（しばらくお待ちください）…"
    pipExit = RunPipInstallWithRefreshedPath(wsh, workDir, setupRel)
    If pipExit <> 0 Then
        MsgBox "setup_environment.py の実行に失敗しました（exitCode=" & CStr(pipExit) & "）。" & vbCrLf & _
               "コマンドプロンプトでブックのフォルダを開き、次を手動実行してください。" & vbCrLf & vbCrLf & _
               "cd /d """ & workDir & """" & vbCrLf & _
               "py -3 " & setupRel, vbCritical
        Exit Sub
    End If

    MacroSplash_SetStep "環境構築が完了しました。xlwings リボンが無い場合は Excel を再起動してください。"
    m_animMacroSucceeded = True
End Sub

