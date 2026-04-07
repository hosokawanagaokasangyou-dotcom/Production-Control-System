Private Function Py3VersionOutput(wsh As Object) As String
    Dim execObj As Object
    Dim s As String
    Py3VersionOutput = ""
    On Error GoTo CleanExit
    Set execObj = wsh.Exec("cmd.exe /c py -3 --version")
    Do While execObj.Status = 0
        Sleep 50
    Loop
    s = execObj.StdOut.ReadAll()
    If Len(Trim$(s)) = 0 Then s = execObj.StdErr.ReadAll()
    Py3VersionOutput = s
CleanExit:
End Function

Private Function IsPython3Available(wsh As Object) As Boolean
    Dim s As String
    s = Py3VersionOutput(wsh)
    IsPython3Available = (InStr(1, s, "Python 3", vbTextCompare) > 0)
End Function

Private Function TryInstallPythonViaWinget(wsh As Object) As Boolean
    On Error GoTo Fail
    Dim wingetBat As String
    wingetBat = "@echo off" & vbCrLf & "winget install -e --id Python.Python.3.12 --silent --accept-package-agreements --accept-source-agreements" & vbCrLf & "exit /b %ERRORLEVEL%"
    RunTempCmdWithConsoleLayout wsh, wingetBat
    TryInstallPythonViaWinget = True
    Exit Function
Fail:
    TryInstallPythonViaWinget = False
End Function

Private Function TryInstallPythonViaOfficialInstaller(wsh As Object) As Boolean
    Dim psCmd As String
    Dim shellCmd As String
    Dim exitCode As Long
    On Error GoTo Fail
    psCmd = "powershell -NoProfile -ExecutionPolicy Bypass -Command """ & _
            "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; " & _
            "$url = '" & PY_OFFICIAL_INSTALLER_URL & "'; " & _
            "$out = Join-Path $env:TEMP 'python_official_installer.exe'; " & _
            "Invoke-WebRequest -Uri $url -OutFile $out -UseBasicParsing; " & _
            "if ((Get-Item $out).Length -lt 1MB) { throw 'ダウンロードに失敗した可能性があります' }; " & _
            "$p = Start-Process -FilePath $out -ArgumentList '/quiet InstallAllUsers=1 PrependPath=1 Include_test=0 Include_pip=1 Include_launcher=1' -Wait -PassThru; " & _
            "Remove-Item $out -ErrorAction SilentlyContinue; " & _
            "exit $p.ExitCode"""
    shellCmd = "cmd.exe /c " & psCmd
    exitCode = wsh.Run(shellCmd, 1, True)
    TryInstallPythonViaOfficialInstaller = (exitCode = 0)
    Exit Function
Fail:
    TryInstallPythonViaOfficialInstaller = False
End Function

' ブック直下または python\ 配下の setup_environment.py。戻り値: 相対パス（例 python\setup_environment.py）または空。
Private Function SetupEnvironmentScriptRelativePath(ByVal workDir As String) As String
    If Len(Dir(workDir & "\python\setup_environment.py")) > 0 Then
        SetupEnvironmentScriptRelativePath = "python\setup_environment.py"
    ElseIf Len(Dir(workDir & "\setup_environment.py")) > 0 Then
        SetupEnvironmentScriptRelativePath = "setup_environment.py"
    Else
        SetupEnvironmentScriptRelativePath = ""
    End If
End Function

Private Function RunPipInstallWithRefreshedPath(wsh As Object, ByVal workDir As String, ByVal setupRel As String) As Long
    Dim ps As String
    Dim shellCmd As String
    Dim wdEsc As String
    Dim setupEsc As String
    ' PATH を再合成したうえで、ブックフォルダで setup_environment.py を実行（pip + requirements + xlwings addin）
    wdEsc = Replace(workDir, "'", "''")
    setupEsc = Replace(setupRel, "'", "''")
    ps = "$env:Path = [System.Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [System.Environment]::GetEnvironmentVariable('Path','User'); " & _
         "$py = Get-Command py -ErrorAction SilentlyContinue; " & _
         "if (-not $py) { Write-Error 'py が見つかりません。Excel を一度終了してから再実行するか、PATH を確認してください。'; exit 91 }; " & _
         "Set-Location -LiteralPath '" & wdEsc & "'; " & _
         "& py -3 -u .\" & setupEsc & "; " & _
         "exit $LASTEXITCODE"
    shellCmd = "powershell -NoProfile -ExecutionPolicy Bypass -Command " & Chr(34) & ps & Chr(34)
    RunPipInstallWithRefreshedPath = wsh.Run(shellCmd, 1, True)
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

' =========================================================
' Gemini 認証: Python は「設定」B1 の JSON ファイルパスからキーを読む（平文または暗号化）。
' 暗号化 JSON の復号は planning_core のソース内定数のみ。パスフレーズはシートに保存しない（B2 は未使用またはクリア）。
' =========================================================
