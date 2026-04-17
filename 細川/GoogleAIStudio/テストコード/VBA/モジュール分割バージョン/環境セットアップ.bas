Option Explicit

Public Function Py3VersionOutput(wsh As Object) As String
    Dim execObj As Object
    Dim s As String
    Py3VersionOutput = ""
    On Error GoTo CleanExit
    Set execObj = wsh.Exec("cmd.exe /c py -" & PM_AI_SETUP_PY_MINOR & " --version")
    Do While execObj.Status = 0
        Sleep 50
    Loop
    s = execObj.StdOut.ReadAll()
    If Len(Trim$(s)) = 0 Then s = execObj.StdErr.ReadAll()
    Py3VersionOutput = s
CleanExit:
End Function

Public Function IsPython3Available(wsh As Object) As Boolean
    Dim s As String
    s = Py3VersionOutput(wsh)
    IsPython3Available = (InStr(1, s, "Python " & PM_AI_SETUP_PY_MINOR, vbTextCompare) > 0)
End Function

Public Function TryInstallPythonViaWinget(wsh As Object) As Boolean
    On Error GoTo Fail
    Dim wingetBat As String
    wingetBat = "@echo off" & vbCrLf & "winget install -e --id " & PM_AI_SETUP_WINGET_PYTHON_ID & " --silent --accept-package-agreements --accept-source-agreements" & vbCrLf & "exit /b %ERRORLEVEL%"
    RunTempCmdWithConsoleLayout wsh, wingetBat
    TryInstallPythonViaWinget = True
    Exit Function
Fail:
    TryInstallPythonViaWinget = False
End Function

Public Function TryInstallPythonViaOfficialInstaller(wsh As Object) As Boolean
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

Private Function BuildPmAiPrependPathScriptBody(ByVal pyMinor As String) As String
    BuildPmAiPrependPathScriptBody = _
        "param(" & vbCrLf & _
        "  [Parameter(Mandatory = $true)]" & vbCrLf & _
        "  [ValidateSet('User', 'Machine')]" & vbCrLf & _
        "  [string] $Scope" & vbCrLf & _
        ")" & vbCrLf & _
        "$ErrorActionPreference = 'Stop'" & vbCrLf & _
        "$minor = '" & pyMinor & "'" & vbCrLf & _
        "$pyRoot = (cmd /c ('py -' + $minor + ' -c ""import os,sys;print(os.path.dirname(sys.executable))""')).Trim()" & vbCrLf & _
        "if (-not $pyRoot) { exit 2 }" & vbCrLf & _
        "$scripts = [System.IO.Path]::Combine($pyRoot, 'Scripts')" & vbCrLf & _
        "$cur = [Environment]::GetEnvironmentVariable('Path', $Scope)" & vbCrLf & _
        "if ($null -eq $cur) { $cur = '' }" & vbCrLf & _
        "$parts = @()" & vbCrLf & _
        "if ($cur) {" & vbCrLf & _
        "  $parts = $cur -split ';' | ForEach-Object {" & vbCrLf & _
        "    $t = $_.Trim()" & vbCrLf & _
        "    if ($t) { $t.TrimEnd([char]92) }" & vbCrLf & _
        "  } | Where-Object { $_ -and ($_ -ne $pyRoot) -and ($_ -ne $scripts) }" & vbCrLf & _
        "}" & vbCrLf & _
        "$new = ($pyRoot + ';' + $scripts + ';' + ($parts -join ';')).TrimEnd(';')" & vbCrLf & _
        "[Environment]::SetEnvironmentVariable('Path', $new, $Scope)" & vbCrLf & _
        "exit 0"
End Function

Private Sub WritePmAiPrependPathScriptFile(ByVal fullPath As String, ByVal pyMinor As String)
    Dim fn As Integer
    Dim txt As String
    Dim lines() As String
    Dim i As Long
    txt = BuildPmAiPrependPathScriptBody(pyMinor)
    fn = FreeFile
    Open fullPath For Output As #fn
    lines = Split(txt, vbCrLf)
    For i = LBound(lines) To UBound(lines)
        Print #fn, lines(i)
    Next i
    Close #fn
End Sub

' Python インストール先と Scripts を PATH 先頭へ（User は通常実行、Machine は UAC 昇格で更新）。
' machinePathUpdated: 昇格プロセスが exit 0 で終わったとき True（キャンセル・失敗は False）。
Public Function PrependPmAiPythonInstallDirsOnRegistryPath(wsh As Object, ByRef machinePathUpdated As Boolean) As Boolean
    Dim ps1 As String
    Dim rcUser As Long
    Dim rcElev As Long
    Dim outerPs As String
    
    machinePathUpdated = False
    On Error GoTo Fail
    
    ps1 = wsh.ExpandEnvironmentStrings("%TEMP%\pm_ai_prepend_python_path.ps1")
    Call WritePmAiPrependPathScriptFile(ps1, CStr(PM_AI_SETUP_PY_MINOR))
    
    rcUser = wsh.Run("powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & ps1 & """ -Scope User", 0, True)
    If rcUser <> 0 Then
        PrependPmAiPythonInstallDirsOnRegistryPath = False
        Exit Function
    End If
    
    MacroSplash_SetStep "環境構築: システムの環境変数 Path を更新します。UAC で許可してください…"
    
    outerPs = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " & Chr(34) & _
              "$p = Join-Path $env:TEMP 'pm_ai_prepend_python_path.ps1'; " & _
              "if (-not (Test-Path -LiteralPath $p)) { exit 9 }; " & _
              "try { " & _
              "  $pr = Start-Process -FilePath powershell.exe -Verb RunAs -Wait -PassThru " & _
              "    -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',$p,'-Scope','Machine'); " & _
              "  if ($null -eq $pr) { exit 6 }; " & _
              "  exit $pr.ExitCode " & _
              "} catch { exit 5 }" & Chr(34)
    
    rcElev = wsh.Run(outerPs, 0, True)
    machinePathUpdated = (rcElev = 0)
    
    PrependPmAiPythonInstallDirsOnRegistryPath = True
    Exit Function
Fail:
    PrependPmAiPythonInstallDirsOnRegistryPath = False
End Function

' ブック直下または python\ 配下の setup_environment.py。戻り値: 相対パス（例 python\setup_environment.py）または空。
Public Function SetupEnvironmentScriptRelativePath(ByVal workDir As String) As String
    If Len(Dir(workDir & "\python\setup_environment.py")) > 0 Then
        SetupEnvironmentScriptRelativePath = "python\setup_environment.py"
    ElseIf Len(Dir(workDir & "\setup_environment.py")) > 0 Then
        SetupEnvironmentScriptRelativePath = "setup_environment.py"
    Else
        SetupEnvironmentScriptRelativePath = ""
    End If
End Function

Public Function RunPipInstallWithRefreshedPath(wsh As Object, ByVal workDir As String, ByVal setupRel As String, ByVal macroBookFullName As String) As Long
    Dim ps As String
    Dim shellCmd As String
    Dim wdEsc As String
    Dim setupEsc As String
    Dim wbEsc As String
    ' PATH を再合成したうえで、ブックフォルダで setup_environment.py を実行（pip + requirements + xlwings addin）
    ' TASK_INPUT_WORKBOOK: Python が「設定_環境変数」シート（例: PM_AI_CMD_PAUSE_ON_ERROR）を読むため
    On Error Resume Next
    wsh.Environment("Process")("TASK_INPUT_WORKBOOK") = macroBookFullName
    On Error GoTo 0
    wdEsc = Replace(workDir, "'", "''")
    setupEsc = Replace(setupRel, "'", "''")
    wbEsc = Replace(macroBookFullName, "'", "''")
    ps = "$env:TASK_INPUT_WORKBOOK='" & wbEsc & "'; $env:PYTHONUTF8='1'; $env:PYTHONIOENCODING='utf-8'; " & _
         "$pmPyRoot = (cmd /c 'py -" & PM_AI_SETUP_PY_MINOR & " -c ""import os,sys;print(os.path.dirname(sys.executable))""').Trim(); " & _
         "$pmPre = ''; if ($pmPyRoot) { $pmPre = ($pmPyRoot + ';' + ([System.IO.Path]::Combine($pmPyRoot, 'Scripts')) + ';') }; " & _
         "$env:Path = $pmPre + [System.Environment]::GetEnvironmentVariable('Path','Machine') + ';' + [System.Environment]::GetEnvironmentVariable('Path','User'); " & _
         "$py = Get-Command py -ErrorAction SilentlyContinue; " & _
         "if (-not $py) { Write-Error 'py が見つかりません。Excel を一度終了してから再実行するか、PATH を確認してください。'; exit 91 }; " & _
         "Set-Location -LiteralPath '" & wdEsc & "'; " & _
         "& py -" & PM_AI_SETUP_PY_MINOR & " -X utf8 -u .\" & setupEsc & "; " & _
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
    Dim machinePathUpdated As Boolean
    
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
        MsgBox "Python " & PM_AI_SETUP_PY_MINOR & "（py -" & PM_AI_SETUP_PY_MINOR & "）が見つかりません。" & vbCrLf & _
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
                       "https://www.python.org/downloads/windows/ から Python " & PM_AI_SETUP_PY_MINOR & " をインストール（Add python.exe to PATH）後、本マクロを再実行してください。", vbCritical
                Exit Sub
            End If
        End If
        
        If Not IsPython3Available(wsh) Then
            MsgBox "インストール後も py -" & PM_AI_SETUP_PY_MINOR & " を認識できませんでした。" & vbCrLf & _
                   "Excel をいったん終了し、再起動してから再度「環境構築」を実行してください。", vbExclamation
            Exit Sub
        End If
        
        msg = Trim$(Py3VersionOutput(wsh))
        MacroSplash_SetStep "Python を検出しました（" & Left$(msg, 80) & "…）。pip でライブラリをインストールします…"
    End If
    
    MacroSplash_SetStep "環境構築: Python " & PM_AI_SETUP_PY_MINOR & " を PATH 先頭に登録しています…"
    machinePathUpdated = False
    If Not PrependPmAiPythonInstallDirsOnRegistryPath(wsh, machinePathUpdated) Then
        MsgBox "Python " & PM_AI_SETUP_PY_MINOR & " は検出できましたが、ユーザー環境変数 Path への先頭登録に失敗しました。" & vbCrLf & _
               "setup_environment.py の実行は続行しますが、手動で Path を確認してください。", vbExclamation
    ElseIf Not machinePathUpdated Then
        MsgBox "システム（マシン）環境変数 Path の先頭登録は、UAC をキャンセルしたか管理者権限での更新に失敗したためスキップされました。" & vbCrLf & _
               "ユーザー Path は更新済みです。システム全体で python を最優先にしたい場合は、環境構築を再実行し UAC で許可してください。" & vbCrLf & _
               "setup_environment.py の実行は続行します。", vbInformation
    End If
    
    MacroSplash_SetStep "環境構築: setup_environment.py を実行しています（しばらくお待ちください）…"
    pipExit = RunPipInstallWithRefreshedPath(wsh, workDir, setupRel, ThisWorkbook.FullName)
    If pipExit <> 0 Then
        MsgBox "setup_environment.py の実行に失敗しました（exitCode=" & CStr(pipExit) & "）。" & vbCrLf & _
               "コマンドプロンプトでブックのフォルダを開き、次を手動実行してください。" & vbCrLf & vbCrLf & _
               "cd /d """ & workDir & """" & vbCrLf & _
               "py -" & PM_AI_SETUP_PY_MINOR & " " & setupRel, vbCritical
        Exit Sub
    End If

    MacroSplash_SetStep "環境構築が完了しました。xlwings リボンが無い場合は Excel を再起動してください。"
    m_animMacroSucceeded = True
End Sub

' =========================================================
' Gemini 認証: Python は「設定」B1 の JSON ファイルパスからキーを読む（平文または暗号化）。
' 暗号化 JSON の復号は planning_core のソース内定数のみ。パスフレーズはシートに保存しない（B2 は未使用またはクリア）。
' =========================================================
