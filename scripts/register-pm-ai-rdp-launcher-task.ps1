# UTF-8 BOM: Windows PowerShell 5.1
param(
    [Parameter(Mandatory = $true)]
    [string]$LauncherExePath,
    [string]$IniPath = '',
    [string]$TaskName = 'PM-AI-RDP-Remote-Launcher',
    [string]$UserId = $(if ($env:USERDOMAIN -and $env:USERNAME) { "$($env:USERDOMAIN)\$($env:USERNAME)" } else { $env:USERNAME }),
    [ValidateSet('LogOn', 'RemoteConnect')]
    [string]$TriggerKind = 'RemoteConnect'
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $LauncherExePath)) {
    throw "Launcher exe not found: $LauncherExePath"
}

$launcherDir = Split-Path -LiteralPath $LauncherExePath -Parent
if ([string]::IsNullOrWhiteSpace($IniPath)) {
  # PM-AI が書き込む ini はサマリ Excel と同階層（共有 DATA）の RAP設定.ini。
  # portable 等 exe だけ別フォルダのときは -IniPath で共有 DATA を明示指定すること。
  $IniPath = Join-Path $launcherDir 'RAP設定.ini'
}

$arguments = "--ini `"$IniPath`""
$action = New-ScheduledTaskAction -Execute $LauncherExePath -Argument $arguments

if ($TriggerKind -eq 'RemoteConnect') {
    $trigger = New-ScheduledTaskTrigger -AtLogOn -User $UserId
    # GUI 上「リモート接続時」と同等にするには SessionStateChange が必要なため XML で登録する。
    $taskXml = @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <Triggers>
    <SessionStateChangeTrigger>
      <Enabled>true</Enabled>
      <UserId>$UserId</UserId>
      <StateChange>SessionRemoteConnect</StateChange>
    </SessionStateChangeTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>$UserId</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>true</RunOnlyIfNetworkAvailable>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>$LauncherExePath</Command>
      <Arguments>$arguments</Arguments>
    </Exec>
  </Actions>
</Task>
"@
    Register-ScheduledTask -TaskName $TaskName -Xml $taskXml -Force | Out-Null
}
else {
    $trigger = New-ScheduledTaskTrigger -AtLogOn -User $UserId
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
    $settings.ExecutionTimeLimit = 'PT0S'
    $settings.RunOnlyIfNetworkAvailable = $true
    $principal = New-ScheduledTaskPrincipal -UserId $UserId -LogonType Interactive -RunLevel Limited
    Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force | Out-Null
}

Write-Host "Registered scheduled task: $TaskName"
Write-Host "  Execute : $LauncherExePath"
Write-Host "  Argument: $arguments"
Write-Host "  Trigger : $TriggerKind"
Write-Host "  Log (UNC): $launcherDir\launcher-yyyyMMdd.log"
Write-Host "  Log (TEMP): $env:TEMP\PM-AI-RDP-Launcher\launcher-yyyyMMdd.log"
