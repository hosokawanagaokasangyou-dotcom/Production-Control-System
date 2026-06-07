# UTF-8 BOM: Windows PowerShell 5.1
param(
    [Parameter(Mandatory = $true)]
    [string]$LauncherExePath,
    [string]$TaskName = 'PM-AI-RDP-Remote-Launcher',
    [string]$UserId = $env:USERNAME
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $LauncherExePath)) {
    throw "Launcher exe not found: $LauncherExePath"
}

$action = New-ScheduledTaskAction -Execute $LauncherExePath
$trigger = New-ScheduledTaskTrigger -AtLogOn -User $UserId
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
$settings.ExecutionTimeLimit = 'PT0S'
$principal = New-ScheduledTaskPrincipal -UserId $UserId -LogonType Interactive -RunLevel Limited

Register-ScheduledTask -TaskName $TaskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force
Write-Host "Registered scheduled task: $TaskName -> $LauncherExePath (user: $UserId)"
