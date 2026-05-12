# PMD.exe が起動し、短時間プロセスが生存することを確認する（自動テスト用）。
#
# 前提: Windows で code_java\package_app.ps1 を実行済み（既定では code_java\dist\PMD\PMD.exe）。
#
# 例:
#   .\scripts\test-pmai-desktop-launch.ps1
#   .\scripts\test-pmai-desktop-launch.ps1 -ExePath 'D:\build\PMD\PMD.exe'
#   .\scripts\test-pmai-desktop-launch.ps1 -LeaveRunning   # 検証後も終了させない

param(
    [string]$ExePath = '',
    [int]$StartupWaitSec = 15,
    [switch]$LeaveRunning
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($ExePath)) {
    $here = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
    $ExePath = Join-Path $here '..\code_java\dist\PMD\PMD.exe'
}

try {
    $ExePath = (Resolve-Path -LiteralPath $ExePath).Path
}
catch {
    Write-Error @"
PMD.exe が見つかりません: $ExePath
先に code_java フォルダで package_app.ps1 を実行して dist を生成してください。
"@
    exit 2
}

Write-Host "--- PMD 起動テスト ---" -ForegroundColor Cyan
Write-Host "Exe: $ExePath"

$workDir = Split-Path -Parent $ExePath
$p = Start-Process -FilePath $ExePath -WorkingDirectory $workDir -PassThru

Start-Sleep -Seconds $StartupWaitSec

$alive = Get-Process -Id $p.Id -ErrorAction SilentlyContinue
if (-not $alive) {
    $code = $p.ExitCode
    Write-Error "プロセスが起動後すぐ終了しました（終了コード: $code）。ログまたはイベントログを確認してください。"
    exit 1
}

Write-Host "OK: プロセス稼働中 (PID $($p.Id))。" -ForegroundColor Green

if (-not $LeaveRunning) {
    Write-Host "テスト完了のためプロセスを終了します…" -ForegroundColor DarkGray
    Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
}

exit 0
