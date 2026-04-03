#requires -Version 5.1
<#
  Excel VBA から起動する cmd.exe のコンソールを、プライマリモニタの左上・幅いっぱい・高さ約 1/4 にしてから終了まで待つ。
  用法: powershell -NoProfile -ExecutionPolicy Bypass -File Run-CmdWithConsoleLayout.ps1 -CmdFile "C:\Temp\xxx.cmd"
#>
param(
  [Parameter(Mandatory = $true)]
  [string] $CmdFile
)

$ErrorActionPreference = 'Stop'
if (-not (Test-Path -LiteralPath $CmdFile)) {
  Write-Error "CmdFile not found: $CmdFile"
  exit 1
}

$full = (Resolve-Path -LiteralPath $CmdFile).Path

Add-Type -AssemblyName System.Windows.Forms | Out-Null
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class CmdLayoutMoveWindow {
  [DllImport("user32.dll", SetLastError = true)]
  public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
}
'@ | Out-Null

$comspec = [Environment]::GetEnvironmentVariable('ComSpec')
if ([string]::IsNullOrEmpty($comspec)) {
  $comspec = Join-Path $env:SystemRoot 'System32\cmd.exe'
}

# cmd /c で .cmd を実行（パスに空白があっても渡せるよう 1 要素でクォート）
$p = Start-Process -FilePath $comspec -ArgumentList @('/c', "`"$full`"") -PassThru -WindowStyle Normal `
  -WorkingDirectory ([System.IO.Path]::GetDirectoryName($full))

$deadline = (Get-Date).AddSeconds(25)
while ((Get-Date) -lt $deadline) {
  if ($p.HasExited) { break }
  try { $p.Refresh() } catch { }
  if ($p.MainWindowHandle -ne [IntPtr]::Zero) { break }
  Start-Sleep -Milliseconds 50
}

if (-not $p.HasExited -and $p.MainWindowHandle -ne [IntPtr]::Zero) {
  $b = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
  $h = [Math]::Max(120, [int][Math]::Floor($b.Height / 4))
  [void][CmdLayoutMoveWindow]::MoveWindow($p.MainWindowHandle, $b.Left, $b.Top, $b.Width, $h, $true)
}

$p.WaitForExit()
exit $p.ExitCode
