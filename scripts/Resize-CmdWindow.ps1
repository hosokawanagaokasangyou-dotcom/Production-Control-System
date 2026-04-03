#requires -Version 5.1
<#
.SYNOPSIS
  現在のコンソル(cmd.exe / PowerShell コンソル)を画面上端左に移動し、
  幅をプライマリモニタ全幅・高さを画面の 1/4 に調整する。
.NOTES
  同じコンソルから実行すること (例: cmd で powershell -File 本スクリプト)。
#>
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class ConsoleWindowLayoutNative {
  [DllImport("kernel32.dll")]
  public static extern IntPtr GetConsoleWindow();
  [DllImport("user32.dll", SetLastError = true)]
  public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);
}
'@

Add-Type -AssemblyName System.Windows.Forms
$bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
$w = [int]$bounds.Width
$h = [int]([Math]::Max(1, [Math]::Floor($bounds.Height / 4)))

$hwnd = [ConsoleWindowLayoutNative]::GetConsoleWindow()
if ($hwnd -eq [IntPtr]::Zero) {
  Write-Error -Message 'GetConsoleWindow returned null (run from cmd.exe or PowerShell console).'
  exit 1
}

$ok = [ConsoleWindowLayoutNative]::MoveWindow($hwnd, 0, 0, $w, $h, $true)
if (-not $ok) {
  Write-Error -Message 'MoveWindow failed.'
  exit 1
}
