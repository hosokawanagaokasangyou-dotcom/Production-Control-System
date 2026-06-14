#Requires -Version 5.1
<#
.SYNOPSIS
  廃止: 単体 RemoteDesktopFxApp の起動は PMD.exe のリモートデスクトップタブへ移行しました。
#>
$ErrorActionPreference = "Stop"
Write-Host "廃止: PmAiRpaLuncher / RemoteDesktopFxApp の単体起動は使用しません。" -ForegroundColor Yellow
Write-Host "  配台 PMD.exe を起動し、メインシェル「リモートデスクトップ」タブを使用してください。" -ForegroundColor Yellow
Write-Host "  接続先 PC では PmAiRdpRemoteLauncher.exe を配備してください。" -ForegroundColor Yellow
exit 1
