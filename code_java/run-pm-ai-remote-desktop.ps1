#Requires -Version 5.1
<#
.SYNOPSIS
  開発用: RemoteDesktopFxApp を Maven exec で起動する（ポータブル配布は fast_package_rdp_launcher.ps1）。

.DESCRIPTION
  本番向けポータブルはリポジトリ直下 rpa_luncher_release\PmAiRpaLuncher\ を使用する。
  配台 PMD は pm-ai-package-release\（fast_package_app.ps1）。

.EXAMPLE
  .\run-pm-ai-remote-desktop.ps1
#>
param(
    [string] $MaxHeap = ""
)

$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

$commonArgs = @("-q")
if (-not [string]::IsNullOrWhiteSpace($MaxHeap)) {
    $commonArgs += "-Djvm.max.heap.rdp=$MaxHeap"
}

& "$PSScriptRoot\mvnw.cmd" @commonArgs @("compile", "exec:exec@pm-ai-remote-desktop")
exit $LASTEXITCODE
