# Remote Desktop RPA Launcher - portable Windows bundle (repo root entry).
#
# Prerequisites: Windows, Maven, network for JDK/JavaFX cache on first run.
# Optional: dotnet SDK for PmAiRdpRemoteLauncher.exe in launcher-deploy-seed.
#
# Output: rpa_luncher_release\PmAiRpaLuncher\PmAiRpaLuncher.exe (double-click)
#         rpa_luncher_release\PmAiRpaLuncher_portable.zip
#
# Usage:
#   .\fast_package_rdp_launcher.ps1
#   .\fast_package_rdp_launcher.ps1 -PackageType app-image
#   .\fast_package_rdp_launcher.ps1 -SkipCanonicalDeploy   # offline / WSL without corp network

# UTF-8 BOM: Windows PowerShell 5.1
[CmdletBinding()]
param(
    [ValidateSet('app-image', 'exe', 'msi')]
    [string]$PackageType = 'app-image',

    [switch]$WinConsole,

    [switch]$SkipJdkPrepare,

    [switch]$SkipJavaFxPrepare,

    [switch]$SkipCsLauncherBuild,

    [switch]$SkipZip,

    [switch]$SkipCanonicalDeploy,

    [string]$CanonicalDeployDir = '',

    [string]$JdkRuntimeImage = '',

    [string]$JpackageDest = ''
)

$ErrorActionPreference = 'Stop'
Write-Host 'DEPRECATED: PmAiRpaLuncher.exe の単体配布は廃止しました。' -ForegroundColor Yellow
Write-Host '  操作者 PC: 配台 PMD.exe → リモートデスクトップタブ' -ForegroundColor Yellow
Write-Host '  接続先 PC: PmAiRdpRemoteLauncher.exe（scripts/build-rdp-remote-launcher.ps1）' -ForegroundColor Yellow
exit 1

$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
$CodeJava = Join-Path $ScriptRoot 'code_java'
$pack = Join-Path $CodeJava 'package_rdp_launcher_app.ps1'
if (-not (Test-Path -LiteralPath $pack)) {
    throw "Missing: $pack"
}

$invokeArgs = @{
    PackageType = $PackageType
}
if ($WinConsole) { $invokeArgs.WinConsole = $true }
if ($SkipJdkPrepare) { $invokeArgs.SkipJdkPrepare = $true }
if ($SkipJavaFxPrepare) { $invokeArgs.SkipJavaFxPrepare = $true }
if ($SkipCsLauncherBuild) { $invokeArgs.SkipCsLauncherBuild = $true }
if ($SkipZip) { $invokeArgs.SkipZip = $true }
if ($SkipCanonicalDeploy) { $invokeArgs.SkipCanonicalDeploy = $true }
if (-not [string]::IsNullOrWhiteSpace($CanonicalDeployDir)) {
    $invokeArgs.CanonicalDeployDir = $CanonicalDeployDir.Trim()
}
if (-not [string]::IsNullOrWhiteSpace($JdkRuntimeImage)) {
    $invokeArgs.JdkRuntimeImage = $JdkRuntimeImage.Trim()
}
if (-not [string]::IsNullOrWhiteSpace($JpackageDest)) {
    $invokeArgs.JpackageDest = $JpackageDest.Trim()
}

Push-Location $CodeJava
try {
    & $pack @invokeArgs
    if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }
}
finally {
    Pop-Location
}
