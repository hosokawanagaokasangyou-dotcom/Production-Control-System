# Remote Desktop RPA Launcher - portable Windows bundle (repo root entry).
#
# Output (repo root):
#   rpa_luncher_release\PmAiRpaLuncher\PmAiRpaLuncher.exe
#   rpa_luncher_release\PmAiRpaLuncher_portable.zip
#
# Dispatch PMD portable is separate: fast_package_app.ps1 -> pm-ai-package-release\
#
# Usage:
#   .\fast_package_rdp_launcher.ps1
#   .\fast_package_rdp_launcher.ps1 -SkipJdkPrepare -SkipJavaFxPrepare -SkipCsLauncherBuild
#   .\fast_package_rdp_launcher.ps1 -SkipCanonicalDeploy

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
