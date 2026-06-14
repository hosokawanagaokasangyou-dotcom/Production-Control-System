# UTF-8 BOM: Windows PowerShell 5.1
param(
    [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
)

$ErrorActionPreference = 'Stop'
$launcherDir = Join-Path $RepoRoot 'tools/pm-ai-rdp-desktop-launcher'
$resourceDir = Join-Path $RepoRoot 'code_java/src/main/resources/jp/co/pm/ai/desktop/rdp-launcher'
$publishName = 'PmAiRdpDesktopLauncher.exe'
$bundleName = 'PmAiRpaLuncher.exe'

Push-Location $launcherDir
try {
    dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true
    $publishDir = Join-Path $launcherDir 'bin/Release/net8.0-windows/win-x64/publish'
    $exe = Join-Path $publishDir $publishName
    if (-not (Test-Path -LiteralPath $exe)) {
        throw "Publish output not found: $exe"
    }

    New-Item -ItemType Directory -Force -Path $resourceDir | Out-Null
    Copy-Item -LiteralPath $exe -Destination (Join-Path $resourceDir $bundleName) -Force
    Write-Host "Bundled desktop launcher -> $resourceDir\$bundleName"
}
finally {
    Pop-Location
}
