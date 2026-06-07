# UTF-8 BOM: Windows PowerShell 5.1
param(
    [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
)

$ErrorActionPreference = 'Stop'
$launcherDir = Join-Path $RepoRoot 'tools/pm-ai-rdp-remote-launcher'
$resourceDir = Join-Path $RepoRoot 'code_java/src/main/resources/jp/co/pm/ai/desktop/rdp-launcher'
$versionTxt = Join-Path $RepoRoot 'version.txt'

Push-Location $launcherDir
try {
    dotnet publish -c Release -r win-x64 --self-contained false
    $publishDir = Join-Path $launcherDir 'bin/Release/net8.0-windows/win-x64/publish'
    $exe = Join-Path $publishDir 'PmAiRdpRemoteLauncher.exe'
    if (-not (Test-Path -LiteralPath $exe)) {
        throw "Publish output not found: $exe"
    }

    New-Item -ItemType Directory -Force -Path $resourceDir | Out-Null
    Copy-Item -LiteralPath $exe -Destination (Join-Path $resourceDir 'PmAiRdpRemoteLauncher.exe') -Force

    $version = '1.00'
    if (Test-Path -LiteralPath $versionTxt) {
        $version = (Get-Content -LiteralPath $versionTxt -TotalCount 1).Trim()
    }
    $versionPath = Join-Path $resourceDir 'PmAiRdpRemoteLauncher.version.txt'
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($versionPath, $version + [Environment]::NewLine, $utf8NoBom)
    Write-Host "Bundled RDP launcher $version -> $resourceDir"
}
finally {
    Pop-Location
}
