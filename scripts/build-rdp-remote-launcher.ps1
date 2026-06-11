# UTF-8 BOM: Windows PowerShell 5.1
param(
    [string]$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
)

$ErrorActionPreference = 'Stop'
$launcherDir = Join-Path $RepoRoot 'tools/pm-ai-rdp-remote-launcher'
$resourceDir = Join-Path $RepoRoot 'code_java/src/main/resources/jp/co/pm/ai/desktop/rdp-launcher'
$repoVersionTxt = Join-Path $RepoRoot 'version.txt'
$launcherVersionBasename = 'PmAiRdpRemoteLauncher.version.txt'

Push-Location $launcherDir
try {
    if (-not (Test-Path -LiteralPath $repoVersionTxt)) {
        throw "Repository version file not found: $repoVersionTxt"
    }
    $version = (Get-Content -LiteralPath $repoVersionTxt -TotalCount 1).Trim()
    if ([string]::IsNullOrWhiteSpace($version)) {
        throw "Repository version is empty: $repoVersionTxt"
    }

    dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true
    $publishDir = Join-Path $launcherDir 'bin/Release/net8.0-windows/win-x64/publish'
    $exe = Join-Path $publishDir 'PmAiRdpRemoteLauncher.exe'
    if (-not (Test-Path -LiteralPath $exe)) {
        throw "Publish output not found: $exe"
    }

    New-Item -ItemType Directory -Force -Path $resourceDir | Out-Null
    Copy-Item -LiteralPath $exe -Destination (Join-Path $resourceDir 'PmAiRdpRemoteLauncher.exe') -Force

    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    $versionLine = $version + [Environment]::NewLine
    $resourceVersionPath = Join-Path $resourceDir $launcherVersionBasename
    [System.IO.File]::WriteAllText($resourceVersionPath, $versionLine, $utf8NoBom)
    [System.IO.File]::WriteAllText((Join-Path $publishDir $launcherVersionBasename), $versionLine, $utf8NoBom)
    Write-Host "Bundled RDP launcher $version (from $repoVersionTxt) -> $resourceDir"
}
finally {
    Pop-Location
}
