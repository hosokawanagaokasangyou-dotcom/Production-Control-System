# UTF-8 BOM: Windows PowerShell 5.1
# PmAiRdpRemoteLauncher.exe の UNC 配備先を診断する（bundle 破損・サイズ不一致の切り分け）。
param(
    [string]$LauncherExePath = '\\192.168.0.101\共有フォルダ\湖南工場\湖南共有\002  加工G\●配台AIシステム\共有DATA\PmAiRdpRemoteLauncher.exe',
    [string]$LogPath = ''
)

$ErrorActionPreference = 'Continue'

function Write-DebugNdjson {
    param(
        [string]$HypothesisId,
        [string]$Location,
        [string]$Message,
        [hashtable]$Data
    )
    $payload = [ordered]@{
        sessionId    = '5dce1b'
        runId        = 'diagnose'
        hypothesisId = $HypothesisId
        location     = $Location
        message      = $Message
        data         = $Data
        timestamp    = [DateTimeOffset]::UtcNow.ToUnixTimeMilliseconds()
    }
    $line = ($payload | ConvertTo-Json -Compress -Depth 6)
    Write-Host $line
    if (-not [string]::IsNullOrWhiteSpace($LogPath)) {
        Add-Content -LiteralPath $LogPath -Value $line -Encoding UTF8
    }
}

function Invoke-LauncherProbe {
    param([string]$Path)
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = $Path
    $psi.UseShellExecute = $false
    $psi.RedirectStandardError = $true
    $psi.RedirectStandardOutput = $true
    $psi.CreateNoWindow = $true
    $p = [System.Diagnostics.Process]::Start($psi)
    $stdout = $p.StandardOutput.ReadToEnd()
    $stderr = $p.StandardError.ReadToEnd()
    $p.WaitForExit()
    return @{
        exitCode = $p.ExitCode
        stdout   = $stdout.Trim()
        stderr   = $stderr.Trim()
    }
}

if ([string]::IsNullOrWhiteSpace($LogPath)) {
    $repoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
    $LogPath = Join-Path $repoRoot '.cursor\debug-5dce1b.log'
}

Write-DebugNdjson -HypothesisId 'H2' -Location 'diagnose:entry' -Message 'diagnose start' -Data @{ exePath = $LauncherExePath }

if (-not (Test-Path -LiteralPath $LauncherExePath)) {
    Write-DebugNdjson -HypothesisId 'H2' -Location 'diagnose:missing' -Message 'exe not found' -Data @{ exePath = $LauncherExePath }
    exit 2
}

$item = Get-Item -LiteralPath $LauncherExePath
Write-DebugNdjson -HypothesisId 'H2' -Location 'diagnose:size' -Message 'share exe metadata' -Data @{
    lengthBytes  = $item.Length
    lastWriteUtc = $item.LastWriteTimeUtc.ToString('o')
    fullName     = $item.FullName
}

$uncProbe = Invoke-LauncherProbe -Path $LauncherExePath
Write-DebugNdjson -HypothesisId 'H1' -Location 'diagnose:unc-run' -Message 'run from share path' -Data $uncProbe

$localCopy = Join-Path $env:TEMP ('PmAiRdpRemoteLauncher-diagnose-' + [Guid]::NewGuid().ToString('N') + '.exe')
Copy-Item -LiteralPath $LauncherExePath -Destination $localCopy -Force
$copySize = (Get-Item -LiteralPath $localCopy).Length
Write-DebugNdjson -HypothesisId 'H2' -Location 'diagnose:local-copy-size' -Message 'copied to temp' -Data @{
    localPath   = $localCopy
    lengthBytes = $copySize
}

$localProbe = Invoke-LauncherProbe -Path $localCopy
Write-DebugNdjson -HypothesisId 'H2' -Location 'diagnose:local-run' -Message 'run from local temp copy' -Data $localProbe

Remove-Item -LiteralPath $localCopy -Force -ErrorAction SilentlyContinue

$stderr = $uncProbe.stderr
if ($stderr -match 'Arithmetic overflow while reading bundle' -or $stderr -match 'Failure processing application bundle') {
    Write-DebugNdjson -HypothesisId 'H2' -Location 'diagnose:verdict' -Message 'bundle read failed — likely truncated/corrupt exe on share' -Data @{
        shareBytes = $item.Length
        hint       = 'Rebuild with scripts/build-rdp-remote-launcher.ps1 and redeploy intact exe (compare file size after copy).'
    }
}

Write-Host ''
Write-Host '診断ログ: ' -NoNewline
Write-Host $LogPath
