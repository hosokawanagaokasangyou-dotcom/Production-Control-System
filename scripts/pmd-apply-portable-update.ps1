# Applies staged portable desktop bundle (PMD.exe, app, runtime) after the running PMD process exits.
# Invoked from Java PortableBundleUpdateLauncher; do not run manually unless debugging.
param(
    [Parameter(Mandatory = $true)][string]$InstallRoot,
    [Parameter(Mandatory = $true)][string]$StagingRoot,
    [Parameter(Mandatory = $true)][long]$WaitPid,
    [Parameter(Mandatory = $true)][string]$LogFile,
    [string]$VersionLabel = '',
    [string]$CanonicalPath = ''
)

$ErrorActionPreference = 'Stop'

function Write-Log([string]$Message) {
    $line = "[{0}] {1}" -f (Get-Date -Format 'o'), $Message
    Add-Content -LiteralPath $LogFile -Value $line -Encoding utf8
}

try {
    $install = [System.IO.Path]::GetFullPath($InstallRoot)
    $staging = [System.IO.Path]::GetFullPath($StagingRoot)
    $logDir = Split-Path -Parent $LogFile
    if ($logDir -and -not (Test-Path -LiteralPath $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }
    Write-Log "=== pmd-apply-portable-update start ==="
    Write-Log "InstallRoot=$install"
    Write-Log "StagingRoot=$staging"
    Write-Log "WaitPid=$WaitPid VersionLabel=$VersionLabel CanonicalPath=$CanonicalPath"

    if (-not (Test-Path -LiteralPath $staging)) {
        throw "StagingRoot does not exist: $staging"
    }
    if (-not (Test-Path -LiteralPath (Join-Path $staging 'PMD.exe'))) {
        throw "StagingRoot is missing PMD.exe: $staging"
    }

    if ($WaitPid -gt 0) {
        Write-Log "Waiting for PID $WaitPid ..."
        try {
            Wait-Process -Id $WaitPid -Timeout 600 -ErrorAction Stop
            Write-Log "Process $WaitPid exited."
        }
        catch {
            Write-Log "Wait-Process: $($_.Exception.Message) (continue if already exited)"
        }
        Start-Sleep -Seconds 2
    }

  foreach ($name in @('app', 'runtime')) {
        $src = Join-Path $staging $name
        if (Test-Path -LiteralPath $src) {
            $dst = Join-Path $install $name
            Write-Log "robocopy $name ..."
            & robocopy $src $dst /E /NFL /NDL /NJH /NJS /nc /ns /np /R:2 /W:2 /XO | Out-Null
            if ($LASTEXITCODE -ge 8) {
                throw "robocopy $name failed with exit $LASTEXITCODE"
            }
        }
    }

    foreach ($leaf in @('PMD.exe', 'launch-pm-ai-desktop.bat', 'version.txt', 'pmd-apply-portable-update.ps1')) {
        $srcFile = Join-Path $staging $leaf
        if (Test-Path -LiteralPath $srcFile) {
            $dstFile = Join-Path $install $leaf
            Write-Log "Copy-Item $leaf"
            Copy-Item -LiteralPath $srcFile -Destination $dstFile -Force
        }
    }

    $pmd = Join-Path $install 'PMD.exe'
    if (-not (Test-Path -LiteralPath $pmd)) {
        throw "PMD.exe missing after apply: $pmd"
    }

    $pendingManifest = Join-Path $env:USERPROFILE '.pm-ai-desktop\pending-portable-update.json'
    if (Test-Path -LiteralPath $pendingManifest) {
        Remove-Item -LiteralPath $pendingManifest -Force -ErrorAction SilentlyContinue
        Write-Log "Removed pending manifest."
    }
    if (Test-Path -LiteralPath $staging) {
        Remove-Item -LiteralPath $staging -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log "Removed staging directory."
    }

    Write-Log "Starting PMD.exe ..."
    Start-Process -FilePath $pmd -WorkingDirectory $install | Out-Null
    Write-Log "=== pmd-apply-portable-update success ==="
    exit 0
}
catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    Write-Log $_.ScriptStackTrace
    exit 1
}
