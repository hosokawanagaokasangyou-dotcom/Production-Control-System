# W6-4 dispatch debug auto loop (max 10 iterations)
param(
    [string]$RepoRoot = "",
    [string]$MasterWorkbook = "",
    [string]$TraceTaskId = "W6-4",
    [string]$DebugSession = "55255a",
    [int]$MaxIterations = 10
)

$ErrorActionPreference = "Stop"
if (-not $RepoRoot) {
    $RepoRoot = $env:PM_AI_REPO_ROOT
}
if (-not $RepoRoot) {
    $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}

$pyDir = Join-Path $RepoRoot "code\python"
$debugLog = Join-Path $RepoRoot ".cursor\debug-$DebugSession.log"
$sidecarLog = Join-Path $RepoRoot "log\w64_agent_debug.ndjson"
$headlessLog = Join-Path $RepoRoot "log\stage2_headless_last.txt"
$corePy = Join-Path $RepoRoot "code\python\planning_core\_core.py"
$coreBackup = Join-Path $RepoRoot "log\_core_w64_loop_backup.py"
$headlessScript = Join-Path $RepoRoot "scripts\run_stage2_headless.ps1"

if (-not $MasterWorkbook) {
    $MasterWorkbook = $env:PM_AI_MASTER_WORKBOOK
}
if (-not $MasterWorkbook) {
    $localMaster = Join-Path $RepoRoot "master.xlsm"
    if (Test-Path -LiteralPath $localMaster) {
        $MasterWorkbook = $localMaster
    }
}
if (-not $MasterWorkbook) {
    Write-Error "Set PM_AI_MASTER_WORKBOOK or place master.xlsm under RepoRoot=$RepoRoot"
    exit 2
}

function Clear-DebugLogs {
    foreach ($p in @($debugLog, $sidecarLog, $headlessLog)) {
        if (Test-Path $p) { Remove-Item $p -Force }
    }
}

function Invoke-Stage2WithEnv {
    param([hashtable]$ExtraEnv = @{})
    Clear-DebugLogs
    foreach ($k in $ExtraEnv.Keys) {
        Set-Item -Path "env:$k" -Value $ExtraEnv[$k]
    }
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $headlessScript `
        -RepoRoot $RepoRoot -MasterWorkbook $MasterWorkbook
    return $LASTEXITCODE
}

function Get-Stage2FailureSummary {
    param([string]$LogPath)
    if (-not (Test-Path $LogPath)) { return @{ remaining = -1; samples = "" } }
    $text = Get-Content -LiteralPath $LogPath -Raw -Encoding UTF8
    $m = [regex]::Match($text, "remaining tasks.*?(\d+)", "IgnoreCase")
    if (-not $m.Success) {
        $m = [regex]::Match($text, "残タスク\s*(\d+)\s*件")
    }
    $samples = [regex]::Match($text, "例:\s*([^\r\n]+)")
    return @{
        remaining = if ($m.Success) { [int]$m.Groups[1].Value } else { -1 }
        samples = if ($samples.Success) { $samples.Groups[1].Value.Trim() } else { "" }
    }
}

function Get-OtherPendingTasks {
    param([string]$LogPath, [string]$TraceId)
    $sum = Get-Stage2FailureSummary -LogPath $LogPath
    $others = @()
    foreach ($part in ($sum.samples -split "、")) {
        $p = $part.Trim()
        if ($p -and ($p -notmatch [regex]::Escape($TraceId))) {
            $others += $p
        }
    }
    return $others
}

function Restore-CoreFromBackup {
    if (Test-Path $coreBackup) {
        Copy-Item -LiteralPath $coreBackup -Destination $corePy -Force
        Write-Host "[loop] restored _core.py from backup"
    }
}

if (-not (Test-Path (Join-Path $RepoRoot "output\plan_input_tasks.xlsx"))) {
    Write-Error "plan_input_tasks.xlsx is missing"
    exit 2
}

Copy-Item -LiteralPath $corePy -Destination $coreBackup -Force

$fixSteps = @(
    @{ name = "baseline_start_date_sync"; env = @{}; revertCore = $false },
    @{ name = "trial_order_strict_off"; env = @{ STAGE2_GLOBAL_DISPATCH_TRIAL_ORDER_STRICT = "0" }; revertCore = $false },
    @{ name = "retry_shift_due_on_partial"; env = @{ STAGE2_RETRY_SHIFT_DUE_ON_PARTIAL_REMAINING = "1" }; revertCore = $false }
)

$iter = 0
foreach ($step in $fixSteps) {
    if ($iter -ge $MaxIterations) { break }
    $iter++
    Write-Host "=== iteration $iter / $MaxIterations : $($step.name) ==="
    if ($step.revertCore) { Restore-CoreFromBackup }
    $exit = Invoke-Stage2WithEnv -ExtraEnv $step.env
    if ($exit -eq 0) {
        Write-Host "[loop] stage2 OK exit=0 ($($step.name))"
        exit 0
    }
    if ($exit -ne 3) {
        Write-Host "[loop] unexpected exit=$exit ($($step.name))"
        continue
    }
    $others = Get-OtherPendingTasks -LogPath $headlessLog -TraceId $TraceTaskId
    if ($others.Count -gt 0) {
        Write-Host "[loop] other pending tasks: $($others -join ', ')"
        continue
    }
    $sum = Get-Stage2FailureSummary -LogPath $headlessLog
    Write-Host "[loop] remaining=$($sum.remaining) samples=$($sum.samples)"
}

Write-Error "W6-4 still pending after $MaxIterations iteration(s)"
exit 3
