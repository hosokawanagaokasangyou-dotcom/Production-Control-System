# Stage-2 debug run for session 8c2e2d
$ErrorActionPreference = "Stop"
$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$PkgData = Join-Path $RepoRoot "pm-ai-package-release\PMD_initial_install\pm-ai-data"

$env:PM_AI_REPO_ROOT = $RepoRoot
$env:PM_AI_CODE_PYTHON_DIR = Join-Path $RepoRoot "code\python"
$env:PM_AI_OUTPUT_DIR = Join-Path $RepoRoot "output"
$e64OnlyPlan = Join-Path $RepoRoot "output\plan_input_e64_only_debug.xlsx"
$pkgPlan = Join-Path $PkgData "output\plan_input_tasks.xlsx"
$localPlan = Join-Path $RepoRoot "output\plan_input_tasks.xlsx"
if (Test-Path -LiteralPath $e64OnlyPlan) {
    $env:PM_AI_PLAN_INPUT_PATH = $e64OnlyPlan
} elseif (Test-Path -LiteralPath $pkgPlan) {
    $env:PM_AI_PLAN_INPUT_PATH = $pkgPlan
} elseif (Test-Path -LiteralPath $localPlan) {
    $env:PM_AI_PLAN_INPUT_PATH = $localPlan
} else {
    Write-Error "plan_input_tasks.xlsx not found"
}
$pkgProc = Join-Path $PkgData ".pm-ai-cache\network-source\task-input-newest.xlsx"
$localProc = Join-Path $RepoRoot ".pm-ai-cache\network-source\task-input-newest.xlsx"
if (Test-Path -LiteralPath $pkgProc) {
    $env:PM_AI_PROCESSING_PLAN_PATH = $pkgProc
} elseif (Test-Path -LiteralPath $localProc) {
    $env:PM_AI_PROCESSING_PLAN_PATH = $localProc
}
$masterSnap = Join-Path $PkgData ".pm-ai-cache\stage2-run-snapshot\master-workbook.xlsm"
$masterBundled = Join-Path $PkgData "code\master.xlsm"
if (Test-Path -LiteralPath $masterSnap) {
    $env:PM_AI_MASTER_WORKBOOK = $masterSnap
} elseif (Test-Path -LiteralPath $masterBundled) {
    $env:PM_AI_MASTER_WORKBOOK = $masterBundled
} else {
    Write-Error "master workbook snapshot not found"
}
$env:PM_AI_AGENT_DEBUG_SESSION = "8c2e2d"
$env:PM_AI_DEBUG_LOG = Join-Path $RepoRoot ".cursor\debug-8c2e2d.log"
$env:PM_AI_STAGE2_WRITE_EXCEL = "0"
$env:PM_AI_STAGE2_SKIP_TODAY_DISPATCH = "1"
$env:PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH = "0"
$inProg = Join-Path $PkgData ".pm-ai-cache\stage2-in-progress-next-day-dispatch.json"
if (-not (Test-Path -LiteralPath $inProg)) {
    $inProg = Join-Path $RepoRoot ".pm-ai-cache\stage2-in-progress-next-day-dispatch.json"
}
if (Test-Path -LiteralPath $inProg) {
    $env:PM_AI_STAGE2_IN_PROGRESS_NEXT_DAY_DISPATCH_JSON = $inProg
}
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"

$logPath = $env:PM_AI_DEBUG_LOG
if (Test-Path -LiteralPath $logPath) { Remove-Item -LiteralPath $logPath -Force }

$pyDir = Join-Path $RepoRoot "code\python"
Set-Location $pyDir
Write-Host "[debug-8c2e2d] repo=$RepoRoot"
Write-Host "[debug-8c2e2d] plan=$($env:PM_AI_PLAN_INPUT_PATH)"
Write-Host "[debug-8c2e2d] master=$($env:PM_AI_MASTER_WORKBOOK)"
Write-Host "[debug-8c2e2d] log=$logPath"
$proc = Start-Process -FilePath "py" -ArgumentList @("-3.14", "-X", "utf8", "-u", "plan_simulation_stage2.py") `
    -WorkingDirectory $pyDir -NoNewWindow -Wait -PassThru `
    -RedirectStandardOutput (Join-Path $RepoRoot "log\stage2_debug_8c2e2d_stdout.txt") `
    -RedirectStandardError (Join-Path $RepoRoot "log\stage2_debug_8c2e2d_stderr.txt")
Write-Host "[debug-8c2e2d] exit=$($proc.ExitCode)"
exit $proc.ExitCode
