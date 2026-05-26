# Stage-2 headless run (W6-4 debug)
param(
    [string]$RepoRoot = "",
    [string]$MasterWorkbook = ""
)

$ErrorActionPreference = "Stop"
if (-not $RepoRoot) {
    $RepoRoot = $env:PM_AI_REPO_ROOT
}
if (-not $RepoRoot) {
    $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}

$pyDir = Join-Path $RepoRoot "code\python"
$logOut = Join-Path $RepoRoot "log\stage2_headless_last.txt"

$env:PM_AI_REPO_ROOT = $RepoRoot
$env:PM_AI_CODE_PYTHON_DIR = $pyDir
$env:PM_AI_OUTPUT_DIR = Join-Path $RepoRoot "output"
$env:PM_AI_PLAN_INPUT_PATH = Join-Path $RepoRoot "output\plan_input_tasks.xlsx"

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
    $defaultsJson = Join-Path $RepoRoot "init_setting\session_defaults_kokubu.json"
    if (Test-Path -LiteralPath $defaultsJson) {
        try {
            $raw = Get-Content -LiteralPath $defaultsJson -Raw -Encoding UTF8 | ConvertFrom-Json
            foreach ($row in $raw.uiEnvRows) {
                if ($row.name -eq "PM_AI_MASTER_WORKBOOK" -and $row.value) {
                    $MasterWorkbook = [string]$row.value
                    break
                }
            }
        } catch {
            Write-Host "[stage2-headless] session_defaults master read failed: $_"
        }
    }
}

if (-not $MasterWorkbook) {
    Write-Error "PM_AI_MASTER_WORKBOOK is not set and master.xlsm was not found under RepoRoot=$RepoRoot"
    exit 2
}

$env:PM_AI_MASTER_WORKBOOK = $MasterWorkbook
$env:PM_AI_SKIP_WORKBOOK_ENV_SHEET = "1"
$env:PM_AI_CMD_PAUSE_ON_ERROR = "0"
$env:PM_AI_STAGE2_WRITE_EXCEL = "0"
$env:STAGE2_SKIP_SNAPSHOT_EXPORT = "1"
$env:STAGE2_SKIP_SHEET_VISIBILITY_APPLY = "1"
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"
$env:PM_AI_AGENT_DEBUG_TRACE_TASK_ID = "W6-4"
$env:PM_AI_AGENT_DEBUG_SESSION = "55255a"
$env:PM_AI_DEBUG_LOG = Join-Path $RepoRoot ".cursor\debug-55255a.log"

Set-Location $pyDir
Write-Host "[stage2-headless] repo=$RepoRoot"
Write-Host "[stage2-headless] master=$MasterWorkbook"
Write-Host "[stage2-headless] cwd=$pyDir"
Write-Host "[stage2-headless] log=$logOut"

$prevEap = $ErrorActionPreference
$ErrorActionPreference = "Continue"
$pyArgs = @("-3.14", "-X", "utf8", "-u", "plan_simulation_stage2.py")
$proc = Start-Process -FilePath "py" -ArgumentList $pyArgs -WorkingDirectory $pyDir `
    -RedirectStandardOutput $logOut -RedirectStandardError (Join-Path $RepoRoot "log\stage2_headless_stderr.txt") `
    -NoNewWindow -Wait -PassThru
$exitCode = $proc.ExitCode
$ErrorActionPreference = $prevEap
Write-Host "[stage2-headless] exit=$exitCode"
exit $exitCode
