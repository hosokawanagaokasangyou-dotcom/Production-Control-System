# Stage-1 headless run (plan_input_tasks.xlsx generation)
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
$logOut = Join-Path $RepoRoot "log\stage1_headless_last.txt"
$defaultsJson = Join-Path $RepoRoot "init_setting\session_defaults_kokubu.json"

function Import-UiEnvFromDefaults {
    param([string]$JsonPath)
    if (-not (Test-Path -LiteralPath $JsonPath)) { return @{} }
    $raw = Get-Content -LiteralPath $JsonPath -Raw -Encoding UTF8 | ConvertFrom-Json
    $map = @{}
    foreach ($row in $raw.uiEnvRows) {
        if ($row.name -and ($null -ne $row.value)) {
            $map[[string]$row.name] = [string]$row.value
        }
    }
    return $map
}

$ui = Import-UiEnvFromDefaults -JsonPath $defaultsJson

$env:PM_AI_REPO_ROOT = $RepoRoot
$env:PM_AI_CODE_PYTHON_DIR = $pyDir
$env:PM_AI_OUTPUT_DIR = if ($ui["PM_AI_OUTPUT_DIR"]) { $ui["PM_AI_OUTPUT_DIR"] } else { Join-Path $RepoRoot "output" }
$env:PM_AI_PLAN_INPUT_PATH = Join-Path $env:PM_AI_OUTPUT_DIR "plan_input_tasks.xlsx"
$env:PM_AI_SKIP_WORKBOOK_ENV_SHEET = "1"
$env:PM_AI_CMD_PAUSE_ON_ERROR = "0"
$env:PM_AI_STAGE2_SKIP_IN_PROGRESS_DISPATCH = "0"
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"

foreach ($key in @(
        "PM_AI_TASK_INPUT_SOURCE_DIR",
        "PM_AI_PROCESSING_PLAN_PATH",
        "PM_AI_EXCLUDE_RULES_JSON",
        "PM_AI_ACTUAL_DETAIL_SOURCE_DIR",
        "PM_AI_ACTUAL_DETAIL_WORKBOOK",
        "PM_AI_WORKSPACE"
    )) {
    if ($ui.ContainsKey($key) -and $ui[$key]) {
        Set-Item -Path "env:$key" -Value $ui[$key]
    }
}

if (-not $MasterWorkbook) {
    $MasterWorkbook = $env:PM_AI_MASTER_WORKBOOK
}
if (-not $MasterWorkbook -and $ui.ContainsKey("PM_AI_MASTER_WORKBOOK")) {
    $MasterWorkbook = $ui["PM_AI_MASTER_WORKBOOK"]
}
if (-not $MasterWorkbook) {
    $localMaster = Join-Path $RepoRoot "master.xlsm"
    if (Test-Path -LiteralPath $localMaster) {
        $MasterWorkbook = $localMaster
    }
}

if (-not $MasterWorkbook) {
    Write-Error "PM_AI_MASTER_WORKBOOK is not set and master.xlsm was not found under RepoRoot=$RepoRoot"
    exit 2
}

$env:PM_AI_MASTER_WORKBOOK = $MasterWorkbook

Set-Location $pyDir
Write-Host "[stage1-headless] repo=$RepoRoot"
Write-Host "[stage1-headless] master=$MasterWorkbook"
Write-Host "[stage1-headless] task_input_dir=$env:PM_AI_TASK_INPUT_SOURCE_DIR"
Write-Host "[stage1-headless] output=$env:PM_AI_PLAN_INPUT_PATH"
Write-Host "[stage1-headless] log=$logOut"

$prevEap = $ErrorActionPreference
$ErrorActionPreference = "Continue"
$pyArgs = @("-3.14", "-X", "utf8", "-u", "task_extract_stage1.py")
$proc = Start-Process -FilePath "py" -ArgumentList $pyArgs -WorkingDirectory $pyDir `
    -RedirectStandardOutput $logOut -RedirectStandardError (Join-Path $RepoRoot "log\stage1_headless_stderr.txt") `
    -NoNewWindow -Wait -PassThru
$exitCode = $proc.ExitCode
$ErrorActionPreference = $prevEap
Write-Host "[stage1-headless] exit=$exitCode"
exit $exitCode
