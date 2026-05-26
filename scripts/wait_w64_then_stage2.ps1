# Wait until plan_input_tasks.xlsx contains W6-4, then run stage2 headless.
param(
    [string]$RepoRoot = "",
    [int]$MaxWaitSec = 300,
    [int]$PollSec = 3
)

if (-not $RepoRoot) {
    $RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
}

$planPath = Join-Path $RepoRoot "output\plan_input_tasks.xlsx"
$headless = Join-Path $RepoRoot "scripts\run_stage2_headless.ps1"
$pyCheck = @"
from openpyxl import load_workbook
p=r'$($planPath.Replace("'","''"))'
wb=load_workbook(p, read_only=True, data_only=True)
ws=wb.active
n=sum(1 for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True)
        if any('W6-4' in str(c) for c in row if c is not None))
print(n)
wb.close()
"@

$deadline = (Get-Date).AddSeconds($MaxWaitSec)
while ((Get-Date) -lt $deadline) {
    if (-not (Test-Path -LiteralPath $planPath)) {
        Write-Host "[wait-w64] missing $planPath"
        Start-Sleep -Seconds $PollSec
        continue
    }
    $count = [int](& py -3.14 -X utf8 -c $pyCheck 2>$null)
    Write-Host "[wait-w64] W6-4 rows=$count size=$((Get-Item -LiteralPath $planPath).Length)"
    if ($count -gt 0) {
        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $headless -RepoRoot $RepoRoot
        exit $LASTEXITCODE
    }
    Start-Sleep -Seconds $PollSec
}

Write-Error "W6-4 not found in plan_input_tasks.xlsx within ${MaxWaitSec}s. Save the workbook and retry."
exit 2
