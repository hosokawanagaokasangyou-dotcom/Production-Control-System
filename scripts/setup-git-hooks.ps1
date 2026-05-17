# One-time: enable pre-commit (version.txt bump + staged *.md UTF-8 check).
# UTF-8 BOM recommended when editing from Windows PowerShell 5.1.
$ErrorActionPreference = 'Stop'
$Root = if ($PSScriptRoot) { Split-Path -Parent $PSScriptRoot } else { Get-Location }
Set-Location -LiteralPath $Root
git config core.hooksPath scripts/git-hooks
$hooks = @(
    (Join-Path $Root 'scripts\git-hooks\pre-commit'),
    (Join-Path $Root 'scripts\resolve_python3.sh'),
    (Join-Path $Root 'scripts\run_python3.sh')
)
foreach ($h in $hooks) {
    if (Test-Path -LiteralPath $h) {
        icacls $h /grant Everyone:RX 2>$null | Out-Null
    }
}
$py = & 'C:\Program Files\Git\bin\bash.exe' -lc "cd '$($Root -replace '\\','/')' && scripts/resolve_python3.sh" 2>&1
Write-Host "core.hooksPath = scripts/git-hooks"
Write-Host "Python for hooks: $py"
Write-Host 'Test: git commit (version.txt should +0.01 in the same commit).'
