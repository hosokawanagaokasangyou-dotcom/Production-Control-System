# Canonical implementation: repository-root fast_package_app.ps1 (run from repo root recommended).
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
$forwardArgs = @($args)
if ($forwardArgs.Count -eq 0) {
    & (Join-Path $repoRoot 'fast_package_app.ps1')
}
else {
    & (Join-Path $repoRoot 'fast_package_app.ps1') @forwardArgs
}
