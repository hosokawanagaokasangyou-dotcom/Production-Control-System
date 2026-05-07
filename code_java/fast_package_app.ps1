# Canonical implementation: repository-root fast_package_app.ps1 (run from repo root recommended).
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent $PSScriptRoot
& (Join-Path $repoRoot 'fast_package_app.ps1') @args
