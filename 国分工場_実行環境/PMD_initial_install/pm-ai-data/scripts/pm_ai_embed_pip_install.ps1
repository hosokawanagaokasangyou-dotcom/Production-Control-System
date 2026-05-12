# Installs code/python/requirements.txt into the bundled Python embed (or any python.exe).
# Run on Windows when dispatch trial fails with "No module named 'pydantic'" etc.
#
# Example:
#   .\scripts\pm_ai_embed_pip_install.ps1
#   .\scripts\pm_ai_embed_pip_install.ps1 -PythonExe 'D:\pm-ai-data\runtime\python-embed\python.exe'
#
[CmdletBinding()]
param(
    [string]$PythonExe = ''
)

$ErrorActionPreference = 'Stop'

$here = $PSScriptRoot
if (-not $here) {
    $here = Split-Path -Parent $MyInvocation.MyCommand.Path
}
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $here '..')).Path
$req = Join-Path $repoRoot 'code\python\requirements.txt'
if (-not (Test-Path -LiteralPath $req)) {
    throw "requirements.txt not found: $req"
}

if ([string]::IsNullOrWhiteSpace($PythonExe)) {
    $PythonExe = 'C:\pm-ai-data\runtime\python-embed\python.exe'
}

if (-not (Test-Path -LiteralPath $PythonExe)) {
    throw @"
Python not found: $PythonExe
Pass -PythonExe to your embed python.exe (e.g. pm-ai-data\runtime\python-embed\python.exe).
"@
}

Write-Host "--- pip upgrade ---" -ForegroundColor Cyan
& $PythonExe -m pip install --upgrade pip --no-warn-script-location
if ($LASTEXITCODE -ne 0) {
    throw 'pip upgrade failed.'
}

Write-Host "--- pip install -r requirements.txt ---" -ForegroundColor Cyan
& $PythonExe -m pip install -r $req --no-warn-script-location
if ($LASTEXITCODE -ne 0) {
    throw 'pip install -r requirements.txt failed.'
}

Write-Host 'Done.' -ForegroundColor Green
