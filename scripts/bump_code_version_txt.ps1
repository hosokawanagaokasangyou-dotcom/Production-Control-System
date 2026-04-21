# Bumps first line of code/version.txt by +0.1 (decimal). UTF-8 no BOM output.
# Setup: git config core.hooksPath .githooks
# Skip: $env:SKIP_BUMP_CODE_VERSION = "1"

$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent $PSScriptRoot
$target = Join-Path $repoRoot 'code\version.txt'

if (-not (Test-Path -LiteralPath $target)) {
    Write-Error "Missing file: $target"
    exit 1
}

$utf8 = New-Object System.Text.UTF8Encoding $false
try {
    $raw = @([System.IO.File]::ReadAllLines($target, $utf8))
} catch {
    Write-Error "Cannot read version.txt: $_"
    exit 1
}
if ($null -eq $raw -or $raw.Count -eq 0 -or ($raw.Count -eq 1 -and [string]::IsNullOrWhiteSpace($raw[0]))) {
    Write-Error 'version.txt is empty or whitespace only'
    exit 1
}

$inv = [cultureinfo]::InvariantCulture
$cur = $raw[0].Trim()
$v = [decimal]0
if (-not [decimal]::TryParse($cur, [System.Globalization.NumberStyles]::Number, $inv, [ref]$v)) {
    Write-Error "First line is not a number: $cur"
    exit 1
}

$newV = [decimal]::Round($v + [decimal]0.1, 1)
$raw[0] = $newV.ToString("0.0", $inv)

$utf8NoBom = New-Object System.Text.UTF8Encoding $false
[System.IO.File]::WriteAllLines($target, $raw, $utf8NoBom)

Write-Host "version.txt: $cur -> $($raw[0])"
