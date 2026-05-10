#Requires -Version 5.1
<#
.SYNOPSIS
  マニュアル用パイプライン（データ準備・プレースホルダ置換・HTML 生成）を 1 回実行する。

.DESCRIPTION
  リポジトリ直下から実行する。
    .\Publish-Manual.ps1

  事前に Python 3.10+ と依存を用意する。
    pip install -r manual/requirements.txt

  画面キャプチャは手動で manual\src\images に PNG を置いてから実行する（自動連続撮影は行わない）。

.PARAMETER Manifest
  既定: manual\pipeline-manifest.yaml

.PARAMETER SkipDataPrep / SkipInject / SkipHtml
  各フェーズをスキップするときに指定。
#>
param(
    [string] $Manifest = "manual\pipeline-manifest.yaml",
    [switch] $SkipDataPrep,
    [switch] $SkipInject,
    [switch] $SkipHtml
)

$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

function Resolve-Python {
    $venvWin = Join-Path $PSScriptRoot "manual\.venv\Scripts\python.exe"
    $venvNix = Join-Path $PSScriptRoot "manual/.venv/bin/python"
    if (Test-Path -LiteralPath $venvWin) { return $venvWin }
    if (Test-Path -LiteralPath $venvNix) { return $venvNix }
    foreach ($cmd in @("python", "py")) {
        try {
            & $cmd -c "import sys; assert sys.version_info >= (3, 10); sys.exit(0)" 2>$null
            if ($LASTEXITCODE -eq 0) { return $cmd }
        } catch {}
    }
    return $null
}

$py = Resolve-Python
if (-not $py) {
    Write-Error "Python 3.10+ が見つかりません。仮想環境か pip install -r manual/requirements.txt を実行してください。"
}

$pubArgs = @(
    "scripts\manual_publish.py",
    "--manifest", $Manifest
)
if ($SkipDataPrep) { $pubArgs += "--skip-data-prep" }
if ($SkipInject) { $pubArgs += "--skip-inject" }
if ($SkipHtml) { $pubArgs += "--skip-html" }

Write-Host "Using $py $($pubArgs -join ' ')"
& $py @pubArgs
exit $LASTEXITCODE
