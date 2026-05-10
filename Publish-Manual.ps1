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

  WSL と Windows で同じ manual\.venv を共用しないこと（pyvenv.cfg の home が別 OS のパスになり失敗する）。

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

function Test-Python310Plus {
    param([string]$Exe)
    try {
        & $Exe -c "import sys; assert sys.version_info >= (3, 10); sys.exit(0)" 2>$null
        return ($LASTEXITCODE -eq 0)
    } catch {
        return $false
    }
}

function Test-VenvPyvenvCfgLooksForeignOnWindows {
    param([string]$RepoRoot)
    $cfgPath = Join-Path $RepoRoot "manual\.venv\pyvenv.cfg"
    if (-not (Test-Path -LiteralPath $cfgPath)) {
        return $false
    }
    $txt = Get-Content -LiteralPath $cfgPath -Raw
    if ($txt -match '(?m)^\s*home\s*=\s*/') {
        return $true
    }
    return $false
}

function Resolve-Python {
    $repo = $PSScriptRoot
    $venvWin = Join-Path $repo "manual\.venv\Scripts\python.exe"
    $venvNix = Join-Path $repo "manual/.venv/bin/python"

    $onWindows = ($env:OS -eq 'Windows_NT')

    if ($onWindows -and (Test-Path -LiteralPath $venvWin)) {
        if (Test-VenvPyvenvCfgLooksForeignOnWindows -RepoRoot $repo) {
            $warnMsg = @(
                "manual\.venv が WSL/Linux 向けです（pyvenv.cfg の home が / で始まります）。",
                "このフォルダを削除し、PowerShell で Windows 用 venv を作り直してください:",
                "  Remove-Item -Recurse -Force manual\.venv",
                "  python -m venv manual\.venv",
                "  manual\.venv\Scripts\pip install -r manual\requirements.txt",
                "いったん PATH 上の Python で続行します。"
            ) -join [Environment]::NewLine
            Write-Warning $warnMsg
        } elseif (Test-Python310Plus -Exe $venvWin) {
            return $venvWin
        } else {
            Write-Warning "manual\.venv\Scripts\python.exe が実行できません。venv を作り直すか PATH の Python を確認してください。"
        }
    }

    if ((-not $onWindows) -and (Test-Path -LiteralPath $venvNix)) {
        if (Test-Python310Plus -Exe $venvNix) {
            return $venvNix
        }
        Write-Warning "manual/.venv/bin/python が実行できません。"
    }

    foreach ($cmd in @("python", "py")) {
        try {
            if (Test-Python310Plus -Exe $cmd) {
                return $cmd
            }
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
