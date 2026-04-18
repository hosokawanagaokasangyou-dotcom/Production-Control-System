# -*- coding: utf-8 -*-
#Requires -Version 5.1
<#
.SYNOPSIS
  起動中の Excel を保存確認付きで終了し、setup_environment.py を実行後、
  終了前に開いていたブック（保存済みパスのみ）を開き直す。

.NOTES
  - xlwings のアドイン配置・pip インストールで Excel / xlwings.xlam がロックされる問題の回避用。
  - 未保存の新規ブックは FullName がディスクパスにならないため、再オープン対象に含めません。
#>
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

function Test-IsSavedWorkbookPath {
    param([string]$FullName)
    if ([string]::IsNullOrWhiteSpace($FullName)) { return $false }
    $t = $FullName.Trim()
    # ドライブレターまたは UNC
    if ($t.Length -ge 2 -and $t[1] -eq ':') { return $true }
    if ($t.StartsWith('\\')) { return $true }
    return $false
}

function Get-ExcelWorkbookPathsFromRunningInstances {
    $paths = New-Object System.Collections.ArrayList
    $seen = @{}

    while ($true) {
        $xl = $null
        try {
            $xl = [System.Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
        } catch {
            break
        }

        try {
            foreach ($wb in $xl.Workbooks) {
                try {
                    $full = $wb.FullName
                    if (-not (Test-IsSavedWorkbookPath $full)) { continue }
                    if (-not (Test-Path -LiteralPath $full)) { continue }
                    $key = $full.ToLowerInvariant()
                    if (-not $seen.ContainsKey($key)) {
                        [void]$seen.Add($key, $null)
                        [void]$paths.Add($full)
                    }
                } catch {
                    # 無視（一時的に COM エラーになり得る）
                }
            }
            $xl.DisplayAlerts = $true
            $xl.Quit()
        } finally {
            if ($null -ne $xl) {
                [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($xl)
                $xl = $null
                [GC]::Collect()
                [GC]::WaitForPendingFinalizers()
            }
        }
        Start-Sleep -Milliseconds 500
    }

    return ,$paths.ToArray()
}

function Wait-ExcelProcessesGone {
    param(
        [int]$TimeoutSec = 1800
    )
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ($true) {
        $procs = Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue
        if (-not $procs) { return $true }
        if ((Get-Date) -gt $deadline) {
            return $false
        }
        Start-Sleep -Seconds 1
    }
}

Write-Host ''
Write-Host '==== 工程管理AI: 環境セットアップ（Excel 終了 → pip / xlwings） ====' -ForegroundColor Cyan
Write-Host ''

$pythonDir = $PSScriptRoot
$setupPy = Join-Path $pythonDir 'setup_environment.py'

if (-not (Test-Path -LiteralPath $setupPy)) {
    Write-Host "見つかりません: $setupPy" -ForegroundColor Red
    Write-Host 'Enter で終了...'
    Read-Host | Out-Null
    exit 1
}

$excelProcBefore = @(Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue)
$savedPaths = @()

if ($excelProcBefore.Count -gt 0) {
    Write-Host '[1/3] 開いている Excel を終了します（未保存はダイアログで保存するか選択してください）。' -ForegroundColor Yellow
    $savedPaths = @(Get-ExcelWorkbookPathsFromRunningInstances)

    if (Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue) {
        Write-Host '       Excel プロセスの完全終了を待っています…' -ForegroundColor DarkGray
        if (-not (Wait-ExcelProcessesGone)) {
            Write-Host ''
            Write-Host 'Excel がタイムアウトまで終了しませんでした。セットアップを中止します。' -ForegroundColor Red
            Write-Host 'Enter で終了...'
            Read-Host | Out-Null
            exit 2
        }
    }
    Write-Host '       Excel は終了しました。' -ForegroundColor Green
} else {
    Write-Host '[1/3] Excel は起動していません（スキップ）。' -ForegroundColor DarkGray
}

Write-Host ''
Write-Host '[2/3] setup_environment.py を実行します（pip / xlwings アドイン）...' -ForegroundColor Yellow
Push-Location $pythonDir
try {
    & py -3.14 -X utf8 -u $setupPy
    $setupCode = $LASTEXITCODE
} catch {
    Write-Host "実行に失敗しました: $_" -ForegroundColor Red
    $setupCode = 1
} finally {
    Pop-Location
}

Write-Host ''

if ($setupCode -ne 0) {
    Write-Host "[2/3] セットアップが終了コード $setupCode で終了しました。ログを確認してください。" -ForegroundColor Red
    Write-Host 'Enter で終了...'
    Read-Host | Out-Null
    exit $setupCode
}

Write-Host '[3/3] 終了前に記録したブックを開き直します…' -ForegroundColor Yellow
if ($savedPaths.Count -eq 0) {
    Write-Host '       再オープン対象の保存済みパスがありません（未保存ブックのみ等）。' -ForegroundColor DarkGray
} else {
    $ordered = $savedPaths | Sort-Object
    foreach ($p in $ordered) {
        if (Test-Path -LiteralPath $p) {
            Write-Host "       開く: $p"
            Start-Process -FilePath $p
            Start-Sleep -Milliseconds 400
        } else {
            Write-Host "       スキップ（見つからない）: $p" -ForegroundColor DarkGray
        }
    }
}

Write-Host ''
Write-Host '手順が完了しました。' -ForegroundColor Green
Write-Host ''
