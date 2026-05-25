#Requires -Version 5.1
<#
.SYNOPSIS
  Windows 向け: compile 後に検証してから JavaFX デスクトップを起動する。

.DESCRIPTION
  同一フォルダの mvnw.cmd で compile → verify-pm-ai-build.ps1 → exec:exec@pm-ai-desktop を実行する。
  pom の JVM オプション（-Xms/-Xmx 等）も適用される。

  重要: このフォルダから実行すること（相対パス前提）:
    .\run-pm-ai-desktop.ps1

  mvnw javafx:run も pm-ai-mvn-route.cmd 経由で同じ検証を行う。
  ClassNotFound 時は .\run-pm-ai-desktop.ps1 を推奨。

.PARAMETER MaxHeap
  Maven プロパティ jvm.max.heap（既定 4g。例: 2g, 4g, 8g）。

.PARAMETER MonitorIntervalSec
  ヒープ監視間隔（秒）。-1 で無効。0 で環境変数 PM_AI_JVM_MEMORY_MONITOR_SEC を使用。

.EXAMPLE
  .\run-pm-ai-desktop.ps1
.EXAMPLE
  .\run-pm-ai-desktop.ps1 -MaxHeap 4g -MonitorIntervalSec 30
#>
param(
    [string] $MaxHeap = "4g",
    [int] $MonitorIntervalSec = -1
)

$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

if ($MonitorIntervalSec -ge 0) {
    $env:PM_AI_JVM_MEMORY_MONITOR_SEC = "$MonitorIntervalSec"
}

$commonArgs = @("-q", "-Djvm.max.heap=$MaxHeap")
$verifyScript = Join-Path $PSScriptRoot "verify-pm-ai-build.ps1"

function Invoke-PmAiBuildVerify {
    & powershell -NoProfile -ExecutionPolicy Bypass -File $verifyScript -ProjectRoot $PSScriptRoot
    return $LASTEXITCODE
}

& "$PSScriptRoot\mvnw.cmd" @commonArgs @("compile")
if ($LASTEXITCODE -ne 0) {
    exit $LASTEXITCODE
}

if ((Invoke-PmAiBuildVerify) -ne 0) {
    Write-Warning "compile 後に必須出力が不足しています。clean compile を実行します..."
    & "$PSScriptRoot\mvnw.cmd" @commonArgs @("clean", "compile")
    if ($LASTEXITCODE -ne 0) {
        exit $LASTEXITCODE
    }
    if ((Invoke-PmAiBuildVerify) -ne 0) {
        Write-Error @"
compile 後も必須 .class / CSS が見つかりません。
対処:
  1. 実行中の Java / Maven プロセスを終了
  2. .\mvnw.cmd clean compile
  3. .\run-pm-ai-desktop.ps1
"@
        exit 1
    }
}

& "$PSScriptRoot\mvnw.cmd" @commonArgs @("validate", "exec:exec@pm-ai-desktop")
exit $LASTEXITCODE
