#Requires -Version 5.1
<#
.SYNOPSIS
  Windows 本番想定: PowerShell から工程管理 JavaFX デスクトップを起動する。

.DESCRIPTION
  同階層の mvnw.cmd で javafx:run します。pom.xml の JVM オプション（-Xms/-Xmx、OOM 時ヒープダンプ等）がそのまま適用されます。

  重要: このフォルダで実行するときはパスの先頭に .\ を付けます。
    .\run-pm-ai-desktop.ps1
  リポジトリ直下からなら:
    .\code_java\run-pm-ai-desktop.ps1
  （run-pm-ai-desktop.ps1 だけでは実行できません。）

  ヒープ監視（stderr に定期サンプル）を有効にする例:
    .\run-pm-ai-desktop.ps1 -MonitorIntervalSec 60

  既定では環境変数 PM_AI_JVM_MEMORY_MONITOR_SEC は上書きしません（-1）。
  0 以上を指定したときだけ、そのセッション用に設定して子 JVM に継承します。

.PARAMETER MaxHeap
  Maven プロパティ jvm.max.heap（例: 2g, 4g）。大きいブック時は 4g 等。

.PARAMETER MonitorIntervalSec
  ヒープ監視の間隔（秒）。-1 なら環境変数を触らない。0 以上で PM_AI_JVM_MEMORY_MONITOR_SEC をその値に設定。

.EXAMPLE
  .\run-pm-ai-desktop.ps1
.EXAMPLE
  .\run-pm-ai-desktop.ps1 -MaxHeap 4g -MonitorIntervalSec 30
#>
param(
    [string] $MaxHeap = "2g",
    [int] $MonitorIntervalSec = -1
)

$ErrorActionPreference = "Stop"
Set-Location -LiteralPath $PSScriptRoot

if ($MonitorIntervalSec -ge 0) {
    $env:PM_AI_JVM_MEMORY_MONITOR_SEC = "$MonitorIntervalSec"
}

$mvnArgs = @(
    "-q",
    "-Djvm.max.heap=$MaxHeap",
    "javafx:run"
)

& "$PSScriptRoot\mvnw.cmd" @mvnArgs
exit $LASTEXITCODE
