#Requires -Version 5.1
<#
.SYNOPSIS
  target\classes に必須 .class / CSS があるか検証する（Unicode パス対応）。

.DESCRIPTION
  cmd の if not exist は日本語パスで誤判定するため、PowerShell Test-Path を使う。
  mvnw javafx:run（pm-ai-mvn-route.cmd）と run-pm-ai-desktop.ps1 から共用。

.EXIT CODES
  0 = OK
  1 = 不足あり
#>
param(
    [string] $ProjectRoot = $PSScriptRoot
)

$ErrorActionPreference = "Stop"
$ProjectRoot = $ProjectRoot.Trim().Trim('"').TrimEnd('\', '/')
if ([string]::IsNullOrWhiteSpace($ProjectRoot)) {
    $ProjectRoot = $PSScriptRoot
}

$requiredRelativePaths = @(
    "target\classes\jp\co\pm\ai\desktop\PmAiFxApp.class",
    "target\classes\jp\co\pm\ai\desktop\io\SkillsSheetMemberReader.class",
    "target\classes\jp\co\pm\ai\desktop\io\PlanInputTabularIo.class",
    "target\classes\jp\co\pm\ai\desktop\io\NetworkSourceFileReloadCache.class",
    "target\classes\jp\co\pm\ai\desktop\dispatch\ResultDispatchInteractiveConsolidator.class",
    "target\classes\jp\co\pm\ai\desktop\ui\EquipmentGraphicGanttPane.class",
    # record / inner: incremental compile on Windows can drop these while outer .class remains
    "target\classes\jp\co\pm\ai\desktop\ui\EquipmentGraphicGanttPane`$VerticalScrollBand.class",
    "target\classes\jp\co\pm\ai\desktop\ui\EquipmentGraphicGanttPane`$LazyBadgeLayoutRequest.class",
    "target\classes\jp\co\pm\ai\desktop\ui\SpreadsheetMultiColumnFilterCoordinator.class",
    "target\classes\jp\co\pm\ai\desktop\ui\TableColumnOrderPersistence.class",
    "target\classes\jp\co\pm\ai\desktop\reconciliation\ReconciliationApp.class",
    # comboRecord button cell (NoClassDefFoundError: ReconciliationApp$2)
    "target\classes\jp\co\pm\ai\desktop\reconciliation\ReconciliationApp`$2.class",
    "target\classes\jp\co\pm\ai\desktop\css\pm-ai-desktop.css",
    "target\classes\jp\co\pm\ai\desktop\css\theme-midnight-blue.css"
)

$missing = @()
foreach ($rel in $requiredRelativePaths) {
    $path = Join-Path $ProjectRoot $rel
    if (-not (Test-Path -LiteralPath $path)) {
        $missing += $rel
    }
}

if ($missing.Count -eq 0) {
    exit 0
}

Write-Host "[pm-ai-desktop] missing build outputs:"
foreach ($rel in $missing) {
    Write-Host "  $rel"
}
exit 1
