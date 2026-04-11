param(
  [Parameter(Mandatory = $true)][string]$InputBas,
  [Parameter(Mandatory = $true)][string]$OutDir
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Ensure-Dir([string]$p) {
  if (-not (Test-Path -LiteralPath $p)) { New-Item -ItemType Directory -Path $p | Out-Null }
}

function StrFromCodePoints([int[]]$cps) {
  $sb = New-Object System.Text.StringBuilder
  foreach ($cp in $cps) { [void]$sb.Append([char]$cp) }
  $sb.ToString()
}

# Module names (ASCII-only script; build Japanese at runtime)
$M_Common = StrFromCodePoints 0x5171,0x901A,0x5B9A,0x7FA9
$M_Stage  = StrFromCodePoints 0x6BB5,0x968E,0x5B9F,0x884C,0x5236,0x5FA1
$M_Splash = StrFromCodePoints 0x30B9,0x30D7,0x30E9,0x30C3,0x30B7,0x30E5,0x8868,0x793A
$M_Sound  = StrFromCodePoints 0x30B5,0x30A6,0x30F3,0x30C9,0x5236,0x5FA1
$M_Gemini = "Gemini" + (StrFromCodePoints 0x9023,0x643A)
$M_Env    = StrFromCodePoints 0x74B0,0x5883,0x30BB,0x30C3,0x30C8,0x30A2,0x30C3,0x30D7
$M_Biz    = StrFromCodePoints 0x696D,0x52D9,0x30ED,0x30B8,0x30C3,0x30AF
$M_TextIO = StrFromCodePoints 0x6587,0x5B57,0x5217,0x5165,0x51FA,0x529B,0x5171,0x901A
$M_Find   = StrFromCodePoints 0x30D5,0x30A1,0x30A4,0x30EB,0x63A2,0x7D22
$M_Font   = StrFromCodePoints 0x30D5,0x30A9,0x30F3,0x30C8,0x7BA1,0x7406
$M_Short  = StrFromCodePoints 0x8D77,0x52D5,0x30B7,0x30E7,0x30FC,0x30C8,0x30AB,0x30C3,0x30C8

Ensure-Dir $OutDir

$lines = Get-Content -LiteralPath $InputBas -Encoding Default

$procStartRegex = '^\s*(Public|Private|Friend)?\s*(Sub|Function)\s+'
$firstProcIdx = ($lines | Select-String -Pattern $procStartRegex | Select-Object -First 1).LineNumber
if (-not $firstProcIdx) { throw "No procedures found in $InputBas" }
$firstProcIdx0 = $firstProcIdx - 1
$header = $lines[0..($firstProcIdx0 - 1)]

# Convert top-level declarations to Public so other modules can reference constants/declares.
$header2 = foreach ($l in $header) {
  $t = $l
  $t = $t -replace '^\s*Private\s+Type\b', 'Public Type'
  $t = $t -replace '^\s*Private\s+Const\b', 'Public Const'
  $t = $t -replace '^\s*Private\s+Declare\b', 'Public Declare'
  $t = $t -replace '^\s*Private\s+Declare\s+PtrSafe\b', 'Public Declare PtrSafe'
  $t = $t -replace '^\s*Private\s+m_', 'Public m_'
  $t
}

function Category-ForProcName([string]$name, [string[]]$body) {
  if ($name -match '^(GeminiReadUtf8FileViaTempCopy|GeminiReadUtf8File|GeminiWriteUtf8File|GeminiJsonStringEscape)$') { return $M_TextIO }
  if ($name -match '^MacroSplash_') { return $M_Splash }
  if ($name -match '^(PlayFinishSound|MacroStartBgm_)') { return $M_Sound }
  if ($name -match '^(Gemini|設定_Gemini|アニメ付き_Gemini|メインシート_Gemini|LOG_AI)') { return $M_Gemini }
  if ($name -match '^(InstallComponents|SetupEnvironment)') { return $M_Env }
  if ($name -match '^ShortcutMainSheet_') { return $M_Short }
  if ($name -match '(フォント|ApplyFont|FontPick|BIZ_UDP)') { return $M_Font }
  if ($name -match '^(WriteTempCmdFile|RunTempCmd|RunCmdFile|BuildStage|ReadStageVbaExitCode|StageVbaExitCode|XwRunConsoleRunner|ParseStage12|Stage12|Stage1Sync)') { return $M_Stage }

  $joined = ($body -join "`n")
  if ($joined -match '\b(CreateObject\("Scripting\.FileSystemObject"\)|Dir\(|FileCopy\b|Kill\b|MkDir\b|GetFolder\(|FolderExists\()') { return $M_Find }

  return $M_Biz
}

function Parse-Procs([string[]]$allLines, [int]$startIdx0) {
  $procs = @()
  $i = $startIdx0
  while ($i -lt $allLines.Count) {
    while ($i -lt $allLines.Count -and $allLines[$i] -notmatch $procStartRegex) { $i++ }
    if ($i -ge $allLines.Count) { break }

    $start = $i
    $sig = $allLines[$i]
    $m = [regex]::Match($sig, '^\s*(Public|Private|Friend)?\s*(Sub|Function)\s+([^\(\s]+)')
    if (-not $m.Success) { throw "Failed to parse signature at line $($i + 1): $sig" }
    $name = $m.Groups[3].Value

    $i++
    while ($i -lt $allLines.Count -and $allLines[$i] -notmatch '^\s*End\s+(Sub|Function)\b') { $i++ }
    if ($i -ge $allLines.Count) { throw "Missing End Sub/Function for $name" }
    $end = $i
    $i++

    $body = $allLines[$start..$end]
    $procs += [pscustomobject]@{ Name = $name; Lines = $body }
  }
  return $procs
}

$procs = Parse-Procs -allLines $lines -startIdx0 $firstProcIdx0

$modules = [ordered]@{
  $M_Common = @()
  $M_Stage  = @()
  $M_Splash = @()
  $M_Sound  = @()
  $M_Gemini = @()
  $M_Env    = @()
  $M_Biz    = @()
  $M_TextIO = @()
  $M_Find   = @()
  $M_Font   = @()
  $M_Short  = @()
}

foreach ($p in $procs) {
  $cat = Category-ForProcName -name $p.Name -body $p.Lines
  $modules[$cat] += $p.Lines
  $modules[$cat] += ""
}

function Write-Bas([string]$modName, [string[]]$bodyLines) {
  $outPath = Join-Path $OutDir ($modName + ".bas")
  $out = @()
  $out += 'Attribute VB_Name = "' + $modName + '"'
  $out += "Option Explicit"
  $out += ""
  if ($modName -eq $M_Common) {
    $out += $header2
    $out += ""
  }
  $out += $bodyLines
  # Shift-JIS (ANSI) output for Japanese Windows.
  $out | Set-Content -LiteralPath $outPath -Encoding Default
}

Write-Bas -modName $M_Common -bodyLines @()
foreach ($k in $modules.Keys) {
  if ($k -eq $M_Common) { continue }
  Write-Bas -modName $k -bodyLines $modules[$k]
}

Write-Host "Wrote modules to: $OutDir"
