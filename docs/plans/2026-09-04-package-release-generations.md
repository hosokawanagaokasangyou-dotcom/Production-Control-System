# パッケージ成果物の世代退避 Implementation Plan

> **For implementers:** Execute this plan task-by-task. Use the `test-driven-development` skill for each task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** `fast_package_app.ps1` の Step 8 で、上書き前のリリース成果物を `previous\v{version}\` に退避し、保持世代数で剪定する。

**Architecture:** Step 8 直前に移動ベースの退避関数を呼び、ルート固定名は従来どおり再生成する。退避・剪定ロジックは `scripts/package_release_archive.ps1` に切り出し、一時ディレクトリでの検証スクリプトで確認する。

**Tech Stack:** Windows PowerShell 5.1、`fast_package_app.ps1`

**仕様正本:** `docs/specs/2026-09-04-version-management-downgrade-design.md`（パッケージ節）

**依存:** なし（本プラン単独で完了可能）。続けてタブ実装は `docs/plans/2026-09-04-version-management-tab.md`。

---

## ファイル構成

| ファイル | 責任 |
|----------|------|
| `scripts/package_release_archive.ps1` | `Move-PmAiReleaseArtifactsToPrevious` / `Remove-PmAiExcessPreviousGenerations` / 版フォルダ名解決 |
| `scripts/test_package_release_archive.ps1` | 一時 dir での退避・衝突・剪定の検証（exit 0/1） |
| `fast_package_app.ps1` | パラメータ・Step 8 配線・README 文言 |

---

### Task 1: 退避・剪定ヘルパ（失敗する検証から）

**Files:**
- Create: `scripts/package_release_archive.ps1`
- Create: `scripts/test_package_release_archive.ps1`

- [ ] **Step 1: Write the failing verification script**

`scripts/test_package_release_archive.ps1`:

```powershell
# UTF-8. Exit 0 on success. Requires package_release_archive.ps1 functions.
$ErrorActionPreference = 'Stop'
$here = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
$lib = Join-Path $here 'package_release_archive.ps1'
if (-not (Test-Path -LiteralPath $lib)) {
    Write-Error "Missing $lib"
    exit 1
}
. $lib

$tmp = Join-Path $env:TEMP ('pm-ai-archive-test-' + [Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp | Out-Null
try {
    $ver = Join-Path $tmp 'version.txt'
    Set-Content -LiteralPath $ver -Value "1.20.0`n" -Encoding utf8
    Set-Content -LiteralPath (Join-Path $tmp 'PMD_initial_install.zip') -Value 'i' -Encoding ascii
    Set-Content -LiteralPath (Join-Path $tmp 'PMD_version_upgrade.zip') -Value 'u' -Encoding ascii
    Set-Content -LiteralPath (Join-Path $tmp 'build-manifest.json') -Value '{}' -Encoding utf8

    $dest = Move-PmAiReleaseArtifactsToPrevious -ReleaseRoot $tmp
    if (-not $dest) { throw 'Move returned empty' }
    $leaf = Split-Path -Leaf $dest
    if ($leaf -ne 'v1.20.0') { throw "Expected v1.20.0 folder, got $leaf" }
    if (Test-Path -LiteralPath $ver) { throw 'version.txt should have been moved' }
    if (-not (Test-Path -LiteralPath (Join-Path $dest 'PMD_version_upgrade.zip'))) {
        throw 'upgrade zip missing in previous'
    }

    # collision: recreate same version artifacts, move again
    Set-Content -LiteralPath $ver -Value "1.20.0`n" -Encoding utf8
    Set-Content -LiteralPath (Join-Path $tmp 'PMD_version_upgrade.zip') -Value 'u2' -Encoding ascii
    $dest2 = Move-PmAiReleaseArtifactsToPrevious -ReleaseRoot $tmp
    $leaf2 = Split-Path -Leaf $dest2
    if ($leaf2 -notmatch '^v1\.20\.0_\d{8}-\d{6}$') {
        throw "Expected timestamped collision folder, got $leaf2"
    }

    # prune: create 3 fake gens, keep 2
    $prev = Join-Path $tmp 'previous'
    1..3 | ForEach-Object {
        $d = Join-Path $prev ("v1.0.$_")
        New-Item -ItemType Directory -Path $d -Force | Out-Null
        Start-Sleep -Milliseconds 50
        Set-Content -LiteralPath (Join-Path $d 'marker.txt') -Value $_ -Encoding ascii
    }
    Remove-PmAiExcessPreviousGenerations -ReleaseRoot $tmp -KeepGenerations 2
    $left = @(Get-ChildItem -LiteralPath $prev -Directory | Sort-Object LastWriteTime -Descending)
    if ($left.Count -ne 2) { throw "Expected 2 gens after prune, got $($left.Count)" }

    Remove-PmAiExcessPreviousGenerations -ReleaseRoot $tmp -KeepGenerations 0
    $left0 = @(Get-ChildItem -LiteralPath $prev -Directory)
    if ($left0.Count -ne 2) { throw 'KeepGenerations 0 must not prune' }

    Write-Host 'OK: package_release_archive tests passed'
    exit 0
}
finally {
    Remove-Item -Recurse -Force -LiteralPath $tmp -ErrorAction SilentlyContinue
}
```

- [ ] **Step 2: Run test to verify it fails**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\test_package_release_archive.ps1
```

Expected: fails（`package_release_archive.ps1` 未存在、または関数未定義）

- [ ] **Step 3: Write minimal implementation**

`scripts/package_release_archive.ps1`:

```powershell
# UTF-8 BOM recommended for PS 5.1. ASCII-safe function bodies.
$ErrorActionPreference = 'Stop'

function Get-PmAiVersionFolderLeaf {
    param([string]$VersionRaw)
    $digits = ''
    if (-not [string]::IsNullOrWhiteSpace($VersionRaw)) {
        $sb = [System.Text.StringBuilder]::new()
        foreach ($ch in $VersionRaw.Trim().ToCharArray()) {
            if ([char]::IsDigit($ch) -or ($ch -eq '.')) { [void]$sb.Append($ch) }
        }
        $digits = $sb.ToString()
        while ($digits.Contains('..')) { $digits = $digits.Replace('..', '.') }
        if ($digits.StartsWith('.')) { $digits = '0' + $digits }
        if ($digits.EndsWith('.')) { $digits = $digits + '0' }
    }
    if ([string]::IsNullOrWhiteSpace($digits)) {
        return ('vunknown_' + (Get-Date -Format 'yyyyMMdd-HHmmss'))
    }
    return ('v' + $digits)
}

function Move-PmAiReleaseArtifactsToPrevious {
    param(
        [Parameter(Mandatory = $true)][string]$ReleaseRoot
    )
    $names = @(
        'version.txt',
        'PMD_initial_install.zip',
        'PMD_version_upgrade.zip',
        'build-manifest.json'
    )
    $existing = @()
    foreach ($n in $names) {
        $p = Join-Path $ReleaseRoot $n
        if (Test-Path -LiteralPath $p) { $existing += $p }
    }
    if ($existing.Count -eq 0) { return $null }

    $verRaw = ''
    $verPath = Join-Path $ReleaseRoot 'version.txt'
    if (Test-Path -LiteralPath $verPath) {
        $verRaw = (Get-Content -LiteralPath $verPath -Raw -Encoding UTF8).Trim()
        $nl = [regex]::Split($verRaw, "`r?`n")
        if ($nl.Length -gt 0) { $verRaw = [string]$nl[0].Trim() }
    }
    $leaf = Get-PmAiVersionFolderLeaf -VersionRaw $verRaw
    $prevRoot = Join-Path $ReleaseRoot 'previous'
    New-Item -ItemType Directory -Path $prevRoot -Force | Out-Null
    $dest = Join-Path $prevRoot $leaf
    if (Test-Path -LiteralPath $dest) {
        $leaf = $leaf + '_' + (Get-Date -Format 'yyyyMMdd-HHmmss')
        $dest = Join-Path $prevRoot $leaf
    }
    New-Item -ItemType Directory -Path $dest -Force | Out-Null
    foreach ($p in $existing) {
        Move-Item -LiteralPath $p -Destination (Join-Path $dest (Split-Path -Leaf $p)) -Force
    }
    return $dest
}

function Remove-PmAiExcessPreviousGenerations {
    param(
        [Parameter(Mandatory = $true)][string]$ReleaseRoot,
        [Parameter(Mandatory = $true)][int]$KeepGenerations
    )
    if ($KeepGenerations -le 0) { return }
    $prevRoot = Join-Path $ReleaseRoot 'previous'
    if (-not (Test-Path -LiteralPath $prevRoot)) { return }
    $dirs = @(Get-ChildItem -LiteralPath $prevRoot -Directory -ErrorAction SilentlyContinue |
        Sort-Object LastWriteTime -Descending)
    if ($dirs.Count -le $KeepGenerations) { return }
    $toRemove = $dirs | Select-Object -Skip $KeepGenerations
    foreach ($d in $toRemove) {
        try {
            Remove-Item -Recurse -Force -LiteralPath $d.FullName -ErrorAction Stop
        }
        catch {
            Write-Warning ("Failed to prune previous generation: {0} ({1})" -f $d.FullName, $_.Exception.Message)
        }
    }
}
```

- [ ] **Step 4: Run test to verify it passes**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\test_package_release_archive.ps1
```

Expected: `OK: package_release_archive tests passed` / exit 0

- [ ] **Step 5: Commit**（ユーザーがコミット依頼したとき、またはリポジトリ方針に従うとき）

```bash
git add scripts/package_release_archive.ps1 scripts/test_package_release_archive.ps1
git commit -m "$(cat <<'EOF'
feat: パッケージ成果物の previous 世代退避ヘルパを追加

EOF
)"
```

---

### Task 2: `fast_package_app.ps1` に配線

**Files:**
- Modify: `fast_package_app.ps1`（param・Step 8・README 1 行）

- [ ] **Step 1: Add parameters**

`param` ブロックに追加:

```powershell
    # Keep N folders under pm-ai-package-release\previous (0 = unlimited). Env: PM_AI_PACKAGE_KEEP_GENERATIONS
    [int]$KeepGenerations = -1,

    # Do not archive prior release artifacts; delete same names as before.
    [switch]$SkipArchive
```

パラメータ解決（`$ReleaseRoot` 確定後〜Step 8 前のどこか、または Step 8 先頭）:

```powershell
$keepGensResolved = $KeepGenerations
if ($keepGensResolved -lt 0) {
    if (-not [string]::IsNullOrWhiteSpace($env:PM_AI_PACKAGE_KEEP_GENERATIONS)) {
        $parsed = 0
        if ([int]::TryParse($env:PM_AI_PACKAGE_KEEP_GENERATIONS.Trim(), [ref]$parsed)) {
            $keepGensResolved = $parsed
        }
        else {
            $keepGensResolved = 5
        }
    }
    else {
        $keepGensResolved = 5
    }
}
```

ヘッダ Usage コメントに `-KeepGenerations` / `-SkipArchive` を追記。

- [ ] **Step 2: Dot-source archive script near other shared scripts**

`package_workspace_copy.ps1` の `.` の直後:

```powershell
$packageReleaseArchiveScript = Join-Path $WorkspaceRoot 'scripts\package_release_archive.ps1'
if (-not (Test-Path -LiteralPath $packageReleaseArchiveScript)) {
    throw "Missing package_release_archive.ps1: $packageReleaseArchiveScript"
}
. (Resolve-Path -LiteralPath $packageReleaseArchiveScript).Path
```

- [ ] **Step 3: Replace Step 8 delete-only block**

現状の「Removing existing release artifacts」ループを次に置き換え:

```powershell
Write-Host "--- Step 8: release version.txt + portable ZIPs ---" -ForegroundColor Cyan
$zipInitial = Join-Path $ReleaseRoot ($BundleInitialName + '.zip')
$zipUpgrade = Join-Path $ReleaseRoot ($BundleUpgradeName + '.zip')
$releaseVersionTxt = Join-Path $ReleaseRoot 'version.txt'

if (-not $SkipArchive) {
    Write-Host "Archiving prior release artifacts to previous\ (KeepGenerations=$keepGensResolved)" -ForegroundColor DarkGray
    $archivedTo = Move-PmAiReleaseArtifactsToPrevious -ReleaseRoot $ReleaseRoot
    if ($archivedTo) {
        Write-Host "Archived prior release to: $archivedTo" -ForegroundColor DarkGray
    }
    Remove-PmAiExcessPreviousGenerations -ReleaseRoot $ReleaseRoot -KeepGenerations $keepGensResolved
}
else {
    Write-Host "Removing existing release artifacts (SkipArchive): version.txt, matching ZIPs" -ForegroundColor DarkGray
    foreach ($p in @($releaseVersionTxt, $zipInitial, $zipUpgrade)) {
        if (Test-Path -LiteralPath $p) {
            Remove-Item -LiteralPath $p -Force -ErrorAction Stop
        }
    }
    $bm = Join-Path $ReleaseRoot 'build-manifest.json'
    if (Test-Path -LiteralPath $bm) {
        Remove-Item -LiteralPath $bm -Force -ErrorAction SilentlyContinue
    }
}
```

注意: `-SkipArchive` でない場合、退避でルート同名は既に無い。ZIP 生成前に同名が残っていた場合のみ `Compress-PortableBundleFolderToZip` 内の削除が効く。退避後に `version.txt` を新規書き込みする既存処理はそのまま。

- [ ] **Step 4: Update README line in Copy-BundleToDist**

`Release: package start bumps...` の行を、退避について触れる文言に更新（ASCII コメント方針を維持）:

```powershell
$rmLines.Add('Release: package start bumps repo version.txt +0.01 (unless -SkipVersionBump). Step 8 archives prior version.txt / ZIPs / build-manifest.json under previous\v{version}\ (unless -SkipArchive), then writes fresh copies at release root; ZIPs omit pm-ai-data/version.txt. KeepGenerations prunes older previous\ folders. Interim bundle folders are removed after zipping.')
```

- [ ] **Step 5: Re-run helper test + smoke header**

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\test_package_release_archive.ps1
```

Expected: OK

フルパッケージ実行は時間がかかるため任意。可能なら既存 `pm-ai-package-release` にダミー ZIP を置いて Step 8 相当だけ手動で関数呼び出し確認。

- [ ] **Step 6: Commit**

```bash
git add fast_package_app.ps1 scripts/package_release_archive.ps1 scripts/test_package_release_archive.ps1
git commit -m "$(cat <<'EOF'
feat: パッケージリリース成果物を previous 世代管理で退避

EOF
)"
```

---

## Spec coverage（本プラン）

| 要件 | Task |
|------|------|
| previous\v{ver} 退避 | Task 1–2 |
| 退避対象 4 ファイル | Task 1 |
| 衝突時タイムスタンプ | Task 1 |
| KeepGenerations / env / 0=無制限 | Task 1–2 |
| SkipArchive | Task 2 |
| 最新はルート固定名 | Task 2（既存生成を維持） |

タブ／強制適用は次プラン。
