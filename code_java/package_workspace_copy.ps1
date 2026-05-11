# Shared workspace mirror for portable pm-ai-data (used by fast_package_app.ps1 at repo root and package_app.ps1).
# UTF-8 BOM expected when edited from Windows PowerShell 5.1.

function Get-PmAiRequirementsFingerprint {
    <#
    .SYNOPSIS
      Short stable fingerprint of requirements.txt for Python embed build_cache directory names.
      When code/python/requirements.txt changes, the cache path changes so pip is re-run even with -SkipPythonPrepare.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RequirementsPath
    )
    if (-not (Test-Path -LiteralPath $RequirementsPath)) {
        return 'noreq'
    }
    $bytes = [System.IO.File]::ReadAllBytes($RequirementsPath)
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $hash = $sha.ComputeHash($bytes)
    }
    finally {
        $sha.Dispose()
    }
    $sb = New-Object System.Text.StringBuilder
    $take = [Math]::Min(8, $hash.Length)
    for ($i = 0; $i -lt $take; $i++) {
        [void]$sb.Append($hash[$i].ToString('x2'))
    }
    return $sb.ToString()
}

function Copy-WorkspaceTreeWithExplicitExclusions {
    <#
    .SYNOPSIS
      Copy repo tree into DestRoot with explicit exclusions (filesystem walk, not git).
    .PARAMETER BundleKind
      InitialInstall: excludes IDE/VBA/packaging dirs; does NOT exclude plan/plans or dispatch outputs.
      VersionUpgrade: stricter (also plan/plans, outputs, .pm-ai-cache, env value files except template TSV).
      DeveloperMirror: legacy package_app.ps1 behavior (excludes plan/plans like Upgrade).
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$RepoRoot,

        [Parameter(Mandatory = $true)]
        [string]$DestRoot,

        [Parameter(Mandatory = $true)]
        [ValidateSet('InitialInstall', 'VersionUpgrade', 'DeveloperMirror')]
        [string]$BundleKind,

        [Parameter(Mandatory = $true)]
        [string]$MandatoryPathsFile,

        [Parameter(Mandatory = $true)]
        [string]$ReleaseFolderRelativePrefix
    )

    if (-not $ReleaseFolderRelativePrefix.EndsWith('/')) {
        $ReleaseFolderRelativePrefix = $ReleaseFolderRelativePrefix.TrimEnd('\') + '/'
    }

    # Must be defined before InitialInstall exclusions reference it (avoid $null -> excludes entire tree).
    $referenceDirRel = 'code/' + (-join @([char]0x53C2, [char]0x7167, [char]0x7528)) + '/'

    # Directory prefixes (repo-relative, slash form, must end with '/').
    $excludedDirPrefixes = [System.Collections.Generic.List[string]]::new()
    foreach ($p in @(
            '.git/',
            '.venv/',
            'code_java/target/',
            'code_java/package_input/',
            'code_java/PMD_fast/',
            'code_java/output/',
            'code_java/dist/',
            'code_java/Cash_PMD/',
            '.cursor/',
            '.vscode/',
            'code/VBA/',
            $ReleaseFolderRelativePrefix
        )) {
        $excludedDirPrefixes.Add($p)
    }

    if ($BundleKind -eq 'InitialInstall') {
        foreach ($p in @(
                '.githooks/',
                '.github/',
                '.pm-ai-cache/network-source/',
                # InitialInstall now also excludes developer-machine artifacts (was permissive; bloats release ZIP).
                'output/',
                'code/output/',
                'code/python/output/',
                $referenceDirRel
            )) {
            $excludedDirPrefixes.Add($p)
        }
    }

    if ($BundleKind -eq 'VersionUpgrade') {
        $excludedDirPrefixes.Add('.pm-ai-cache/')
        $excludedDirPrefixes.Add('code/output/')
        $excludedDirPrefixes.Add('output/')
        $excludedDirPrefixes.Add('code/python/output/')
    }

    # Directory base names matched anywhere in the path.
    $excludedDirNames = [System.Collections.Generic.List[string]]::new()
    foreach ($n in @('__pycache__', '.pytest_cache', 'build_cache', '.venv')) {
        $excludedDirNames.Add($n)
    }
    # Note: repo-root prefix '.venv/' also exists in excludedDirPrefixes; '.venv' here excludes nested trees
    # (e.g. manual/.venv/) which must not ship — portable runtime is pm-ai-data/runtime/python-embed only.
    # plan / plans are dev-time dispatch outputs; never ship in release bundles regardless of profile.
    if ($BundleKind -eq 'InitialInstall' -or $BundleKind -eq 'VersionUpgrade' -or $BundleKind -eq 'DeveloperMirror') {
        $excludedDirNames.Add('plan')
        $excludedDirNames.Add('plans')
    }
    # User profile exports under repo (e.g. init_setting/user-profiles/) must not ship in InitialInstall / VersionUpgrade ZIPs.
    if ($BundleKind -eq 'InitialInstall' -or $BundleKind -eq 'VersionUpgrade') {
        $excludedDirNames.Add('user-profiles')
    }

    $excludedFileNamePatterns = @(
        '*.log',
        '~$*',
        # JVM heap dumps (-XX:+HeapDumpOnOutOfMemoryError) and partitioned siblings (.p1 / .p2 / .p3 ...)
        # These can reach multi-GB each on JavaFX / Maven runs and would otherwise inflate the portable ZIP.
        '*.hprof',
        '*.hprof.*',
        # Chromium-style heap snapshots, generic dumps, and Windows minidumps.
        '*.heapsnapshot',
        '*.dump',
        '*.mdmp',
        # Generic tmp files and Windows Explorer metadata. ~$* already handles Office locks.
        '*.tmp',
        'Thumbs.db',
        'desktop.ini'
    )

    # Env template TSV (must stay bundled in all modes). Built without non-ASCII literals in source.
    $templateLeaf = -join @(
        [char]0x8A2D, [char]0x5B9A, [char]0x5F,
        [char]0x74B0, [char]0x5883, [char]0x5909, [char]0x6570,
        [char]0x5F,
        [char]0x96DB, [char]0x5F62,
        '.tsv'
    )
    $envPrefix = -join @(
        [char]0x8A2D, [char]0x5B9A, [char]0x5F,
        [char]0x74B0, [char]0x5883, [char]0x5909, [char]0x6570
    )
    $xlwingsInstallBatRel = 'xlwings' + (-join @(
            [char]0x30A4, [char]0x30F3, [char]0x30B9, [char]0x30C8,
            [char]0x30FC, [char]0x30EB
        )) + '.bat'
    $workspaceLeaf = (-join @(
            [char]0x5DE5, [char]0x7A0B, [char]0x7BA1, [char]0x7406,
            'AI',
            [char]0x30D7, [char]0x30ED, [char]0x30B8, [char]0x30A7,
            [char]0x30AF, [char]0x30C8
        )) + '.code-workspace'

    function Test-IsExcludedDir {
        param([string]$RelSlash)
        if ([string]::IsNullOrEmpty($RelSlash)) { return $false }
        $withSlash = if ($RelSlash.EndsWith('/')) { $RelSlash } else { $RelSlash + '/' }
        foreach ($prefix in $excludedDirPrefixes) {
            if ($withSlash.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
                return $true
            }
        }
        foreach ($seg in $RelSlash.Split('/')) {
            foreach ($name in $excludedDirNames) {
                if ($seg.Equals($name, [StringComparison]::OrdinalIgnoreCase)) {
                    return $true
                }
            }
        }
        return $false
    }

    function Test-IsExcludedFileLeaf {
        param([string]$Leaf)
        foreach ($pat in $excludedFileNamePatterns) {
            if ($Leaf -like $pat) {
                return $true
            }
        }
        return $false
    }

    function Test-IsExcludedExactRepoRelativeFile {
        param([string]$RelSlash)
        $norm = $RelSlash -replace '\\', '/'
        if ($BundleKind -eq 'InitialInstall') {
            foreach ($x in @(
                    $xlwingsInstallBatRel,
                    ('code/' + $workspaceLeaf),
                    'code/----AI------.code-workspace'
                )) {
                if ($norm.Equals($x, [StringComparison]::OrdinalIgnoreCase)) {
                    return $true
                }
            }
        }
        if ($BundleKind -eq 'VersionUpgrade') {
            foreach ($x in @(
                    'config/bundled_session_ui_defaults.json',
                    'config/bundled_table_column_order.json'
                )) {
                if ($norm.Equals($x, [StringComparison]::OrdinalIgnoreCase)) {
                    return $true
                }
            }
        }
        return $false
    }

    function Test-IsExcludedUpgradeEnvFile {
        param([string]$RelSlash, [string]$Leaf)
        if ($BundleKind -ne 'VersionUpgrade') {
            return $false
        }
        if ($Leaf -eq '.env' -or ($Leaf -like '.env.*')) {
            return $true
        }
        if (-not ($Leaf.EndsWith('.tsv', [StringComparison]::OrdinalIgnoreCase))) {
            return $false
        }
        if ($Leaf -eq $templateLeaf) {
            return $false
        }
        if ($Leaf.StartsWith($envPrefix + '_', [StringComparison]::Ordinal)) {
            return $true
        }
        return $false
    }

    $rootFull = (Resolve-Path -LiteralPath $RepoRoot).Path
    $rootLen = $rootFull.Length

    $queue = New-Object System.Collections.Queue
    $queue.Enqueue($rootFull)

    while ($queue.Count -gt 0) {
        $cur = [string]$queue.Dequeue()
        $children = Get-ChildItem -LiteralPath $cur -Force -ErrorAction SilentlyContinue
        foreach ($child in $children) {
            $full = $child.FullName
            if (-not $full.StartsWith($rootFull, [StringComparison]::OrdinalIgnoreCase)) {
                continue
            }
            $rel = $full.Substring($rootLen).TrimStart('\', '/')
            if ([string]::IsNullOrWhiteSpace($rel)) {
                continue
            }
            $relSlash = $rel -replace '\\', '/'

            if ($child.PSIsContainer) {
                if (Test-IsExcludedDir -RelSlash $relSlash) {
                    continue
                }
                $queue.Enqueue($full)
                continue
            }

            if (Test-IsExcludedFileLeaf -Leaf $child.Name) {
                continue
            }
            if (Test-IsExcludedExactRepoRelativeFile -RelSlash $relSlash) {
                continue
            }
            if (Test-IsExcludedUpgradeEnvFile -RelSlash $relSlash -Leaf $child.Name) {
                continue
            }

            $dst = Join-Path $DestRoot $rel
            $parent = Split-Path -Parent $dst
            if (-not [string]::IsNullOrWhiteSpace($parent) -and -not (Test-Path -LiteralPath $parent)) {
                New-Item -ItemType Directory -Path $parent -Force | Out-Null
            }
            try {
                Copy-Item -LiteralPath $full -Destination $dst -Force
            }
            catch [System.IO.IOException] {
                # ASCII only: PS 5.1 loads this file as system ANSI unless UTF-8 BOM (see file header).
                $ex = $_.Exception
                $hr = [System.Runtime.InteropServices.Marshal]::GetHRForException($ex)
                # 0x80070070 ERROR_DISK_FULL (works for localized IOException messages)
                $hint = if ($hr -eq -2147024784 -or $ex.Message -match '(?i)\b(space|full)\b') {
                    'Insufficient disk space likely: remove/move old pm-ai-package-release or free space on destination drive, then retry.'
                }
                else {
                    'Possible file lock or path-too-long; see paths below.'
                }
                throw ("Workspace mirror copy failed. $hint`nOriginal error: $($ex.Message)`nSource: $full`nDestination: $dst")
            }
        }
    }

    if (-not (Test-Path -LiteralPath $MandatoryPathsFile)) {
        throw "Missing mandatory paths list: $MandatoryPathsFile"
    }
    $mandatoryCodeRootTxt = @(
        Get-Content -LiteralPath $MandatoryPathsFile -Encoding UTF8 |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -ne '' -and ($_ -notmatch '^\s*#') }
    )
    foreach ($relSlash in $mandatoryCodeRootTxt) {
        $relOs = $relSlash -replace '/', [System.IO.Path]::DirectorySeparatorChar
        $src = Join-Path $RepoRoot $relOs
        if (-not (Test-Path -LiteralPath $src)) {
            continue
        }
        $dst = Join-Path $DestRoot $relOs
        $parent = Split-Path -Parent $dst
        if (-not [string]::IsNullOrWhiteSpace($parent) -and -not (Test-Path -LiteralPath $parent)) {
            New-Item -ItemType Directory -Path $parent -Force | Out-Null
        }
        try {
            Copy-Item -LiteralPath $src -Destination $dst -Force
        }
        catch [System.IO.IOException] {
            $ex = $_.Exception
            $hr = [System.Runtime.InteropServices.Marshal]::GetHRForException($ex)
            $hint = if ($hr -eq -2147024784 -or $ex.Message -match '(?i)\b(space|full)\b') {
                'Insufficient disk space likely: remove/move old pm-ai-package-release or free space on destination drive, then retry.'
            }
            else {
                'Possible file lock or path-too-long; see paths below.'
            }
            throw ("Mandatory-path copy failed. $hint`nOriginal error: $($ex.Message)`nSource: $src`nDestination: $dst")
        }
    }
}
