# Shared workspace mirror for portable pm-ai-data (used by fast_package_app.ps1 at repo root and package_app.ps1).
# UTF-8 BOM expected when edited from Windows PowerShell 5.1.

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
    foreach ($n in @('__pycache__', '.pytest_cache', 'build_cache')) {
        $excludedDirNames.Add($n)
    }
    if ($BundleKind -eq 'VersionUpgrade' -or $BundleKind -eq 'DeveloperMirror') {
        $excludedDirNames.Add('plan')
        $excludedDirNames.Add('plans')
    }

    $excludedFileNamePatterns = @(
        '*.log',
        '~$*'
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
            Copy-Item -LiteralPath $full -Destination $dst -Force
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
        Copy-Item -LiteralPath $src -Destination $dst -Force
    }
}
