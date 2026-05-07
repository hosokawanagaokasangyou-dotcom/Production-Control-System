# Production desktop (JavaFX) - Windows app bundle builder (jpackage + bundled runtime).
#
# Prerequisites:
#   - Run build on Windows (OpenJFX win classifier required).
#   - Maven uses JAVA_HOME on PATH (compile must match maven.compiler.release).
#   - Bundled JVM: Temurin JDK zip -> build_cache -> jpackage --runtime-image. Override: -JdkRuntimeImage or PM_AI_JDK_RUNTIME_IMAGE.
#   - JavaFX: OpenJFX Windows win jars downloaded from Maven Central into package_input (same version as pom javafx.version).
#   - For --type exe/msi: WiX Toolset on PATH (candle/light).
#   - Bundled Python: pip runs at build time - internet on first run or empty cache.
#
# Usage:
#   .\package_app.ps1
#   .\package_app.ps1 -PackageType exe
#   .\package_app.ps1 -SkipPythonPrepare   # reuse existing build_cache python (faster)
#   .\package_app.ps1 -SkipJdkPrepare      # reuse existing build_cache JDK extract (faster)
#   .\package_app.ps1 -SkipJavaFxPrepare   # reuse cached OpenJFX win jars (faster)
#   .\package_app.ps1 -WinConsole
#   .\package_app.ps1 -JpackageDest C:\pm-ai-out   # ASCII-only parent for jpackage --dest (if launchers missing)
#   .\package_app.ps1 -JdkRuntimeImage C:\path\to\jdk   # skip download; needs bin\java.exe and bin\jpackage.exe
# Env: PM_AI_JPACKAGE_DEST, PM_AI_JDK_RUNTIME_IMAGE (optional)

# UTF-8 BOM: Windows PowerShell 5.1 parses this file as UTF-8. Body is ASCII-only; Japanese paths live in package_app_mandatory_code_paths.txt.
[CmdletBinding()]
param(
    [ValidateSet('app-image', 'exe', 'msi')]
    [string]$PackageType = 'app-image',

    [switch]$WinConsole,

    [switch]$SkipPythonPrepare,

    [switch]$SkipJdkPrepare,

    [switch]$SkipJavaFxPrepare,

    # JDK root for --runtime-image (bin\java.exe). Empty = download per pom.xml into build_cache.
    [string]$JdkRuntimeImage = '',

    # Parent directory for jpackage --dest only (must be ASCII-only on some JDK/Windows builds).
    [string]$JpackageDest = ''
)

$ErrorActionPreference = 'Stop'

$Root = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
Set-Location $Root

$WorkspaceRoot = (Resolve-Path -LiteralPath (Join-Path $Root '..')).Path

function Copy-WorkspaceTreeWithExplicitExclusions {
    # Walk the workspace by filesystem (no git, no .gitignore). Exclusions are individually listed
    # below so adding/removing one path is a one-liner edit. Excluded dirs are pruned during
    # traversal so we never recurse into .git/ or huge build_cache/ trees.
    param(
        [string]$RepoRoot,
        [string]$DestRoot
    )

    # Directory prefixes (repo-relative, slash form, must end with '/'). Subtree pruned at this point.
    $excludedDirPrefixes = @(
        '.git/',
        '.venv/',
        'code_java/target/',
        'code_java/build_cache/',
        'code_java/package_input/',
        'code_java/dist/',
        'code_java/output/'
    )

    # Directory base names matched anywhere in the path (Python / pytest caches).
    $excludedDirNames = @(
        '__pycache__',
        '.pytest_cache'
    )

    # File-name globs (matched against the leaf name only). Excel lock + log noise.
    $excludedFileNamePatterns = @(
        '*.log',
        '~$*'
    )

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

    # Ancestors are already pruned during BFS, so file check is leaf-name only.
    function Test-IsExcludedFile {
        param([string]$Leaf)
        foreach ($pat in $excludedFileNamePatterns) {
            if ($Leaf -like $pat) {
                return $true
            }
        }
        return $false
    }

    $rootFull = (Resolve-Path -LiteralPath $RepoRoot).Path
    $rootLen = $rootFull.Length

    # Iterative BFS so we can prune excluded directories without descending into them.
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

            if (Test-IsExcludedFile -Leaf $child.Name) {
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

    # Always copy master *.txt under code/ (paths listed in UTF-8 sidecar file).
    $pathsFile = Join-Path $Root 'package_app_mandatory_code_paths.txt'
    if (-not (Test-Path -LiteralPath $pathsFile)) {
        throw "Missing mandatory paths list: $pathsFile"
    }
    $mandatoryCodeRootTxt = @(
        Get-Content -LiteralPath $pathsFile -Encoding UTF8 |
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

function Read-MavenPomProperties {
    param([string]$PomPath)
    [xml]$xml = Get-Content -LiteralPath $PomPath -Encoding UTF8
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('m', 'http://maven.apache.org/POM/4.0.0')
    $props = @{}
    foreach ($n in $xml.SelectNodes('/m:project/m:properties/*', $ns)) {
        $props[$n.LocalName] = $n.InnerText.Trim()
    }
    return $props
}

function Expand-PomPropertyPlaceholder {
    param(
        [string]$Raw,
        [hashtable]$Props
    )
    if ($null -eq $Raw -or [string]::IsNullOrWhiteSpace($Raw)) {
        return ''
    }
    $current = $Raw.Trim()
    for ($iter = 0; $iter -lt 4; $iter++) {
        $m = [regex]::Match($current, '^\$\{([^}]+)\}$')
        if (-not $m.Success) {
            break
        }
        $innerKey = $m.Groups[1].Value
        if (-not $Props.ContainsKey($innerKey)) {
            break
        }
        $next = [string]$Props[$innerKey]
        if ([string]::IsNullOrWhiteSpace($next)) {
            break
        }
        $current = $next.Trim()
    }
    return $current
}

function Get-MavenProjectInfo {
    param([string]$PomPath)
    [xml]$xml = Get-Content -LiteralPath $PomPath -Encoding UTF8
    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('m', 'http://maven.apache.org/POM/4.0.0')
    $artifact = $xml.SelectSingleNode('/m:project/m:artifactId', $ns).InnerText.Trim()
    $versionNode = $xml.SelectSingleNode('/m:project/m:version', $ns)
    if (-not $versionNode -or [string]::IsNullOrWhiteSpace($versionNode.InnerText)) {
        $parentVer = $xml.SelectSingleNode('/m:project/m:parent/m:version', $ns)
        if ($parentVer) {
            $version = $parentVer.InnerText.Trim()
        }
        else {
            throw 'Could not read version from pom.xml.'
        }
    }
    else {
        $version = $versionNode.InnerText.Trim()
    }
    if (-not $artifact -or -not $version) {
        throw 'Could not read artifactId / version from pom.xml.'
    }
    $mainJar = "$artifact-$version.jar"
    @{
        ArtifactId = $artifact
        Version    = $version
        MainJar    = $mainJar
    }
}

function Copy-JpackageInputDirectory {
    param(
        [string]$RootPath,
        [string]$MainJarName,
        [string]$DestPath
    )
    if (Test-Path -LiteralPath $DestPath) {
        Remove-Item -Recurse -Force $DestPath
    }
    New-Item -ItemType Directory -Path $DestPath | Out-Null

    $mainSrc = Join-Path (Join-Path $RootPath 'target') $MainJarName
    if (-not (Test-Path -LiteralPath $mainSrc)) {
        throw "Main JAR not found: $mainSrc"
    }
    Copy-Item -LiteralPath $mainSrc -Destination $DestPath

    $depDir = Join-Path (Join-Path $RootPath 'target') 'dependency'
    if (-not (Test-Path -LiteralPath $depDir)) {
        throw "dependency folder not found: $depDir"
    }
    Copy-Item -Path (Join-Path $depDir '*') -Destination $DestPath -Force
}

function Ensure-JdkWindowsEmbedCache {
    param(
        [string]$CacheRoot,
        [string]$JdkRelease,
        [string]$ZipUrlOverride,
        [bool]$Skip
    )

    $dest = Join-Path $CacheRoot ('jdk-embed-' + $JdkRelease + '-windows-amd64')
    $javaExe = Join-Path $dest 'bin\java.exe'
    $jpkgExe = Join-Path $dest 'bin\jpackage.exe'

    if ($Skip -and (Test-Path -LiteralPath $javaExe) -and (Test-Path -LiteralPath $jpkgExe)) {
        Write-Host "SkipJdkPrepare: using cache: $dest" -ForegroundColor DarkGray
        return [string]$dest
    }

    if (Test-Path -LiteralPath $dest) {
        Remove-Item -Recurse -Force -LiteralPath $dest
    }
    New-Item -ItemType Directory -Path $dest -Force | Out-Null

    $zipPath = Join-Path $dest 'jdk-bundle.zip'
    if (-not [string]::IsNullOrWhiteSpace($ZipUrlOverride)) {
        $url = $ZipUrlOverride.Trim()
        Write-Host "--- Download JDK zip (pom pm.ai.bundle.jdk.windows.zip.url): $url ---" -ForegroundColor Cyan
    }
    else {
        $url = "https://api.adoptium.net/v3/binary/latest/$JdkRelease/ga/windows/x64/jdk/hotspot/normal/eclipse"
        Write-Host "--- Download JDK zip (Adoptium API, Windows x64 release $JdkRelease): $url ---" -ForegroundColor Cyan
    }

    Invoke-WebRequest -Uri $url -OutFile $zipPath -UseBasicParsing

    $extractTmp = Join-Path $dest '_ext'
    New-Item -ItemType Directory -Path $extractTmp -Force | Out-Null
    try {
        Expand-Archive -LiteralPath $zipPath -DestinationPath $extractTmp -Force
    }
    finally {
        Remove-Item -LiteralPath $zipPath -Force -ErrorAction SilentlyContinue
    }

    $javaFound = Get-ChildItem -Path $extractTmp -Recurse -Filter 'java.exe' -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Directory.Name -ieq 'bin' } |
        Select-Object -First 1
    if (-not $javaFound) {
        throw "JDK zip did not contain bin\java.exe under: $extractTmp"
    }

    $jdkHome = $javaFound.Directory.Parent.FullName
    Get-ChildItem -LiteralPath $jdkHome -ErrorAction SilentlyContinue | ForEach-Object {
        Move-Item -LiteralPath $_.FullName -Destination $dest -Force
    }
    Remove-Item -LiteralPath $extractTmp -Recurse -Force -ErrorAction SilentlyContinue

    if (-not (Test-Path -LiteralPath $javaExe)) {
        throw "JDK layout error: missing $javaExe"
    }
    if (-not (Test-Path -LiteralPath $jpkgExe)) {
        throw "JDK layout error: missing $jpkgExe"
    }

    return [string]$dest
}

function Ensure-PythonEmbedCache {
    param(
        [string]$WorkspaceRootPath,
        [string]$PythonVersion,
        [string]$CacheRoot,
        [bool]$Skip
    )

    $dest = Join-Path $CacheRoot "python-embed-$PythonVersion-amd64"
    $pyExe = Join-Path $dest 'python.exe'
    $req = Join-Path $WorkspaceRootPath 'code\python\requirements.txt'

    if ($Skip -and (Test-Path -LiteralPath $pyExe)) {
        Write-Host "SkipPythonPrepare: using cache: $dest" -ForegroundColor DarkGray
        return [string]$dest
    }

    if (-not (Test-Path -LiteralPath $req)) {
        throw "requirements.txt not found: $req"
    }

    New-Item -ItemType Directory -Path $CacheRoot -Force | Out-Null
    if (Test-Path -LiteralPath $dest) {
        Remove-Item -Recurse -Force $dest
    }
    New-Item -ItemType Directory -Path $dest | Out-Null

    $zipUrl = "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-embed-amd64.zip"
    $zipPath = Join-Path $dest 'python-embed.zip'
    Write-Host "--- Download Python embed: $zipUrl ---" -ForegroundColor Cyan
    Invoke-WebRequest -Uri $zipUrl -OutFile $zipPath -UseBasicParsing
    Expand-Archive -LiteralPath $zipPath -DestinationPath $dest -Force
    Remove-Item -LiteralPath $zipPath -Force

    Get-ChildItem -LiteralPath $dest -Filter '*._pth' | ForEach-Object {
        $t = Get-Content -LiteralPath $_.FullName -Raw
        if ($t -notmatch '(?m)^import site\s*$') {
            Add-Content -LiteralPath $_.FullName -Value "`r`nimport site`r`n" -Encoding UTF8
        }
    }

    $getPip = Join-Path $dest 'get-pip.py'
    Write-Host "--- Download get-pip.py ---" -ForegroundColor Cyan
    Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile $getPip -UseBasicParsing

    Push-Location $dest
    try {
        # PS 5.1 + ErrorAction Stop: python stderr WARNING may trigger NativeCommandError.
        $prevEa = $ErrorActionPreference
        try {
            $ErrorActionPreference = 'SilentlyContinue'
            $env:PIP_NO_WARN_SCRIPT_LOCATION = '1'
            & .\python.exe $getPip *> $null
            if ($LASTEXITCODE -ne 0) {
                throw 'get-pip failed.'
            }
            & .\python.exe -m pip install -q --upgrade pip --no-warn-script-location *> $null
            if ($LASTEXITCODE -ne 0) {
                throw 'pip upgrade failed.'
            }
            & .\python.exe -m pip install -q -r $req --no-warn-script-location *> $null
            if ($LASTEXITCODE -ne 0) {
                throw 'pip install -r requirements.txt failed.'
            }
        }
        finally {
            $ErrorActionPreference = $prevEa
            Remove-Item Env:PIP_NO_WARN_SCRIPT_LOCATION -ErrorAction SilentlyContinue
        }
    }
    finally {
        Pop-Location
    }

    return [string]$dest
}

function Normalize-JvmHeapToken {
    param([string]$Raw)
    $t = ($Raw -replace '[\r\n\t]', '').Trim()
    if ([string]::IsNullOrWhiteSpace($t)) {
        return '512m'
    }
    return $t
}

function Build-PmAiDesktopLauncherBatContent {
    param(
        [string]$JavafxVersion,
        [string]$JvmInitial,
        [string]$JvmMax,
        [string]$LauncherExeBaseName = 'PMD'
    )
    $jv = ($JavafxVersion -replace '[\r\n\t]', '').Trim()
    if ([string]::IsNullOrWhiteSpace($jv)) {
        $jv = '26.0.1'
    }
    $xms = Normalize-JvmHeapToken $JvmInitial
    $xmx = Normalize-JvmHeapToken $JvmMax

    $artifacts = @('javafx-base', 'javafx-controls', 'javafx-fxml', 'javafx-graphics', 'javafx-swing')
    $lines = [System.Collections.Generic.List[string]]::new()
    $lines.Add('@echo off')
    $lines.Add('rem ASCII-only. Generated by package_app.ps1 from pom javafx.version / jvm heap.')
    $lines.Add('rem Do not paste into PowerShell; run: .\launch-pm-ai-desktop.bat')
    $lines.Add('setlocal EnableExtensions EnableDelayedExpansion')
    $lines.Add('')
    $lines.Add('set "ROOT=%~dp0"')
    $lines.Add('if "%ROOT:~-1%"=="\" set "ROOT=%ROOT:~0,-1%"')
    $lines.Add('cd /d "%ROOT%"')
    $lines.Add('')
    $lines.Add('if not exist "%ROOT%\app" (')
    $lines.Add('    echo [ERROR] Missing app folder. Put this bat next to ' + $LauncherExeBaseName + '.exe / app / runtime.')
    $lines.Add('    echo Current: "%ROOT%"')
    $lines.Add('    pause')
    $lines.Add('    exit /b 1')
    $lines.Add(')')
    $lines.Add('')
    $lines.Add('set "JAVA_EXE=%ROOT%\runtime\bin\java.exe"')
    $lines.Add('if exist "%JAVA_EXE%" goto :have_java')
    $lines.Add('')
    $lines.Add('if defined JAVA_HOME (')
    $lines.Add('    if exist "%JAVA_HOME%\bin\java.exe" (')
    $lines.Add('        set "JAVA_EXE=%JAVA_HOME%\bin\java.exe"')
    $lines.Add('        echo [WARN] Using JAVA_HOME java.exe (bundled runtime missing).')
    $lines.Add('        goto :have_java')
    $lines.Add('    )')
    $lines.Add(')')
    $lines.Add('')
    $lines.Add('echo [ERROR] Java not found: "%ROOT%\runtime\bin\java.exe"')
    $lines.Add('pause')
    $lines.Add('exit /b 1')
    $lines.Add('')
    $lines.Add(':have_java')
    $lines.Add('')
    $lines.Add('set "PM_AI_JFX_MODPATH=%ROOT%\app\' + $artifacts[0] + '-' + $jv + '-win.jar"')
    for ($ai = 1; $ai -lt $artifacts.Count; $ai++) {
        $lines.Add('set "PM_AI_JFX_MODPATH=!PM_AI_JFX_MODPATH!;%ROOT%\app\' + $artifacts[$ai] + '-' + $jv + '-win.jar"')
    }
    $lines.Add('')
    # Must match jpackage --java-options ($javaOpts): JavaFX --module-path/--add-modules + ControlsFX (internal JavaFX) opens/exports.
    $compatJvm = '--add-opens=javafx.base/com.sun.javafx.event=ALL-UNNAMED --add-opens=javafx.controls/javafx.scene.control.skin=ALL-UNNAMED --add-exports=javafx.controls/com.sun.javafx.scene.control.behavior=ALL-UNNAMED --enable-native-access=javafx.graphics'
    $javaLine = '"%JAVA_EXE%" -Dfile.encoding=UTF-8 -Xms' + $xms + ' -Xmx' + $xmx + ' -XX:+HeapDumpOnOutOfMemoryError -XX:+UseStringDeduplication -Dprism.order=sw ' + $compatJvm + ' --module-path "!PM_AI_JFX_MODPATH!" --add-modules javafx.controls,javafx.fxml,javafx.graphics,javafx.base,javafx.swing -classpath "%ROOT%\app\*" jp.co.pm.ai.desktop.PmAiFxApp %*'
    $lines.Add($javaLine)
    $lines.Add('')
    $lines.Add('set EXITCODE=!ERRORLEVEL!')
    $lines.Add('')
    $lines.Add('if !EXITCODE! neq 0 (')
    $lines.Add('    echo.')
    $lines.Add('    echo [Exit !EXITCODE!] Logs: !USERPROFILE!\.pm-ai-desktop\startup.log  or  !TEMP!\pm-ai-desktop-startup.log')
    $lines.Add(')')
    $lines.Add('')
    $lines.Add('exit /b !EXITCODE!')
    $lines.Add('')
    return ($lines.ToArray() -join "`r`n")
}

function Sync-JavaFxWindowsRuntimeFromMavenCentral {
    param(
        [string]$PackageInputDir,
        [string]$JavafxVersion,
        [string]$CacheRoot,
        [bool]$Skip
    )

    $artifacts = @(
        'javafx-base',
        'javafx-controls',
        'javafx-fxml',
        'javafx-graphics',
        'javafx-swing'
    )
    $cacheDir = Join-Path $CacheRoot "javafx-openjfx-$JavafxVersion-windows-amd64"
    New-Item -ItemType Directory -Path $PackageInputDir -Force | Out-Null
    New-Item -ItemType Directory -Path $cacheDir -Force | Out-Null

    foreach ($aid in $artifacts) {
        $fn = "$aid-$JavafxVersion-win.jar"
        $cached = Join-Path $cacheDir $fn
        $url = "https://repo1.maven.org/maven2/org/openjfx/$aid/$JavafxVersion/$fn"

        $needDownload = $true
        if ($Skip -and (Test-Path -LiteralPath $cached)) {
            $fi = Get-Item -LiteralPath $cached -ErrorAction SilentlyContinue
            if ($null -ne $fi -and $fi.Length -gt 4096) {
                $needDownload = $false
                Write-Host "SkipJavaFxPrepare: using cache $fn" -ForegroundColor DarkGray
            }
        }

        if ($needDownload) {
            Write-Host "--- Download JavaFX runtime: $fn ---" -ForegroundColor Cyan
            try {
                Invoke-WebRequest -Uri $url -OutFile $cached -UseBasicParsing
            }
            catch {
                throw "JavaFX download failed: $url $($_.Exception.Message)"
            }
            $fi2 = Get-Item -LiteralPath $cached -ErrorAction SilentlyContinue
            if ($null -eq $fi2 -or $fi2.Length -lt 4096) {
                throw "JavaFX download invalid or empty: $url"
            }
        }

        Copy-Item -LiteralPath $cached -Destination (Join-Path $PackageInputDir $fn) -Force
    }
}

function Copy-BundleToDist {
    param(
        [string]$WorkspaceRootPath,
        [string]$DistAppRoot,
        [string]$PythonEmbedSourceDir,
        [string]$AppExeBaseName = 'PMD'
    )

    if ([string]::IsNullOrWhiteSpace($PythonEmbedSourceDir) -or -not (Test-Path -LiteralPath $PythonEmbedSourceDir)) {
        throw "Invalid Python embed path: '$PythonEmbedSourceDir'"
    }

    $data = Join-Path $DistAppRoot 'pm-ai-data'
    if (Test-Path -LiteralPath $data) {
        Remove-Item -Recurse -Force $data
    }

    New-Item -ItemType Directory -Path $data -Force | Out-Null

    Write-Host "--- Copy workspace into pm-ai-data (explicit exclude list + mandatory code/*.txt) ---" -ForegroundColor Cyan
    Copy-WorkspaceTreeWithExplicitExclusions -RepoRoot $WorkspaceRootPath -DestRoot $data

    New-Item -ItemType Directory -Path (Join-Path $data 'input\task-input') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $data 'input\actual-detail') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $data 'output') -Force | Out-Null

    $rt = Join-Path $data 'runtime\python-embed'
    New-Item -ItemType Directory -Path $rt -Force | Out-Null
    Write-Host "--- Copy Python runtime into pm-ai-data ---" -ForegroundColor Cyan
    & robocopy $PythonEmbedSourceDir $rt /E /NFL /NDL /NJH /NJS /nc /ns /np | Out-Host
    $rc2 = $LASTEXITCODE
    if ($rc2 -ge 8) {
        throw "robocopy python-embed failed (exit $rc2)"
    }

    $readme = Join-Path $data 'README_PORTABLE.txt'
    @(
        'JVM: bundled via jpackage --runtime-image from Temurin JDK (see pom.xml / package_app.ps1).',
        'JavaFX: OpenJFX Windows jars pinned from Maven Central into package_input during package_app.ps1 (javafx.version).',
        'Portable bundle generated by package_app.ps1.',
        'Workspace mirror source: filesystem walk with explicit exclude list in package_app.ps1.',
        'Master *.txt under code/ are always copied (see package_app_mandatory_code_paths.txt).',
        'Excluded dirs: .git, .venv, code_java/{target,build_cache,package_input,dist,output}, **/__pycache__, **/.pytest_cache.',
        'Excluded files: *.log, ~$* (Excel lock).',
        "This folder sits next to $($AppExeBaseName).exe.",
        'Version: repo-root version.txt (included here). Optional sync via PM_AI_PORTABLE_BUNDLE_SOURCE_DIR.',
        'Python: pm-ai-data\runtime\python-embed\python.exe (from build_cache embed + pip).',
        'Default inputs: input\task-input , input\actual-detail.',
        'Per-user session data: ~/.pm-ai-desktop (initialized per machine/user).',
        ''
    ) | Set-Content -LiteralPath $readme -Encoding UTF8
}

$POM = Join-Path $Root 'pom.xml'
$pomProps = Read-MavenPomProperties -PomPath $POM
$jvmInitial = $pomProps['jvm.initial.heap']
$jvmMax = $pomProps['jvm.max.heap']
$prismOrder = $pomProps['javafx.prism.order']
$pyEmbedVer = $pomProps['pm.ai.bundle.python.embed.version']
if ([string]::IsNullOrWhiteSpace($jvmInitial)) { $jvmInitial = '512m' }
if ([string]::IsNullOrWhiteSpace($jvmMax)) { $jvmMax = '3g' }
if ([string]::IsNullOrWhiteSpace($prismOrder)) { $prismOrder = 'sw' }
if ([string]::IsNullOrWhiteSpace($pyEmbedVer)) {
    throw 'pom.properties missing pm.ai.bundle.python.embed.version.'
}

$proj = Get-MavenProjectInfo -PomPath $POM

$APP_NAME = 'PMD'
$VersionTxtPath = Join-Path $WorkspaceRoot 'version.txt'
$APP_VERSION = $proj.Version -replace '-SNAPSHOT$', '.0'
if (Test-Path -LiteralPath $VersionTxtPath) {
    $rawTxt = (Get-Content -LiteralPath $VersionTxtPath -Raw -Encoding UTF8).Trim()
    $firstLine = ($rawTxt -split "`r?`n")[0].Trim()
    if (-not [string]::IsNullOrWhiteSpace($firstLine)) {
        $sbDigits = [System.Text.StringBuilder]::new()
        foreach ($ch in $firstLine.ToCharArray()) {
            if ([char]::IsDigit($ch) -or ($ch -eq '.')) {
                [void]$sbDigits.Append($ch)
            }
        }
        $digitsdots = $sbDigits.ToString()
        if (-not [string]::IsNullOrWhiteSpace($digitsdots)) {
            while ($digitsdots.Contains('..')) {
                $digitsdots = $digitsdots.Replace('..', '.')
            }
            if ($digitsdots.StartsWith('.')) {
                $digitsdots = '0' + $digitsdots
            }
            if ($digitsdots.EndsWith('.')) {
                $digitsdots = $digitsdots + '0'
            }
            $APP_VERSION = $digitsdots
            if ($APP_VERSION -match '^\d+\.\d+$') {
                $APP_VERSION = "$APP_VERSION.0"
            }
        }
    }
}

Write-Host "--- Step 1: Maven package ---" -ForegroundColor Cyan
$mvnw = Join-Path $Root 'mvnw.cmd'
if (-not (Test-Path -LiteralPath $mvnw)) {
    throw "Maven Wrapper not found: $mvnw"
}
& $mvnw @('clean', 'package', '-DskipTests')
if ($LASTEXITCODE -ne 0) {
    Write-Error 'Maven build failed.'
    exit $LASTEXITCODE
}

Write-Host "--- Step 2: jpackage input directory ---" -ForegroundColor Cyan
$packageInput = Join-Path $Root 'package_input'
Copy-JpackageInputDirectory -RootPath $Root -MainJarName $proj.MainJar -DestPath $packageInput

$cacheRoot = Join-Path $Root 'build_cache'

Write-Host "--- Step 3: JavaFX Windows runtime (Maven Central win jars -> package_input) ---" -ForegroundColor Cyan
$javafxVer = Expand-PomPropertyPlaceholder -Raw ([string]$pomProps['javafx.version']) -Props $pomProps
if ([string]::IsNullOrWhiteSpace($javafxVer)) {
    throw 'pom.xml: javafx.version is required for JavaFX bundle.'
}
Sync-JavaFxWindowsRuntimeFromMavenCentral -PackageInputDir $packageInput -JavafxVersion $javafxVer -CacheRoot $cacheRoot -Skip:$SkipJavaFxPrepare

Write-Host "--- Step 4: Windows JDK bundle (Temurin zip -> jpackage --runtime-image) ---" -ForegroundColor Cyan
$jdkRelease = Expand-PomPropertyPlaceholder -Raw ([string]$pomProps['pm.ai.bundle.jdk.windows.release']) -Props $pomProps
if ([string]::IsNullOrWhiteSpace($jdkRelease)) {
    $jdkRelease = Expand-PomPropertyPlaceholder -Raw ([string]$pomProps['maven.compiler.release']) -Props $pomProps
}
if ([string]::IsNullOrWhiteSpace($jdkRelease)) {
    throw 'pom.xml: set maven.compiler.release or pm.ai.bundle.jdk.windows.release.'
}
$jdkZipUrlOverride = ''
if ($pomProps.ContainsKey('pm.ai.bundle.jdk.windows.zip.url')) {
    $jdkZipUrlOverride = [string]$pomProps['pm.ai.bundle.jdk.windows.zip.url']
}

if (-not [string]::IsNullOrWhiteSpace($JdkRuntimeImage)) {
    $jdkRoot = $JdkRuntimeImage.TrimEnd('\', '/')
}
elseif (-not [string]::IsNullOrWhiteSpace($env:PM_AI_JDK_RUNTIME_IMAGE)) {
    $jdkRoot = $env:PM_AI_JDK_RUNTIME_IMAGE.Trim().TrimEnd('\', '/')
}
else {
    $jdkRoot = [string](Ensure-JdkWindowsEmbedCache -CacheRoot $cacheRoot -JdkRelease $jdkRelease `
            -ZipUrlOverride $jdkZipUrlOverride -Skip:$SkipJdkPrepare)
}

$jdkJavaExe = Join-Path $jdkRoot 'bin\java.exe'
$JPACKAGE = Join-Path $jdkRoot 'bin\jpackage.exe'
if (-not (Test-Path -LiteralPath $jdkJavaExe)) {
    throw "JDK folder missing bin\java.exe (runtime-image): $jdkRoot"
}
if (-not (Test-Path -LiteralPath $JPACKAGE)) {
    throw "JDK folder missing bin\jpackage.exe: $jdkRoot"
}
Write-Host "Using JDK for jpackage + bundled runtime: $jdkRoot" -ForegroundColor DarkGray

Write-Host "--- Step 5: Python embed cache (pip) ---" -ForegroundColor Cyan
$pythonSrc = [string](Ensure-PythonEmbedCache -WorkspaceRootPath $WorkspaceRoot -PythonVersion $pyEmbedVer `
        -CacheRoot $cacheRoot -Skip:$SkipPythonPrepare)

Write-Host "--- Step 6: jpackage (type=$PackageType) ---" -ForegroundColor Cyan

# Final output always under code_java\dist. jpackage --dest may use a staging folder: non-ASCII paths
# can produce runtime\bin with DLLs but no java.exe on some JDK builds.
$distFinal = Join-Path $Root 'dist'
$jpkgDestParent = $distFinal
if (-not [string]::IsNullOrWhiteSpace($JpackageDest)) {
    $jpkgDestParent = $JpackageDest.TrimEnd('\', '/')
}
elseif (-not [string]::IsNullOrWhiteSpace($env:PM_AI_JPACKAGE_DEST)) {
    $jpkgDestParent = $env:PM_AI_JPACKAGE_DEST.Trim().TrimEnd('\', '/')
}
elseif ($Root -match '[^\x00-\x7F]') {
    $jpkgDestParent = Join-Path $env:TEMP ("pm-ai-jpackage-" + [Guid]::NewGuid().ToString('N'))
    Write-Host "Repo path contains non-ASCII: staging jpackage --dest to ASCII-only: $jpkgDestParent" -ForegroundColor Cyan
}
$usedStagingForJpackage = ($jpkgDestParent -ne $distFinal)

# When staging to TEMP, do not pre-delete code_java\dist: Explorer / running exe often locks it.
# jpackage only writes to $jpkgDestParent; we replace dist\<APP_NAME> after success.
$pathsToClean = @($jpkgDestParent)
if (-not $usedStagingForJpackage) {
    $pathsToClean = @($distFinal)
}
$pathsToClean = $pathsToClean | Select-Object -Unique
foreach ($p in $pathsToClean) {
    if ([string]::IsNullOrWhiteSpace($p)) {
        continue
    }
    if (-not (Test-Path -LiteralPath $p)) {
        continue
    }
    $removed = $false
    for ($i = 0; $i -lt 5; $i++) {
        try {
            Remove-Item -Recurse -Force -LiteralPath $p -ErrorAction Stop
            $removed = $true
            break
        }
        catch {
            Write-Warning "Cannot remove $p (locked?). Retry ($i/5)..."
            Start-Sleep -Seconds 2
        }
    }
    if (-not $removed -and (Test-Path -LiteralPath $p)) {
        throw "Cannot remove folder (close Explorer/app using it): $p"
    }
}
if ($usedStagingForJpackage) {
    New-Item -ItemType Directory -Path $jpkgDestParent -Force | Out-Null
}

# Native exe launch uses only jpackage --java-options (see dist\<APP_NAME>\app\<APP_NAME>.cfg).
# Match launch-pm-ai-desktop.bat: JavaFX modular jars on --module-path + --add-modules.
# jpackage cfg understands $APPDIR (bundle root); jars land under app\ next to the launcher.
# Oracle: custom --module-path in --java-options is appended to any default module path.
$jvForJpkgOpts = ($javafxVer -replace '[\r\n\t]', '').Trim()
$jfxModsForJpkg = @('javafx-base', 'javafx-controls', 'javafx-fxml', 'javafx-graphics', 'javafx-swing')
$modPathJpkgSb = [System.Text.StringBuilder]::new()
for ($mi = 0; $mi -lt $jfxModsForJpkg.Count; $mi++) {
    if ($mi -gt 0) {
        [void]$modPathJpkgSb.Append(';')
    }
    [void]$modPathJpkgSb.Append('$APPDIR/app/')
    [void]$modPathJpkgSb.Append($jfxModsForJpkg[$mi])
    [void]$modPathJpkgSb.Append('-')
    [void]$modPathJpkgSb.Append($jvForJpkgOpts)
    [void]$modPathJpkgSb.Append('-win.jar')
}
$jpackageModulePathJavaOpt = '--module-path=' + $modPathJpkgSb.ToString()

$javaOpts = @(
    '-Dfile.encoding=UTF-8',
    "-Xms$jvmInitial",
    "-Xmx$jvmMax",
    '-XX:+HeapDumpOnOutOfMemoryError',
    '-XX:+UseStringDeduplication',
    "-Dprism.order=$prismOrder",
    $jpackageModulePathJavaOpt,
    '--add-modules=javafx.controls,javafx.fxml,javafx.graphics,javafx.base,javafx.swing',
    '--add-opens=javafx.base/com.sun.javafx.event=ALL-UNNAMED',
    '--add-opens=javafx.controls/javafx.scene.control.skin=ALL-UNNAMED',
    '--add-exports=javafx.controls/com.sun.javafx.scene.control.behavior=ALL-UNNAMED',
    '--enable-native-access=javafx.graphics'
)

$jpkgArgs = [System.Collections.Generic.List[string]]::new()
$jpkgArgs.Add('--type')
$jpkgArgs.Add($PackageType)
$jpkgArgs.Add('--input')
$jpkgArgs.Add($packageInput)
$jpkgArgs.Add('--runtime-image')
$jpkgArgs.Add($jdkRoot)
$jpkgArgs.Add('--dest')
$jpkgArgs.Add($jpkgDestParent)
$jpkgArgs.Add('--name')
$jpkgArgs.Add($APP_NAME)
$jpkgArgs.Add('--main-jar')
$jpkgArgs.Add($proj.MainJar)
$jpkgArgs.Add('--main-class')
$jpkgArgs.Add('jp.co.pm.ai.desktop.PmAiFxApp')
$jpkgArgs.Add('--app-version')
$jpkgArgs.Add($APP_VERSION)
$jpkgArgs.Add('--vendor')
$jpkgArgs.Add('jp.co.pm.ai')
$jpkgArgs.Add('--copyright')
$jpkgArgs.Add('Copyright (C) 2026')
$jpkgArgs.Add('--description')
$jpkgArgs.Add('Production Control Desktop (JavaFX)')

foreach ($opt in $javaOpts) {
    $jpkgArgs.Add('--java-options')
    $jpkgArgs.Add($opt)
}

if ($WinConsole) {
    $jpkgArgs.Add('--win-console')
}

if ($PackageType -eq 'exe' -or $PackageType -eq 'msi') {
    $jpkgArgs.Add('--win-shortcut')
    $jpkgArgs.Add('--win-menu')
    $jpkgArgs.Add('--win-dir-chooser')
}

& $JPACKAGE @($jpkgArgs.ToArray())

if ($LASTEXITCODE -ne 0) {
    Write-Error 'jpackage failed.'
    exit $LASTEXITCODE
}

if ($usedStagingForJpackage) {
    $stagedApp = Join-Path $jpkgDestParent $APP_NAME
    if (-not (Test-Path -LiteralPath $stagedApp)) {
        throw "jpackage staging output missing: $stagedApp"
    }
    New-Item -ItemType Directory -Path $distFinal -Force | Out-Null
    $destAppPath = Join-Path $distFinal $APP_NAME
    $copyTarget = $destAppPath
    if (Test-Path -LiteralPath $destAppPath) {
        $removedOld = $false
        for ($ri = 0; $ri -lt 8; $ri++) {
            try {
                Remove-Item -Recurse -Force -LiteralPath $destAppPath -ErrorAction Stop
                $removedOld = $true
                break
            }
            catch {
                Write-Warning "Cannot remove old bundle folder (close $($APP_NAME).exe / Explorer on dist\$APP_NAME). Retry ($ri/8)..."
                Start-Sleep -Seconds 3
            }
        }
        if (-not $removedOld) {
            $copyTarget = Join-Path $distFinal ($APP_NAME + '-build-' + [Guid]::NewGuid().ToString('N').Substring(0, 12))
            Write-Warning "Old folder is locked; publishing fresh bundle to $copyTarget instead. Close handles on $destAppPath and rename/swap when possible."
        }
    }
    Copy-Item -Recurse -Force -LiteralPath $stagedApp -Destination $copyTarget
    Remove-Item -Recurse -Force -LiteralPath $jpkgDestParent -ErrorAction SilentlyContinue
    $publishedBundleRoot = (Resolve-Path -LiteralPath $copyTarget).Path
    Write-Host "Copied jpackage output from staging to: $publishedBundleRoot" -ForegroundColor DarkGray
}
else {
    $publishedBundleRoot = Join-Path $distFinal $APP_NAME
}

$dist = $distFinal

$postJpkgRoot = $publishedBundleRoot
if (Test-Path -LiteralPath $postJpkgRoot) {
    $diagBin = Join-Path $postJpkgRoot 'runtime\bin'
    $bundledJavaExe = Join-Path $diagBin 'java.exe'
    # LiteralPath can mis-detect on some Unicode/long paths; confirm with *.exe listing.
    $exeInBin = @()
    if (Test-Path -LiteralPath $diagBin) {
        $exeInBin = @(Get-ChildItem -LiteralPath $diagBin -Filter '*.exe' -File -ErrorAction SilentlyContinue)
    }
    $hasJavaExe = Test-Path -LiteralPath $bundledJavaExe
    if (-not $hasJavaExe -and $exeInBin.Count -gt 0) {
        $hasJavaExe = [bool]($exeInBin | Where-Object { $_.Name -ieq 'java.exe' })
    }

    if (-not $hasJavaExe) {
        Write-Warning @"
Bundled launcher not found: $bundledJavaExe
Common causes:
  1) Windows Defender / AV removed java.exe (DLLs often remain). Check Protection history.
  2) Very long or non-ASCII path - this script stages jpackage --dest under %TEMP% when the repo path has non-ASCII; override with -JpackageDest or PM_AI_JPACKAGE_DEST, or clone to e.g. C:\work\pm-ai.
  3) Stale dist - ensure code_java\dist was removed before jpackage.
Java runtime is next to $($APP_NAME).exe (from --runtime-image JDK); pm-ai-data\runtime is Python only.
Step 7 continues; output may be incomplete.
"@
        $diagRt = Join-Path $postJpkgRoot 'runtime'
        if (Test-Path -LiteralPath $diagBin) {
            Write-Host '--- Diagnostic: runtime\bin *.exe (all) ---' -ForegroundColor Yellow
            if ($exeInBin.Count -eq 0) {
                Write-Host '  (no .exe files - launchers missing or blocked)' -ForegroundColor Yellow
            }
            else {
                $exeInBin | Sort-Object Name | Format-Table Name, Length -AutoSize
            }
            Write-Host '--- Diagnostic: runtime\bin non-directory count by extension (sample) ---' -ForegroundColor DarkGray
            Get-ChildItem -LiteralPath $diagBin -File -ErrorAction SilentlyContinue |
                Group-Object Extension |
                Sort-Object Count -Descending |
                Select-Object -First 15 Name, Count |
                Format-Table -AutoSize
        }
        elseif (Test-Path -LiteralPath $diagRt) {
            Write-Host '--- Diagnostic: runtime exists but bin\ is missing ---' -ForegroundColor Yellow
            Get-ChildItem -LiteralPath $diagRt -ErrorAction SilentlyContinue | Select-Object Name
        }
        else {
            Write-Host '--- Diagnostic: runtime folder missing under app-image root ---' -ForegroundColor Yellow
        }
    }
}

Write-Host "--- Step 7: bundle pm-ai-data (Python + code/python + default dirs) ---" -ForegroundColor Cyan
$distRoot = $publishedBundleRoot
if (-not (Test-Path -LiteralPath $distRoot)) {
    throw "Distribution folder missing: $distRoot"
}
Copy-BundleToDist -WorkspaceRootPath $WorkspaceRoot -DistAppRoot $distRoot -PythonEmbedSourceDir $pythonSrc -AppExeBaseName $APP_NAME

$launcherBatDst = Join-Path $distRoot 'launch-pm-ai-desktop.bat'
$javafxVerForLauncher = Expand-PomPropertyPlaceholder -Raw ([string]$pomProps['javafx.version']) -Props $pomProps
if ([string]::IsNullOrWhiteSpace($javafxVerForLauncher)) {
    $javafxVerForLauncher = '26.0.1'
}
$batBody = Build-PmAiDesktopLauncherBatContent -JavafxVersion $javafxVerForLauncher -JvmInitial $jvmInitial -JvmMax $jvmMax -LauncherExeBaseName $APP_NAME
[System.IO.File]::WriteAllText($launcherBatDst, $batBody, [System.Text.UTF8Encoding]::new($false))
Write-Host "Launcher bat: $launcherBatDst (JavaFX module-path from pom javafx.version=$javafxVerForLauncher)" -ForegroundColor DarkGray

Write-Host "--- Done ---" -ForegroundColor Green
$mainExePath = Join-Path $distRoot ($APP_NAME + '.exe')
Write-Host "App: $mainExePath"
Write-Host "Portable data: $(Join-Path $distRoot 'pm-ai-data')"
Write-Host "JVM: -Xms$jvmInitial -Xmx$jvmMax (same as pom.xml properties)"
if ($PackageType -ne 'app-image') {
    Write-Host 'Check dist for installer output.'
}
if (-not $WinConsole -and $PackageType -eq 'app-image') {
    Write-Host 'Hint: console build: .\package_app.ps1 -WinConsole' -ForegroundColor DarkGray
}
