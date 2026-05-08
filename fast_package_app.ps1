# Production desktop (JavaFX) - Windows app bundle builder (jpackage + bundled runtime).
#
# Prerequisites:
#   - Run build on Windows (OpenJFX win classifier required).
#   - Maven uses JAVA_HOME on PATH (compile must match maven.compiler.release).
#   - Bundled JVM: Temurin JDK zip -> code_java/Cash_PMD -> jpackage --runtime-image. Override: -JdkRuntimeImage or PM_AI_JDK_RUNTIME_IMAGE.
#   - JavaFX: OpenJFX Windows win jars downloaded from Maven Central into package_input (same version as pom javafx.version).
#   - For --type exe/msi: WiX Toolset on PATH (candle/light).
#   - Bundled Python: pip runs at build time - internet on first run or empty cache.
#
# Usage (run from repository root):
#   .\fast_package_app.ps1
#   .\fast_package_app.ps1 -PackageType exe
#   .\fast_package_app.ps1 -RefreshCache        # Force re-download of all components (ignore cache)
#   .\fast_package_app.ps1 -WinConsole
#   .\fast_package_app.ps1 -JpackageDest C:\pm-ai-out   # ASCII-only parent for jpackage --dest (if launchers missing)
#   .\fast_package_app.ps1 -JdkRuntimeImage C:\path\to\jdk   # skip download; needs bin\java.exe and bin\jpackage.exe
# Env: PM_AI_JPACKAGE_DEST, PM_AI_JDK_RUNTIME_IMAGE (optional)

# UTF-8 BOM: Windows PowerShell 5.1 parses this file as UTF-8. Body is ASCII-only; Japanese paths live in package_app_mandatory_code_paths.txt.
[CmdletBinding()]
param(
    [ValidateSet('app-image', 'exe', 'msi')]
    [string]$PackageType = 'app-image',

    [switch]$WinConsole,

    # キャッシュを無視して強制的に再ダウンロードする場合に使用します
    [switch]$RefreshCache,

    # JDK root for --runtime-image (bin\java.exe). Empty = download per pom.xml into code_java/Cash_PMD.
    [string]$JdkRuntimeImage = '',

    # Parent directory for jpackage --dest only (must be ASCII-only on some JDK/Windows builds).
    [string]$JpackageDest = ''
)

$ErrorActionPreference = 'Stop'

# Script lives at repo root; Maven project is code_java/.
$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
$WorkspaceRoot = $ScriptRoot
$CodeJavaRoot = Join-Path $WorkspaceRoot 'code_java'
$ReleaseDirName = 'pm-ai-package-release'
$ReleaseRoot = Join-Path $WorkspaceRoot $ReleaseDirName
$BundleInitialName = 'PMD_initial_install'
$BundleUpgradeName = 'PMD_version_upgrade'

# Load shared copy logic before changing directory; Maven runs under code_java via Push-Location only.
$packageWorkspaceCopyScript = Join-Path $CodeJavaRoot 'package_workspace_copy.ps1'
if (-not (Test-Path -LiteralPath $packageWorkspaceCopyScript)) {
    throw "Missing package_workspace_copy.ps1: $packageWorkspaceCopyScript (pull latest or restore from repo)."
}
. (Resolve-Path -LiteralPath $packageWorkspaceCopyScript).Path

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
        [switch]$RefreshCache # [bool]から[switch]に変更
    )

    $dest = Join-Path $CacheRoot ('jdk-embed-' + $JdkRelease + '-windows-amd64')
    $javaExe = Join-Path $dest 'bin\java.exe'
    $jpkgExe = Join-Path $dest 'bin\jpackage.exe'

    # キャッシュが存在し、リフレッシュフラグがない場合は再利用
    if (-not $RefreshCache -and (Test-Path -LiteralPath $javaExe) -and (Test-Path -LiteralPath $jpkgExe)) {
        Write-Host "Using cached JDK: $dest" -ForegroundColor DarkGray
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
        [switch]$RefreshCache # [bool]から[switch]に変更
    )

    $dest = Join-Path $CacheRoot "python-embed-$PythonVersion-amd64"
    $pyExe = Join-Path $dest 'python.exe'
    $req = Join-Path $WorkspaceRootPath 'code\python\requirements.txt'

    # キャッシュが存在し、リフレッシュフラグがない場合は再利用
    if (-not $RefreshCache -and (Test-Path -LiteralPath $pyExe)) {
        Write-Host "Using cached Python embed: $dest" -ForegroundColor DarkGray
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
    $lines.Add('rem ASCII-only. Generated by fast_package_app.ps1 from pom javafx.version / jvm heap.')
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
        [switch]$RefreshCache # [bool]から[switch]に変更
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
        # キャッシュが存在し、リフレッシュフラグがない場合は再利用
        if (-not $RefreshCache -and (Test-Path -LiteralPath $cached)) {
            $fi = Get-Item -LiteralPath $cached -ErrorAction SilentlyContinue
            if ($null -ne $fi -and $fi.Length -gt 4096) {
                $needDownload = $false
                Write-Host "Using cached JavaFX runtime: $fn" -ForegroundColor DarkGray
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
        [ValidateSet('InitialInstall', 'VersionUpgrade')]
        [string]$BundleKind,
        [string]$MandatoryPathsFile,
        [string]$ReleaseFolderRelativePrefix,
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

    Write-Host "--- Copy workspace into pm-ai-data (bundle=$BundleKind) ---" -ForegroundColor Cyan
    Copy-WorkspaceTreeWithExplicitExclusions -RepoRoot $WorkspaceRootPath -DestRoot $data `
        -BundleKind $BundleKind -MandatoryPathsFile $MandatoryPathsFile -ReleaseFolderRelativePrefix $ReleaseFolderRelativePrefix

    $verifyPcInit = Join-Path $data 'code\python\planning_core\__init__.py'
    if (-not (Test-Path -LiteralPath $verifyPcInit)) {
        throw @"
Bundle incomplete: missing planning_core package.
Expected: $verifyPcInit
Ensure the repo workspace contains code/python/planning_core (clone depth / sparse checkout).
"@
    }

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
    $rmLines = [System.Collections.Generic.List[string]]::new()
    $rmLines.Add('JVM: bundled via jpackage --runtime-image from Temurin JDK (see pom.xml / fast_package_app.ps1).')
    $rmLines.Add('JavaFX: OpenJFX Windows jars pinned from Maven Central into package_input during fast_package_app.ps1 (javafx.version).')
    $rmLines.Add('Portable bundle generated by fast_package_app.ps1.')
    $rmLines.Add('Workspace mirror: package_workspace_copy.ps1 (shared with package_app.ps1).')
    $rmLines.Add('Master *.txt under code/ are always copied (see package_app_mandatory_code_paths.txt).')
    if ($BundleKind -eq 'InitialInstall') {
        $rmLines.Add('Bundle profile: InitialInstall - excludes .git, .venv, .cursor, .vscode, code/VBA, code_java build/cache dirs, pm-ai-package-release/, **/__pycache__, **/.pytest_cache, build_cache.')
        $rmLines.Add('Does NOT exclude plan/plans or code/output (may include local artifacts if present).')
    }
    else {
        $rmLines.Add('Bundle profile: VersionUpgrade - also excludes **/plan, **/plans, code/output/, repo output/, code/python/output/, .pm-ai-cache/, extra env-var TSVs (template TSV still bundled), .env.')
        $rmLines.Add('See package_workspace_copy.ps1 for exact rules.')
    }
    $rmLines.Add('Excluded files (all profiles): *.log, ~$* (Excel lock).')
    $rmLines.Add("This folder sits next to $($AppExeBaseName).exe.")
    $rmLines.Add('Release ZIPs omit pm-ai-data/version.txt; repo-root version.txt is shipped next to the ZIPs in pm-ai-package-release (interim bundle folders are deleted after zipping).')
    $rmLines.Add('First launch: if the empty marker file next to this app exe exists, the desktop resets env-tab defaults once then deletes it (Initial install bundle only). See Java AppPaths.PORTABLE_FIRST_LAUNCH_MARKER_FILE.')
    $rmLines.Add('Portable sync: PM_AI_PORTABLE_BUNDLE_SOURCE_DIR may be a folder (repo root layout under pm-ai-data on share) or a path to PMD_version_upgrade_*.zip with version.txt beside the zip.')
    $rmLines.Add('Python: pm-ai-data\runtime\python-embed\python.exe (build cache: code_java\Cash_PMD, not bundled).')
    $rmLines.Add('Default inputs: input\task-input , input\actual-detail.')
    $rmLines.Add('Per-user session data: ~/.pm-ai-desktop (initialized per machine/user).')
    $rmLines.Add('')
    $rmLines | Set-Content -LiteralPath $readme -Encoding UTF8
}

function Compress-PortableBundleFolderToZip {
    <#
    .SYNOPSIS
      Zip an app-image folder; omits pm-ai-data/version.txt (release version is beside the zip).
      Overwrites ZipFilePath when a file with that name already exists.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceDir,
        [Parameter(Mandatory = $true)]
        [string]$ZipFilePath
    )
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $sourceFull = (Resolve-Path -LiteralPath $SourceDir).Path.TrimEnd('\')
    if (Test-Path -LiteralPath $ZipFilePath) {
        Remove-Item -LiteralPath $ZipFilePath -Force -ErrorAction SilentlyContinue
    }
    $zipParent = Split-Path -Parent $ZipFilePath
    if (-not [string]::IsNullOrWhiteSpace($zipParent) -and -not (Test-Path -LiteralPath $zipParent)) {
        New-Item -ItemType Directory -Path $zipParent -Force | Out-Null
    }
    # Create (not CreateNew): always overwrite an existing ZIP with the same path.
    $fs = [System.IO.File]::Open($ZipFilePath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create, $false)
        try {
            Get-ChildItem -LiteralPath $sourceFull -Recurse -File -Force | ForEach-Object {
                $full = $_.FullName
                $rel = $full.Substring($sourceFull.Length).TrimStart('\')
                $relNorm = $rel -replace '\\', '/'
                if ($relNorm -ieq 'pm-ai-data/version.txt') {
                    return
                }
                $entryName = $rel -replace '\\', '/'
                if ($entryName -match '\.\./|^\.\.(/|\\)|(/|\\)\.\.(/|\\)') {
                    throw "Unsafe zip entry name: $entryName"
                }
                $entry = $zip.CreateEntry($entryName, [System.IO.Compression.CompressionLevel]::Optimal)
                $es = $entry.Open()
                try {
                    $srcFs = [System.IO.File]::OpenRead($full)
                    try {
                        $srcFs.CopyTo($es)
                    }
                    finally {
                        $srcFs.Dispose()
                    }
                }
                finally {
                    $es.Dispose()
                }
            }
        }
        finally {
            $zip.Dispose()
        }
    }
    finally {
        $fs.Dispose()
    }
}

$POM = Join-Path $CodeJavaRoot 'pom.xml'
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
$mvnw = Join-Path $CodeJavaRoot 'mvnw.cmd'
if (-not (Test-Path -LiteralPath $mvnw)) {
    throw "Maven Wrapper not found: $mvnw"
}
Push-Location $CodeJavaRoot
try {
    & $mvnw @('clean', 'package', '-DskipTests')
    if ($LASTEXITCODE -ne 0) {
        Write-Error 'Maven build failed.'
        exit $LASTEXITCODE
    }
}
finally {
    Pop-Location
}

Write-Host "--- Step 2: jpackage input directory ---" -ForegroundColor Cyan
$packageInput = Join-Path $CodeJavaRoot 'package_input'
Copy-JpackageInputDirectory -RootPath $CodeJavaRoot -MainJarName $proj.MainJar -DestPath $packageInput

New-Item -ItemType Directory -Path $ReleaseRoot -Force | Out-Null
# Download cache (JDK/JavaFX/Python embed) lives under code_java/Cash_PMD; pm-ai-package-release ends with ZIPs + version.txt only.
$cacheRoot = Join-Path $CodeJavaRoot 'Cash_PMD'
New-Item -ItemType Directory -Path $cacheRoot -Force | Out-Null
$legacyCashUnderRelease = Join-Path $ReleaseRoot 'Cash_PMD'
if (Test-Path -LiteralPath $legacyCashUnderRelease) {
    Write-Host "--- Remove legacy Cash_PMD under $ReleaseDirName (cache moved to code_java\Cash_PMD) ---" -ForegroundColor DarkGray
    Remove-Item -Recurse -Force -LiteralPath $legacyCashUnderRelease -ErrorAction SilentlyContinue
}

Write-Host "--- Step 3: JavaFX Windows runtime (Maven Central win jars -> package_input) ---" -ForegroundColor Cyan
$javafxVer = Expand-PomPropertyPlaceholder -Raw ([string]$pomProps['javafx.version']) -Props $pomProps
if ([string]::IsNullOrWhiteSpace($javafxVer)) {
    throw 'pom.xml: javafx.version is required for JavaFX bundle.'
}
Sync-JavaFxWindowsRuntimeFromMavenCentral -PackageInputDir $packageInput -JavafxVersion $javafxVer -CacheRoot $cacheRoot -RefreshCache:$RefreshCache

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
            -ZipUrlOverride $jdkZipUrlOverride -RefreshCache:$RefreshCache)
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
        -CacheRoot $cacheRoot -RefreshCache:$RefreshCache)

Write-Host "--- Step 6: jpackage (type=$PackageType) ---" -ForegroundColor Cyan

# Final output under repo-root pm-ai-package-release\. jpackage --dest may use TEMP when paths are non-ASCII.
$distFinal = $ReleaseRoot
$jpkgDestParent = $distFinal
if (-not [string]::IsNullOrWhiteSpace($JpackageDest)) {
    $jpkgDestParent = $JpackageDest.TrimEnd('\', '/')
}
elseif (-not [string]::IsNullOrWhiteSpace($env:PM_AI_JPACKAGE_DEST)) {
    $jpkgDestParent = $env:PM_AI_JPACKAGE_DEST.Trim().TrimEnd('\', '/')
}
elseif ($WorkspaceRoot -match '[^\x00-\x7F]') {
    $jpkgDestParent = Join-Path $env:TEMP ("pm-ai-jpackage-" + [Guid]::NewGuid().ToString('N'))
    Write-Host "Repo path contains non-ASCII: staging jpackage --dest to ASCII-only: $jpkgDestParent" -ForegroundColor Cyan
}
$usedStagingForJpackage = ($jpkgDestParent -ne $distFinal)

# Remove only prior jpackage app folder and bundle outputs (download cache is code_java\Cash_PMD, not under $ReleaseRoot).
$bundleOutInitial = Join-Path $ReleaseRoot $BundleInitialName
$bundleOutUpgrade = Join-Path $ReleaseRoot $BundleUpgradeName
$pathsToClean = @()
if ($usedStagingForJpackage) {
    $pathsToClean += $jpkgDestParent
}
else {
    $pathsToClean += (Join-Path $ReleaseRoot $APP_NAME)
}
$pathsToClean += $bundleOutInitial
$pathsToClean += $bundleOutUpgrade
$pathsToClean = $pathsToClean | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
foreach ($p in $pathsToClean) {
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

# Native exe launch uses only jpackage --java-options (see PMD_fast\<APP_NAME>\app\<APP_NAME>.cfg).
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
    # $APPDIR\ に設定 (appフォルダ自身の参照)
    [void]$modPathJpkgSb.Append('$APPDIR\')
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
                Write-Warning "Cannot remove old bundle folder (close $($APP_NAME).exe / Explorer on ${ReleaseDirName}\$APP_NAME). Retry ($ri/8)..."
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
  3) Stale bundle folder - remove ${ReleaseDirName}\$APP_NAME under repo root before jpackage.
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

Write-Host "--- Step 7: bundle pm-ai-data (Initial + VersionUpgrade) ---" -ForegroundColor Cyan
if (-not (Test-Path -LiteralPath $publishedBundleRoot)) {
    throw "Distribution folder missing: $publishedBundleRoot"
}

$mandatoryFile = Join-Path $CodeJavaRoot 'package_app_mandatory_code_paths.txt'
$relPref = "$ReleaseDirName/"

foreach ($destBundle in @($bundleOutInitial, $bundleOutUpgrade)) {
    if (Test-Path -LiteralPath $destBundle) {
        Remove-Item -Recurse -Force -LiteralPath $destBundle -ErrorAction Stop
    }
}

Write-Host "--- robocopy jpackage -> $BundleInitialName ---" -ForegroundColor Cyan
& robocopy $publishedBundleRoot $bundleOutInitial /E /NFL /NDL /NJH /NJS /nc /ns /np | Out-Host
if ($LASTEXITCODE -ge 8) {
    throw "robocopy to $BundleInitialName failed (exit $LASTEXITCODE)"
}
Write-Host "--- robocopy jpackage -> $BundleUpgradeName ---" -ForegroundColor Cyan
& robocopy $publishedBundleRoot $bundleOutUpgrade /E /NFL /NDL /NJH /NJS /nc /ns /np | Out-Host
if ($LASTEXITCODE -ge 8) {
    throw "robocopy to $BundleUpgradeName failed (exit $LASTEXITCODE)"
}

Write-Host "--- Remove intermediate folder $($APP_NAME) ---" -ForegroundColor DarkGray
Remove-Item -Recurse -Force -LiteralPath $publishedBundleRoot -ErrorAction SilentlyContinue

Copy-BundleToDist -WorkspaceRootPath $WorkspaceRoot -DistAppRoot $bundleOutInitial -PythonEmbedSourceDir $pythonSrc `
    -BundleKind InitialInstall -MandatoryPathsFile $mandatoryFile -ReleaseFolderRelativePrefix $relPref -AppExeBaseName $APP_NAME
Copy-BundleToDist -WorkspaceRootPath $WorkspaceRoot -DistAppRoot $bundleOutUpgrade -PythonEmbedSourceDir $pythonSrc `
    -BundleKind VersionUpgrade -MandatoryPathsFile $mandatoryFile -ReleaseFolderRelativePrefix $relPref -AppExeBaseName $APP_NAME

$javafxVerForLauncher = Expand-PomPropertyPlaceholder -Raw ([string]$pomProps['javafx.version']) -Props $pomProps
if ([string]::IsNullOrWhiteSpace($javafxVerForLauncher)) {
    $javafxVerForLauncher = '26.0.1'
}
$batBody = Build-PmAiDesktopLauncherBatContent -JavafxVersion $javafxVerForLauncher -JvmInitial $jvmInitial -JvmMax $jvmMax -LauncherExeBaseName $APP_NAME
foreach ($bd in @($bundleOutInitial, $bundleOutUpgrade)) {
    $launcherBatDst = Join-Path $bd 'launch-pm-ai-desktop.bat'
    [System.IO.File]::WriteAllText($launcherBatDst, $batBody, [System.Text.UTF8Encoding]::new($false))
    Write-Host "Launcher bat: $launcherBatDst" -ForegroundColor DarkGray
}

# Initial install only: empty marker for first-launch env reset (Java deletes after success).
# Build leaf name as UTF-16 code units to avoid script-file encoding mojibake on Windows PowerShell 5.1 (must match AppPaths.PORTABLE_FIRST_LAUNCH_MARKER_FILE).
$firstLaunchLeaf = (-join @([char]0x521d, [char]0x56de, [char]0x8d77, [char]0x52d5)) + '.txt'
$firstLaunchMarker = Join-Path $bundleOutInitial $firstLaunchLeaf
[System.IO.File]::WriteAllText($firstLaunchMarker, '', [System.Text.UTF8Encoding]::new($false))
Write-Host "First-launch marker (Initial only): $firstLaunchMarker" -ForegroundColor DarkGray

Write-Host "--- Step 8: release version.txt + portable ZIPs (overwrite same names; pm-ai-data/version.txt omitted inside ZIP) ---" -ForegroundColor Cyan
if (Test-Path -LiteralPath $VersionTxtPath) {
    $releaseVersionTxt = Join-Path $ReleaseRoot 'version.txt'
    Copy-Item -LiteralPath $VersionTxtPath -Destination $releaseVersionTxt -Force
    Write-Host "Copied/overwrote: $releaseVersionTxt" -ForegroundColor DarkGray
}
else {
    Write-Warning "Repo version.txt missing; skipped copy to $ReleaseRoot"
}

$zipVerSafe = $APP_VERSION -replace '[^\w\.\-]', '_'
$zipInitial = Join-Path $ReleaseRoot ("PMD_initial_install_$zipVerSafe.zip")
$zipUpgrade = Join-Path $ReleaseRoot ("PMD_version_upgrade_$zipVerSafe.zip")
Write-Host "Zipping Initial -> $zipInitial" -ForegroundColor Cyan
Compress-PortableBundleFolderToZip -SourceDir $bundleOutInitial -ZipFilePath $zipInitial
Write-Host "Zipping Upgrade -> $zipUpgrade" -ForegroundColor Cyan
Compress-PortableBundleFolderToZip -SourceDir $bundleOutUpgrade -ZipFilePath $zipUpgrade

if (-not (Test-Path -LiteralPath $zipInitial) -or -not (Test-Path -LiteralPath $zipUpgrade)) {
    throw "ZIP output missing after compress (initial or upgrade)."
}
Write-Host "--- Remove interim bundle folders (release artifacts: 2 ZIPs + version.txt only) ---" -ForegroundColor DarkGray
Remove-Item -Recurse -Force -LiteralPath $bundleOutInitial -ErrorAction Stop
Remove-Item -Recurse -Force -LiteralPath $bundleOutUpgrade -ErrorAction Stop

Write-Host "--- Done ---" -ForegroundColor Green
Write-Host "Release: $ReleaseRoot — PMD_initial_install_$zipVerSafe.zip, PMD_version_upgrade_$zipVerSafe.zip, version.txt"
Write-Host "Download cache: $cacheRoot"
Write-Host "JVM: -Xms$jvmInitial -Xmx$jvmMax (same as pom.xml properties)"
if ($PackageType -ne 'app-image') {
    Write-Host "Check $ReleaseDirName for installer output."
}
if (-not $WinConsole -and $PackageType -eq 'app-image') {
    Write-Host 'Hint: console build: .\fast_package_app.ps1 -WinConsole (run from repo root)' -ForegroundColor DarkGray
}