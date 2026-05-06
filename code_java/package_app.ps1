# Production desktop (JavaFX) — Windows app bundle builder (jpackage + bundled runtime).
#
# Prerequisites:
#   - Run build on Windows (OpenJFX win classifier required).
#   - Full JDK matching pom maven.compiler.release (jpackage). Currently JDK 26 expected.
#   - For --type exe/msi: WiX Toolset on PATH (candle/light).
#   - Bundled Python: pip runs at build time — internet on first run or empty cache.
#
# Usage:
#   .\package_app.ps1
#   .\package_app.ps1 -PackageType exe
#   .\package_app.ps1 -SkipPythonPrepare   # reuse existing build_cache python (faster)
#   .\package_app.ps1 -WinConsole

# UTF-8 BOM helps Windows PowerShell 5.1 display Japanese in README / paths consistently.
[CmdletBinding()]
param(
    [ValidateSet('app-image', 'exe', 'msi')]
    [string]$PackageType = 'app-image',

    [switch]$WinConsole,

    [switch]$SkipPythonPrepare
)

$ErrorActionPreference = 'Stop'

$Root = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
Set-Location $Root

$WorkspaceRoot = (Resolve-Path -LiteralPath (Join-Path $Root '..')).Path

function Copy-WorkspaceTreeRespectingGitIgnore {
    param(
        [string]$RepoRoot,
        [string]$DestRoot
    )

    $gitMarker = Join-Path $RepoRoot '.git'
    if (-not (Test-Path -LiteralPath $gitMarker)) {
        throw @'
Repository root has no .git.
package_app.ps1 bundles pm-ai-data using git ls-files and .gitignore. Run from a Git checkout.
'@
    }

    if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
        throw @'
git is not on PATH.
Install Git for Windows so git ls-files can honor .gitignore.
'@
    }

    Push-Location $RepoRoot
    try {
        $stdout = & git -c core.quotepath=false ls-files -co --exclude-standard 2>$null
        if ($LASTEXITCODE -ne 0) {
            throw "git ls-files failed (exit $LASTEXITCODE). Does git status work in this folder?"
        }
    }
    finally {
        Pop-Location
    }

    $lines = @()
    if ($null -eq $stdout) {
        $lines = @()
    }
    elseif ($stdout -is [System.Array]) {
        $lines = $stdout
    }
    else {
        $lines = @($stdout.ToString() -split "`r?`n")
    }

    # Skip packaging scratch dirs even if listed (avoid duplicate Python / bloat).
    function Test-IsPackagingScratchPath {
        param([string]$RelSlash)
        foreach ($prefix in @(
                'code_java/build_cache/',
                'code_java/package_input/',
                'code_java/dist/'
            )) {
            if ($RelSlash.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
                return $true
            }
        }
        return $false
    }

    foreach ($raw in $lines) {
        if ($null -eq $raw) {
            continue
        }
        $relSlash = ($raw.ToString().Trim() -replace '\\', '/')
        if ([string]::IsNullOrWhiteSpace($relSlash)) {
            continue
        }

        if (Test-IsPackagingScratchPath -RelSlash $relSlash) {
            continue
        }

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

    # Always copy master *.txt under code/ (repo-relative UTF-8 paths).
    $mandatoryCodeRootTxt = @(
        "code/使用原反, 加工幅.txt",
        "code/使用原反,ロール単位の長さ.txt",
        "code/製品名,ロール単位の長さ.txt",
        "code/製品名,製品厚み.txt",
        "code/製品名,製品長.txt",
        "code/製品名, 製品幅.txt"
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

function Add-JpackageCandidatesFromDir {
    param(
        [string]$BasePath,
        [string]$NamePattern
    )
    $list = [System.Collections.Generic.List[string]]::new()
    if (-not (Test-Path -LiteralPath $BasePath)) {
        return $list
    }
    Get-ChildItem -LiteralPath $BasePath -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match $NamePattern } |
        Sort-Object Name -Descending |
        ForEach-Object {
            $exe = Join-Path $_.FullName 'bin\jpackage.exe'
            $list.Add($exe)
        }
    return $list
}

function Resolve-JpackageExe {
    $candidates = [System.Collections.Generic.List[string]]::new()

    if ($env:JAVA_HOME) {
        $candidates.Add((Join-Path $env:JAVA_HOME 'bin\jpackage.exe'))
    }

    foreach ($pair in @(
            @{ Base = 'C:\Program Files\Eclipse Adoptium'; Pattern = '^jdk-(2[1-9]|[34][0-9])' },
            @{ Base = 'C:\Program Files\Microsoft';        Pattern = '^jdk-(2[1-9]|[34][0-9])' },
            @{ Base = 'C:\Program Files\Java';             Pattern = '^jdk-(2[1-9]|[34][0-9])' }
        )) {
        foreach ($p in (Add-JpackageCandidatesFromDir -BasePath $pair.Base -NamePattern $pair.Pattern)) {
            $candidates.Add($p)
        }
    }

    @(
        'C:\Program Files\Java\jdk-26\bin\jpackage.exe',
        'C:\Program Files\Java\jdk-25\bin\jpackage.exe',
        'C:\Program Files\Java\jdk-21\bin\jpackage.exe'
    ) | ForEach-Object { $candidates.Add($_) }

    foreach ($p in $candidates) {
        if ($p -and (Test-Path -LiteralPath $p)) {
            return (Resolve-Path -LiteralPath $p).Path
        }
    }

    $cmd = Get-Command jpackage -ErrorAction SilentlyContinue
    if ($cmd) {
        return $cmd.Source
    }

    throw @'
jpackage not found.
Install a JDK matching pom maven.compiler.release and set JAVA_HOME or PATH.
'@
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

function Copy-BundleToDist {
    param(
        [string]$WorkspaceRootPath,
        [string]$DistAppRoot,
        [string]$PythonEmbedSourceDir
    )

    if ([string]::IsNullOrWhiteSpace($PythonEmbedSourceDir) -or -not (Test-Path -LiteralPath $PythonEmbedSourceDir)) {
        throw "Invalid Python embed path: '$PythonEmbedSourceDir'"
    }

    $data = Join-Path $DistAppRoot 'pm-ai-data'
    if (Test-Path -LiteralPath $data) {
        Remove-Item -Recurse -Force $data
    }

    New-Item -ItemType Directory -Path $data -Force | Out-Null

    Write-Host "--- Copy workspace into pm-ai-data (git ls-files / .gitignore + mandatory code/*.txt) ---" -ForegroundColor Cyan
    Copy-WorkspaceTreeRespectingGitIgnore -RepoRoot $WorkspaceRootPath -DestRoot $data

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
        'このフォルダは package_app.ps1 が生成したポータブル用データです。',
        'ワークスペースの複製元: git ls-files -co --exclude-standard（.gitignore で無視されるものは含みません）。',
        'code 直下のマスタ .txt（使用原反／製品名・ロール単位などの一覧）は package_app.ps1 で必ず同梱します。',
        '除外追加: code_java/build_cache, package_input, dist（パッケージ作業用）。',
        'PmAiDesktop.exe と同じ階層にあります。',
        '版はリポジトリ直下 version.txt（この複製に含まれる）。正本フォルダは環境変数 PM_AI_PORTABLE_BUNDLE_SOURCE_DIR で指定して起動時同期できます。',
        'Python: runtime\python-embed\python.exe（requirements 済みキャッシュを複製）',
        '入力フォルダの既定: input\task-input , input\actual-detail（アプリ起動時に参照されます）',
        'セッション ~/.pm-ai-desktop はユーザーごとに別 PC で初期化されます。',
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

$JPACKAGE = Resolve-JpackageExe
$proj = Get-MavenProjectInfo -PomPath $POM

$APP_NAME = 'PmAiDesktop'
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

Write-Host "--- Step 3: Python embed cache (pip) ---" -ForegroundColor Cyan
$cacheRoot = Join-Path $Root 'build_cache'
$pythonSrc = [string](Ensure-PythonEmbedCache -WorkspaceRootPath $WorkspaceRoot -PythonVersion $pyEmbedVer `
        -CacheRoot $cacheRoot -Skip:$SkipPythonPrepare)

Write-Host "--- Step 4: jpackage (type=$PackageType) ---" -ForegroundColor Cyan
$dist = Join-Path $Root 'dist'
if (Test-Path $dist) {
    $removed = $false
    for ($i = 0; $i -lt 5; $i++) {
        try {
            Remove-Item -Recurse -Force -LiteralPath $dist -ErrorAction Stop
            $removed = $true
            break
        }
        catch {
            Write-Warning "Cannot remove dist (file may be locked). Close PmAiDesktop / Explorer using dist, retry ($i/5)..."
            Start-Sleep -Seconds 2
        }
    }
    if (-not $removed -and (Test-Path $dist)) {
        throw @'
Cannot remove dist. Close any handle on code_java\dist (explorer, running app), then re-run package_app.ps1.
'@
    }
}

$javaOpts = @(
    '-Dfile.encoding=UTF-8',
    "-Xms$jvmInitial",
    "-Xmx$jvmMax",
    '-XX:+HeapDumpOnOutOfMemoryError',
    '-XX:+UseStringDeduplication',
    "-Dprism.order=$prismOrder",
    '--add-opens=javafx.base/com.sun.javafx.event=ALL-UNNAMED',
    '--add-opens=javafx.controls/javafx.scene.control.skin=ALL-UNNAMED',
    '--add-exports=javafx.controls/com.sun.javafx.scene.control.behavior=ALL-UNNAMED'
)

$jpkgArgs = [System.Collections.Generic.List[string]]::new()
$jpkgArgs.Add('--type')
$jpkgArgs.Add($PackageType)
$jpkgArgs.Add('--input')
$jpkgArgs.Add($packageInput)
$jpkgArgs.Add('--dest')
$jpkgArgs.Add($dist)
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

$postJpkgRoot = Join-Path $dist $APP_NAME
if (Test-Path -LiteralPath $postJpkgRoot) {
    $bundledJavaExe = Join-Path $postJpkgRoot 'runtime\bin\java.exe'
    if (-not (Test-Path -LiteralPath $bundledJavaExe)) {
        Write-Warning @"
Bundled JRE not found: $bundledJavaExe
Common causes:
  1) Windows Defender / AV removed java.exe (often DLLs remain). Check Protection history.
  2) Very long or non-ASCII path (e.g. Japanese folder names) — try cloning to a short path like C:\work\pm-ai and run package_app.ps1 there.
  3) Stale dist — ensure Step 4 deleted code_java\dist before jpackage.
App runtime is next to PmAiDesktop.exe; pm-ai-data\runtime is Python only.
Step 5 continues; output may be incomplete.
"@
        $diagBin = Join-Path $postJpkgRoot 'runtime\bin'
        $diagRt = Join-Path $postJpkgRoot 'runtime'
        if (Test-Path -LiteralPath $diagBin) {
            Write-Host '--- Diagnostic: runtime\bin (first 40 entries) ---' -ForegroundColor Yellow
            Get-ChildItem -LiteralPath $diagBin -ErrorAction SilentlyContinue |
                Select-Object -First 40 Name, Length |
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

Write-Host "--- Step 5: bundle pm-ai-data (Python + code/python + default dirs) ---" -ForegroundColor Cyan
$distRoot = Join-Path $dist $APP_NAME
if (-not (Test-Path -LiteralPath $distRoot)) {
    throw "Distribution folder missing: $distRoot"
}
Copy-BundleToDist -WorkspaceRootPath $WorkspaceRoot -DistAppRoot $distRoot -PythonEmbedSourceDir $pythonSrc

$launcherBatSrc = Join-Path $Root 'launch-pm-ai-desktop-portable.bat'
$launcherBatDst = Join-Path $distRoot 'launch-pm-ai-desktop.bat'
if (Test-Path -LiteralPath $launcherBatSrc) {
    Copy-Item -LiteralPath $launcherBatSrc -Destination $launcherBatDst -Force
    Write-Host "Launcher bat: $launcherBatDst (alternative to exe)" -ForegroundColor DarkGray
}

Write-Host "--- Done ---" -ForegroundColor Green
Write-Host "App: $(Join-Path $distRoot "$APP_NAME.exe")"
Write-Host "Portable data: $(Join-Path $distRoot 'pm-ai-data')"
Write-Host "JVM: -Xms$jvmInitial -Xmx$jvmMax (same as pom.xml properties)"
if ($PackageType -ne 'app-image') {
    Write-Host 'Check dist for installer output.'
}
if (-not $WinConsole -and $PackageType -eq 'app-image') {
    Write-Host 'Hint: console build: .\package_app.ps1 -WinConsole' -ForegroundColor DarkGray
}
