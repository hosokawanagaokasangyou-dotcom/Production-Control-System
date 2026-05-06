# 生産管理デスクトップ（JavaFX）? Windows 11 向け配布物作成（jpackage + 同梱ランタイム）
#
# 前提:
#   - ビルドは Windows 上で実行すること（OpenJFX の win 分類器が有効）。
#   - JDK は pom の maven.compiler.release と一致するフル JDK（jpackage 同梱）。現状は 26 系を想定。
#   - --type exe / msi には WiX Toolset（PATH に candle/light）が必要。
#   - 同梱 Python はビルド機で pip を実行するため、インターネット接続が必要（初回またはキャッシュ無し時）。
#
# 使い方:
#   .\package_app.ps1
#   .\package_app.ps1 -PackageType exe
#   .\package_app.ps1 -SkipPythonPrepare   # 既存 build_cache の Python をそのまま流用（高速）
#   .\package_app.ps1 -WinConsole

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
jpackage が見つかりません。
  JDK（pom の maven.compiler.release と一致）をインストールし、JAVA_HOME または PATH を設定してください。
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
            throw 'pom.xml から version を読み取れませんでした。'
        }
    }
    else {
        $version = $versionNode.InnerText.Trim()
    }
    if (-not $artifact -or -not $version) {
        throw 'pom.xml から artifactId / version を読み取れませんでした。'
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
        throw "メイン JAR が見つかりません: $mainSrc"
    }
    Copy-Item -LiteralPath $mainSrc -Destination $DestPath

    $depDir = Join-Path (Join-Path $RootPath 'target') 'dependency'
    if (-not (Test-Path -LiteralPath $depDir)) {
        throw "依存 JAR フォルダがありません: $depDir"
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
        Write-Host "SkipPythonPrepare: キャッシュをそのまま使います: $dest" -ForegroundColor DarkGray
        return $dest
    }

    if (-not (Test-Path -LiteralPath $req)) {
        throw "requirements.txt がありません: $req"
    }

    New-Item -ItemType Directory -Path $CacheRoot -Force | Out-Null
    if (Test-Path -LiteralPath $dest) {
        Remove-Item -Recurse -Force $dest
    }
    New-Item -ItemType Directory -Path $dest | Out-Null

    $zipUrl = "https://www.python.org/ftp/python/$PythonVersion/python-$PythonVersion-embed-amd64.zip"
    $zipPath = Join-Path $dest 'python-embed.zip'
    Write-Host "--- Python embed 取得: $zipUrl ---" -ForegroundColor Cyan
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
    Write-Host "--- get-pip 取得 ---" -ForegroundColor Cyan
    Invoke-WebRequest -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile $getPip -UseBasicParsing

    Push-Location $dest
    try {
        & .\python.exe $getPip
        if ($LASTEXITCODE -ne 0) {
            throw 'get-pip が失敗しました。'
        }
        & .\python.exe -m pip install --upgrade pip
        if ($LASTEXITCODE -ne 0) {
            throw 'pip のアップグレードに失敗しました。'
        }
        & .\python.exe -m pip install -r $req
        if ($LASTEXITCODE -ne 0) {
            throw 'pip install -r requirements.txt が失敗しました。'
        }
    }
    finally {
        Pop-Location
    }

    return $dest
}

function Copy-BundleToDist {
    param(
        [string]$WorkspaceRootPath,
        [string]$DistAppRoot,
        [string]$PythonEmbedSourceDir
    )

    $data = Join-Path $DistAppRoot 'pm-ai-data'
    if (Test-Path -LiteralPath $data) {
        Remove-Item -Recurse -Force $data
    }

    $codePySrc = Join-Path $WorkspaceRootPath 'code\python'
    if (-not (Test-Path -LiteralPath $codePySrc)) {
        throw "code\python が見つかりません: $codePySrc"
    }

    New-Item -ItemType Directory -Path $data -Force | Out-Null
    $destPy = Join-Path $data 'code\python'
    New-Item -ItemType Directory -Path (Split-Path $destPy -Parent) -Force | Out-Null

    Write-Host "--- pm-ai-data に code\python を複製 ---" -ForegroundColor Cyan
    & robocopy $codePySrc $destPy /E /NFL /NDL /NJH /NJS /nc /ns /np | Out-Host
    $rc = $LASTEXITCODE
    if ($rc -ge 8) {
        throw "robocopy が失敗しました (exit $rc)"
    }

    foreach ($name in @('特別ルール.md', '特別ルール列挙.md')) {
        $srcMd = Join-Path $WorkspaceRootPath $name
        if (Test-Path -LiteralPath $srcMd) {
            Copy-Item -LiteralPath $srcMd -Destination $data -Force
        }
    }

    New-Item -ItemType Directory -Path (Join-Path $data 'input\task-input') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $data 'input\actual-detail') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $data 'output') -Force | Out-Null

    $rt = Join-Path $data 'runtime\python-embed'
    New-Item -ItemType Directory -Path $rt -Force | Out-Null
    Write-Host "--- Python ランタイムを pm-ai-data に複製 ---" -ForegroundColor Cyan
    & robocopy $PythonEmbedSourceDir $rt /E /NFL /NDL /NJH /NJS /nc /ns /np | Out-Host
    $rc2 = $LASTEXITCODE
    if ($rc2 -ge 8) {
        throw "robocopy python-embed が失敗しました (exit $rc2)"
    }

    $readme = Join-Path $data 'README_PORTABLE.txt'
    @(
        'このフォルダは package_app.ps1 が生成したポータブル用データです。',
        'PmAiDesktop.exe と同じ階層にあります。',
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
    throw 'pom.properties に pm.ai.bundle.python.embed.version がありません。'
}

$JPACKAGE = Resolve-JpackageExe
$proj = Get-MavenProjectInfo -PomPath $POM

$APP_NAME = 'PmAiDesktop'
$APP_VERSION = $proj.Version -replace '-SNAPSHOT$', '.0'

Write-Host "--- Step 1: Maven package ---" -ForegroundColor Cyan
$mvnw = Join-Path $Root 'mvnw.cmd'
if (-not (Test-Path -LiteralPath $mvnw)) {
    throw "Maven Wrapper が見つかりません: $mvnw"
}
& $mvnw @('clean', 'package', '-DskipTests')
if ($LASTEXITCODE -ne 0) {
    Write-Error 'Maven のビルドに失敗しました。'
    exit $LASTEXITCODE
}

Write-Host "--- Step 2: jpackage 入力ディレクトリ ---" -ForegroundColor Cyan
$packageInput = Join-Path $Root 'package_input'
Copy-JpackageInputDirectory -RootPath $Root -MainJarName $proj.MainJar -DestPath $packageInput

Write-Host "--- Step 3: Python 同梱キャッシュ（pip）---" -ForegroundColor Cyan
$cacheRoot = Join-Path $Root 'build_cache'
$pythonSrc = Ensure-PythonEmbedCache -WorkspaceRootPath $WorkspaceRoot -PythonVersion $pyEmbedVer `
    -CacheRoot $cacheRoot -Skip:$SkipPythonPrepare

Write-Host "--- Step 4: jpackage（type=$PackageType）---" -ForegroundColor Cyan
$dist = Join-Path $Root 'dist'
if (Test-Path $dist) {
    Remove-Item -Recurse -Force $dist
}

$javaOpts = @(
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
    Write-Error 'jpackage が失敗しました。'
    exit $LASTEXITCODE
}

Write-Host "--- Step 5: pm-ai-data 同梱（Python + code/python + 既定フォルダ）---" -ForegroundColor Cyan
$distRoot = Join-Path $dist $APP_NAME
if (-not (Test-Path -LiteralPath $distRoot)) {
    throw "配布フォルダがありません: $distRoot"
}
Copy-BundleToDist -WorkspaceRootPath $WorkspaceRoot -DistAppRoot $distRoot -PythonEmbedSourceDir $pythonSrc

Write-Host "--- 完了 ---" -ForegroundColor Green
Write-Host "アプリ本体: $(Join-Path $distRoot "$APP_NAME.exe")"
Write-Host "ポータブルデータ: $(Join-Path $distRoot 'pm-ai-data')"
Write-Host "JVM: -Xms$jvmInitial -Xmx$jvmMax （pom.xml properties と同一）"
if ($PackageType -ne 'app-image') {
    Write-Host "インストーラーは dist 直下を確認してください。"
}
if (-not $WinConsole -and $PackageType -eq 'app-image') {
    Write-Host "ヒント: コンソール付きは .\package_app.ps1 -WinConsole" -ForegroundColor DarkGray
}
