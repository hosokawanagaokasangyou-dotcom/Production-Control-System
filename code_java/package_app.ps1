# 生産管理デスクトップ（JavaFX）? Windows 11 向け配布物作成（jpackage）
#
# 前提:
#   ? ビルドは Windows 上で実行すること（OpenJFX の win 分類器が有効になる）。
#   ? JDK は pom の maven.compiler.release と一致するフル JDK（jpackage 同梱）。
#     本プロジェクトは release 26 を想定（Temurin 等の JDK 26+）。
#   ? --type exe / msi を使う場合は WiX Toolset（PATH に candle/light）が必要。
#
# 使い方:
#   .\package_app.ps1                          # アプリフォルダ（app-image）? 別 PC へフォルダごとコピー可
#   .\package_app.ps1 -PackageType exe         # 単体インストーラー（.exe）
#   .\package_app.ps1 -PackageType msi         # MSI インストーラー
#   .\package_app.ps1 -WinConsole              # コンソール付き exe（ログ確認用）

param(
    [ValidateSet('app-image', 'exe', 'msi')]
    [string]$PackageType = 'app-image',

    [switch]$WinConsole
)

$ErrorActionPreference = 'Stop'

$Root = if ($PSScriptRoot) { $PSScriptRoot } else { Get-Location }
Set-Location $Root

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
  ? pom の maven.compiler.release と一致するフル JDK をインストールし、JAVA_HOME を設定するか PATH に含めてください。
  ? Windows 11 では Eclipse Temurin / Microsoft Build of OpenJDK などの JDK 配布でも jpackage は bin に含まれます。
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
        throw "メイン JAR が見つかりません: $mainSrc （mvn package が成功しているか確認してください）"
    }
    Copy-Item -LiteralPath $mainSrc -Destination $DestPath

    $depDir = Join-Path (Join-Path $RootPath 'target') 'dependency'
    if (-not (Test-Path -LiteralPath $depDir)) {
        throw "依存 JAR フォルダがありません: $depDir （pom の maven-dependency-plugin が有効か、mvn package を再実行してください）"
    }
    Copy-Item -Path (Join-Path $depDir '*') -Destination $DestPath -Force
}

$JPACKAGE = Resolve-JpackageExe
$POM = Join-Path $Root 'pom.xml'
$proj = Get-MavenProjectInfo -PomPath $POM

$APP_NAME = 'PmAiDesktop'
# jpackage の --app-version は数値ドット区切りのみ（0.1.0-SNAPSHOT は不可）
$APP_VERSION = $proj.Version -replace '-SNAPSHOT$', '.0'

Write-Host "--- Step 1: Maven package（依存 JAR を target/dependency に複製）---" -ForegroundColor Cyan
$mvnw = Join-Path $Root 'mvnw.cmd'
if (-not (Test-Path -LiteralPath $mvnw)) {
    throw "Maven Wrapper が見つかりません: $mvnw"
}
& $mvnw @('clean', 'package', '-DskipTests')
if ($LASTEXITCODE -ne 0) {
    Write-Error 'Maven のビルドに失敗しました。'
    exit $LASTEXITCODE
}

Write-Host "--- Step 2: jpackage 用入力ディレクトリを準備 ---" -ForegroundColor Cyan
$packageInput = Join-Path $Root 'package_input'
Copy-JpackageInputDirectory -RootPath $Root -MainJarName $proj.MainJar -DestPath $packageInput

Write-Host "--- Step 3: jpackage（type=$PackageType）---" -ForegroundColor Cyan
$dist = Join-Path $Root 'dist'
if (Test-Path $dist) {
    Remove-Item -Recurse -Force $dist
}

# javafx-maven-plugin と同等の実行時オプション（PmAiFxApp / Prism / ControlsFX）
$javaOpts = @(
    '-Xms512m',
    '-Xmx3g',
    '-XX:+HeapDumpOnOutOfMemoryError',
    '-XX:+UseStringDeduplication',
    '-Dprism.order=sw',
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

Write-Host "--- 完了 ---" -ForegroundColor Green
if ($PackageType -eq 'app-image') {
    $distRoot = Join-Path $dist $APP_NAME
    Write-Host "出力フォルダ: $distRoot"
    Write-Host "実行ファイル: $(Join-Path $distRoot "$APP_NAME.exe")"
    Write-Host "別 PC では上記フォルダ一式をコピーし、exe を実行してください（Python / マスタ等は別途配置・環境変数で指定）。"
}
else {
    Write-Host "インストーラー出力先: $dist （.exe / .msi を確認してください）"
}

if (-not $WinConsole -and $PackageType -eq 'app-image') {
    Write-Host "ヒント: コンソールが必要な場合は .\package_app.ps1 -WinConsole を実行してください。" -ForegroundColor DarkGray
}
